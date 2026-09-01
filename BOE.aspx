<%@ Import Namespace="System.Xml" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Page Language="VB" Title="Export BOE Result" ContentType="text/xml" Debug="true" %>

<script runat="server">

    Dim processClass As New ProcessClass
    Dim myConn As String = ConfigurationManager.ConnectionStrings("DefaultConnection").ConnectionString

    Public Function GetDataTable(thisString As String) As DataTable
        Try
            Using thisConn As New SqlConnection(myConn)
                Using thisCmd As New SqlCommand(thisString, thisConn)
                    Using thisAdapter As New SqlDataAdapter(thisCmd)
                        Dim thisTable As New DataTable()
                        thisAdapter.Fill(thisTable)
                        Return thisTable
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Function GetItemData(thisString As String) As String
        Dim result As String = String.Empty
        Using thisConn As New SqlConnection(myConn)
            thisConn.Open()
            Using thisCmd As New SqlCommand(thisString, thisConn)
                Using rdResult = thisCmd.ExecuteReader
                    While rdResult.Read
                        result = rdResult.Item(0).ToString()
                    End While
                End Using
            End Using
            thisConn.Close()
        End Using
        Return result
    End Function

    Protected Sub Page_Load(sender As Object, e As EventArgs)
        Dim type As String = Request.QueryString("type").ToString()
        Dim company As String = Request.QueryString("company").ToString()
        Dim companyId As String = company
        Dim action As String = String.Empty
        If Not String.IsNullOrEmpty(Request.QueryString("action")) Then
            action = Request.QueryString("action").ToString()
        End If

        Dim status As String = String.Empty
        If Not String.IsNullOrEmpty(Request.QueryString("status")) Then
            status = Request.QueryString("status").ToString()
        End If

        Dim stringCompany As String = String.Empty
        stringCompany = "AND Customers.CompanyId='" & company & "'"
        If company = "jpmd" Then
            stringCompany = "AND Customers.CompanyId='2'" : companyId = "2"
        End If
        If company = "local" Then
            stringCompany = "AND Customers.CompanyId='3'" : companyId = "3"
        End If

        If action = "check" Then
            Dim stringStatus As String = "AND OrderHeaders.Status='" & status & "'"
            If type = "header" Then
                Using thisConn As New SqlConnection(myConn)
                    Using thisCmd As New SqlCommand(String.Format("SELECT OrderHeaders.*, Customers.Name AS CustomerName, Customers.DebtorCode AS DebtorCode, Logins.UserName AS UserName FROM OrderHeaders INNER JOIN Customers ON OrderHeaders.CustomerId=Customers.Id INNER JOIN Logins ON OrderHeaders.CreatedBy=Logins.Id WHERE OrderHeaders.Active=1 {0} {1} ORDER BY OrderHeaders.Id DESC", stringStatus, stringCompany), thisConn)
                        thisConn.Open()
                        Using reader As SqlDataReader = thisCmd.ExecuteReader()
                            DataHeader(reader)
                        End Using
                    End Using
                End Using
            End If

            If type = "detail" Then
                Using thisConn As New SqlConnection(myConn)
                    Using thisCmd As New SqlCommand(String.Format("SELECT OrderDetails.*, Products.DesignId AS DesignId, Products.BlindId AS BlindId, Designs.Name AS DesignName, Blinds.Name AS BlindName, Products.Name AS ProductName, ProductTubes.Name AS TubeName, ProductControls.Name AS ControlName, ProductColours.Name AS ColourName FROM OrderDetails INNER JOIN OrderHeaders ON OrderDetails.HeaderId=OrderHeaders.Id INNER JOIN Customers ON OrderHeaders.CustomerId=Customers.Id LEFT JOIN Products ON OrderDetails.ProductId=Products.Id LEFT JOIN Designs ON Products.DesignId=Designs.Id LEFT JOIN Blinds ON Products.BlindId=Blinds.Id LEFT JOIN ProductTubes ON Products.TubeType=ProductTubes.Id LEFT JOIN ProductControls ON Products.ControlType=ProductControls.Id LEFT JOIN ProductColours ON Products.ColourType=ProductColours.Id WHERE OrderDetails.Active=1 AND OrderHeaders.Active=1 {0} {1} ORDER BY OrderDetails.Id ASC", stringStatus, stringCompany), thisConn)
                        thisConn.Open()
                        Using reader As SqlDataReader = thisCmd.ExecuteReader()
                            DataDetail(reader)
                        End Using
                    End Using
                End Using
            End If
        End If

        If action = "download" Then
            Using thisConn As New SqlConnection(myConn)
                Using thisCmd As New SqlCommand("sp_OrderHeaders_Update_Production", thisConn)
                    thisCmd.CommandType = CommandType.StoredProcedure
                    thisCmd.Parameters.Add("@ActionBy", SqlDbType.Int).Value = 2
                    thisConn.Open()
                    thisCmd.ExecuteNonQuery()
                End Using
            End Using

            processClass.RefreshSalesData(companyId)

            If type = "header" Then
                Using thisConn As New SqlConnection(myConn)
                    Using thisCmd As New SqlCommand("sp_OrderHeaders_Download", thisConn)
                        thisCmd.CommandType = CommandType.StoredProcedure
                        thisCmd.Parameters.AddWithValue("@Company", company)

                        thisConn.Open()
                        Using reader As SqlDataReader = thisCmd.ExecuteReader()
                            DataHeader(reader)
                        End Using
                    End Using
                End Using
            End If

            If type = "detail" Then
                Using thisConn As New SqlConnection(myConn)
                    Using thisCmd As New SqlCommand("sp_OrderDetails_Download", thisConn)
                        thisCmd.CommandType = CommandType.StoredProcedure
                        thisCmd.Parameters.AddWithValue("@Company", company)
                        thisConn.Open()
                        Using reader As SqlDataReader = thisCmd.ExecuteReader()
                            DataDetail(reader)
                        End Using
                    End Using
                End Using
            End If

            DownloadDate()
        End If
    End Sub

    Protected Sub DownloadDate()
        Try
            Using thisConn As New SqlConnection(myConn)
                Using thisCmd As New SqlCommand("UPDATE OrderHeaders SET DownloadDate=GETDATE() WHERE Download='Yes' AND DownloadDate IS NULL AND Active=1", thisConn)
                    thisConn.Open()
                    thisCmd.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As Exception
        End Try
    End Sub

    Protected Sub DataHeader(myData As SqlDataReader)
        Dim writer As New XmlTextWriter(Response.OutputStream, Encoding.ASCII)
        writer.WriteStartDocument()
        writer.WriteStartElement("Orders")

        While myData.Read()
            writer.WriteStartElement("OrderHeader")
            writer.WriteAttributeString("Id", myData("Id").ToString())
            writer.WriteAttributeString("OrderId", myData("OrderId").ToString())
            writer.WriteAttributeString("CustomerId", myData("CustomerId").ToString())
            writer.WriteAttributeString("OrderNumber", myData("OrderNumber").ToString())
            writer.WriteAttributeString("OrderName", myData("OrderName").ToString())
            writer.WriteAttributeString("OrderNote", myData("OrderNote").ToString())
            writer.WriteAttributeString("OrderState", myData("OrderState").ToString())
            writer.WriteAttributeString("OrderAddress", myData("OrderAddress").ToString())
            writer.WriteAttributeString("OrdID", myData("Id").ToString())
            writer.WriteAttributeString("StoreOrderNo", myData("OrderNumber").ToString())
            writer.WriteAttributeString("StoreCustomer", myData("OrderName").ToString())
            writer.WriteAttributeString("DebtorCode", myData("DebtorCode").ToString())
            writer.WriteAttributeString("Zone", myData("OrderState").ToString())
            writer.WriteAttributeString("Address", myData("OrderAddress").ToString())

            Dim prodDate As DateTime
            If DateTime.TryParse(myData("ProductionDate").ToString(), prodDate) Then
                writer.WriteAttributeString("SubmittedDate", prodDate.ToString("dd/MM/yyyy"))
            Else
                writer.WriteAttributeString("SubmittedDate", "")
            End If
            writer.WriteAttributeString("Production", "BIG")
            writer.WriteAttributeString("SubProduction", "Universal")
            writer.WriteAttributeString("Status", myData("Status").ToString())
            writer.WriteEndElement()
        End While

        writer.WriteEndElement()
        writer.WriteEndDocument()
        writer.Close()
    End Sub

    Protected Sub DataDetail(thisData As SqlDataReader)
        Dim writer As New XmlTextWriter(Response.OutputStream, Encoding.ASCII)
        writer.WriteStartDocument()
        writer.WriteStartElement("Orders")

        While thisData.Read()
            Dim productId As String = thisData("ProductId").ToString()
            Dim productName As String = thisData("ProductName").ToString()

            Dim designName As String = thisData("DesignName").ToString()
            Dim blindName As String = thisData("BlindName").ToString()

            Dim tubeName As String = thisData("TubeName").ToString()
            Dim controlName As String = thisData("ControlName").ToString()
            Dim colourName As String = thisData("ColourName").ToString()

            If designName = "Aluminium Blind" Then
                Dim subType As String = thisData("SubType").ToString()
                Dim venId As String = GetItemData("SELECT VenId FROM ProductKits WHERE ProductId='" & productId & "'")

                Dim venIdB As String = String.Empty
                If subType.Contains("2 on 1") Then venIdB = venId

                Dim newSubType As String = "Single"
                If subType = "2 on 1 Left-Left" Then newSubType = "2 on 1 Aluminium Left-Left"
                If subType = "2 on 1 Right-Right" Then newSubType = "2 on 1 Aluminium Right-Right"
                If subType = "2 on 1 Left-Right" Then newSubType = "2 on 1 Aluminium Left-Right"

                If String.IsNullOrEmpty(venId) Then Continue While

                writer.WriteStartElement("OrderDetails")
                writer.WriteAttributeString("OrddID", thisData("Id").ToString())
                writer.WriteAttributeString("FKOrdID", thisData("HeaderId").ToString())
                writer.WriteAttributeString("Qty", thisData("Qty").ToString())
                writer.WriteAttributeString("Room", thisData("Room").ToString())
                writer.WriteAttributeString("Mounting", thisData("Mounting").ToString())

                writer.WriteAttributeString("BlindType", "Ven")
                writer.WriteAttributeString("OrderType", blindName)
                writer.WriteAttributeString("SubType", newSubType)
                writer.WriteAttributeString("IDHK", venId)
                writer.WriteAttributeString("IDHK2", venIdB)
                writer.WriteAttributeString("Width", thisData("Width").ToString())
                writer.WriteAttributeString("Width2c", thisData("WidthB").ToString())
                writer.WriteAttributeString("Drop", thisData("Drop").ToString())
                writer.WriteAttributeString("Drop2c", thisData("DropB").ToString())
                writer.WriteAttributeString("ControlPosition", thisData("ControlPosition").ToString())
                writer.WriteAttributeString("ControlPosition_LSDouble3rd", thisData("ControlPositionB").ToString())
                writer.WriteAttributeString("TilterPosition", thisData("TilterPosition").ToString())
                writer.WriteAttributeString("Roll_LSDouble3rd", thisData("TilterPositionB").ToString())
                writer.WriteAttributeString("PullCordLength", thisData("ControlLengthValue").ToString())
                writer.WriteAttributeString("Chain_LSDouble3rd", thisData("ControlLengthValueB").ToString())
                writer.WriteAttributeString("WandLength", thisData("WandLengthValue").ToString())
                writer.WriteAttributeString("Additional", thisData("WandLengthValueB").ToString())
                writer.WriteAttributeString("Supply", thisData("Supply").ToString())

                writer.WriteAttributeString("TotalItems", thisData("TotalItems").ToString())
                writer.WriteAttributeString("MarkUp", thisData("MarkUp").ToString())
                writer.WriteAttributeString("Notes", thisData("Notes").ToString())
                writer.WriteEndElement()
            End If

            If designName = "Cellular Shades" Then
                Dim kitId As String = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "'")
                Dim webFabricId As String = thisData("FabricColourId").ToString()
                Dim webFabricIdB As String = thisData("FabricColourIdB").ToString()

                Dim boeFabricId As String = GetItemData("SELECT BoeId FROM FabricColours WHERE Id='" & webFabricId & "'")
                Dim boeFabricIdB As String = GetItemData("SELECT BoeId FROM FabricColours WHERE Id='" & webFabricIdB & "'")

                If String.IsNullOrEmpty(kitId) Then Continue While
                If String.IsNullOrEmpty(boeFabricId) Then Continue While

                Dim controlPosition As String = thisData("ControlPosition").ToString()
                If controlPosition = "Both Sides" Then controlPosition = "Left and Right"

                writer.WriteStartElement("OrderDetails")
                writer.WriteAttributeString("OrddID", thisData("Id").ToString())
                writer.WriteAttributeString("FKOrdID", thisData("HeaderId").ToString())
                writer.WriteAttributeString("Qty", thisData("Qty").ToString())
                writer.WriteAttributeString("Room", thisData("Room").ToString())
                writer.WriteAttributeString("Mounting", thisData("Mounting").ToString())

                writer.WriteAttributeString("BlindType", "Cellular Shades")
                writer.WriteAttributeString("OrderType", blindName)
                writer.WriteAttributeString("IDHK", kitId)
                writer.WriteAttributeString("FabricID", boeFabricId)
                writer.WriteAttributeString("FabricID_DoubleBracket", boeFabricIdB)
                writer.WriteAttributeString("Width", thisData("Width").ToString())
                writer.WriteAttributeString("Drop", thisData("Drop").ToString())
                writer.WriteAttributeString("ControlPosition", controlPosition)
                writer.WriteAttributeString("ControlLength", thisData("ControlLengthValue").ToString())
                writer.WriteAttributeString("Supply", thisData("Supply").ToString())

                writer.WriteAttributeString("TotalItems", thisData("TotalItems").ToString())
                writer.WriteAttributeString("MarkUp", thisData("MarkUp").ToString())
                writer.WriteAttributeString("Notes", thisData("Notes").ToString())
                writer.WriteEndElement()
            End If

            If designName = "Curtain" Then
                Dim heading As String = thisData("Heading").ToString()
                Dim headingB As String = thisData("HeadingB").ToString()

                Dim orderType As String = blindName

                Dim kitName As String = String.Format("{0} {1}", blindName, heading)
                Dim kitId As String = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "'")
                Dim kitIdB As String = String.Empty

                If blindName = "Track Only" Then
                    kitId = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "'")
                End If
                If blindName = "Fabric Only" Then
                    kitId = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "'")
                End If

                If blindName = "Complete Set (Double)" Then
                    orderType = "Double Bracket"

                    kitName = String.Format("{0} {1}", blindName, headingB)
                    kitIdB = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "'")
                End If

                Dim webFabricId As String = thisData("FabricColourId").ToString()
                Dim boeFabricId As String = GetItemData("SELECT BoeId FROM FabricColours WHERE Id='" & webFabricId & "'")

                Dim webFabricIdB As String = thisData("FabricColourIdB").ToString()
                Dim boeFabricIdB As String = GetItemData("SELECT BoeId FROM FabricColours WHERE Id='" & webFabricIdB & "'")

                If String.IsNullOrEmpty(kitId) Then Continue While

                writer.WriteStartElement("OrderDetails")
                writer.WriteAttributeString("OrddID", thisData("Id").ToString())
                writer.WriteAttributeString("FKOrdID", thisData("HeaderId").ToString())
                writer.WriteAttributeString("Qty", thisData("Qty").ToString())
                writer.WriteAttributeString("Room", thisData("Room").ToString())
                writer.WriteAttributeString("Mounting", thisData("Mounting").ToString())
                writer.WriteAttributeString("BlindType", "Curtain")
                writer.WriteAttributeString("OrderType", orderType)
                writer.WriteAttributeString("IDHK", kitId)
                writer.WriteAttributeString("IDHK2", kitIdB)
                writer.WriteAttributeString("FabricID", boeFabricId)
                writer.WriteAttributeString("FabricID_DoubleBracket", boeFabricIdB)
                writer.WriteAttributeString("Width", thisData("Width").ToString())
                writer.WriteAttributeString("Drop", thisData("Drop").ToString())
                writer.WriteAttributeString("StackOption", thisData("StackPosition").ToString())
                writer.WriteAttributeString("ControlType", "Hand")
                writer.WriteAttributeString("BottomRailColour", thisData("BottomHem").ToString())
                writer.WriteAttributeString("Additional1", thisData("Supply").ToString())
                writer.WriteAttributeString("ControlLength", thisData("ReturnLengthValue").ToString())
                writer.WriteAttributeString("ControlLength2", thisData("ReturnLengthValueB").ToString())
                writer.WriteAttributeString("Tracks", thisData("TrackType").ToString())
                writer.WriteAttributeString("LayoutCode", thisData("TrackDraw").ToString())
                writer.WriteAttributeString("WandLength", thisData("LouvreSize").ToString())

                writer.WriteAttributeString("TotalItems", thisData("TotalItems").ToString())
                writer.WriteAttributeString("MarkUp", thisData("MarkUp").ToString())
                writer.WriteAttributeString("Notes", thisData("Notes").ToString())
                writer.WriteEndElement()
            End If

            If designName = "Design Shades" Then
                Dim kitId As String = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "'")
                Dim webFabricId As String = thisData("FabricColourId").ToString()
                Dim boeFabricId As String = GetItemData("SELECT BoeId FROM FabricColours WHERE Id='" & webFabricId & "'")

                Dim chainId As String = thisData("ChainId").ToString()
                Dim boeChainId As String = String.Empty
                Dim controlColour As String = GetItemData("SELECT Name FROM Chains WHERE Id='" & chainId & "'")
                Dim controlLength As String = thisData("ControlLengthValue").ToString()
                If controlName = "Wand" Then
                    controlColour = thisData("WandColour").ToString()
                    controlLength = thisData("WandLengthValue").ToString()
                End If
                If controlName = "Chain" Then
                    boeChainId = GetItemData("SELECT BoeId FROM Chains WHERE Id='" & chainId & "'")
                End If

                If String.IsNullOrEmpty(kitId) Then Continue While
                If String.IsNullOrEmpty(boeFabricId) Then Continue While

                writer.WriteStartElement("OrderDetails")
                writer.WriteAttributeString("OrddID", thisData("Id").ToString())
                writer.WriteAttributeString("FKOrdID", thisData("HeaderId").ToString())
                writer.WriteAttributeString("Qty", thisData("Qty").ToString())
                writer.WriteAttributeString("Room", thisData("Room").ToString())
                writer.WriteAttributeString("Mounting", thisData("Mounting").ToString())

                writer.WriteAttributeString("BlindType", "Profile")
                writer.WriteAttributeString("OrderType", "Profile")
                writer.WriteAttributeString("IDHK", kitId)
                writer.WriteAttributeString("IDChain", boeChainId)
                writer.WriteAttributeString("FabricID", boeFabricId)
                writer.WriteAttributeString("Width", thisData("Width").ToString())
                writer.WriteAttributeString("Drop", thisData("Drop").ToString())
                writer.WriteAttributeString("ControlPosition", thisData("ControlPosition").ToString())
                writer.WriteAttributeString("StackOption", thisData("StackPosition").ToString())
                writer.WriteAttributeString("ControlLength", controlLength)
                writer.WriteAttributeString("ControlColour", controlColour)

                writer.WriteAttributeString("TotalItems", thisData("TotalItems").ToString())
                writer.WriteAttributeString("MarkUp", thisData("MarkUp").ToString())
                writer.WriteAttributeString("Notes", thisData("Notes").ToString())
                writer.WriteEndElement()
            End If

            If designName = "Linea Valance" Then
                Dim kitId As String = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "'")
                Dim webFabricId As String = thisData("FabricColourId").ToString()
                Dim boeFabricId As String = GetItemData("SELECT BoeId FROM FabricColours WHERE Id='" & webFabricId & "'")

                Dim valanceSize As String = String.Empty
                If blindName = "Valance 100mm" Then valanceSize = "100"
                If blindName = "Valance 140mm" Then valanceSize = "140"

                Dim valancePosition As String = thisData("ReturnPosition").ToString()
                If valancePosition = "Both Sides" Then valancePosition = "RL"

                If String.IsNullOrEmpty(kitId) Then Continue While

                writer.WriteStartElement("OrderDetails")
                writer.WriteAttributeString("OrddID", thisData("Id").ToString())
                writer.WriteAttributeString("FKOrdID", thisData("HeaderId").ToString())
                writer.WriteAttributeString("Qty", thisData("Qty").ToString())
                writer.WriteAttributeString("Room", thisData("Room").ToString())
                writer.WriteAttributeString("Mounting", thisData("Mounting").ToString())

                writer.WriteAttributeString("BlindType", "Linea Valance")
                writer.WriteAttributeString("OrderType", blindName)
                writer.WriteAttributeString("IDHK", kitId)
                writer.WriteAttributeString("FabricInsert", thisData("FabricInsert").ToString())
                writer.WriteAttributeString("FabricID", boeFabricId)
                writer.WriteAttributeString("Width", thisData("Width").ToString())
                writer.WriteAttributeString("Drop", thisData("Drop").ToString())
                writer.WriteAttributeString("LineaValance_BracketType", thisData("BracketType").ToString())
                writer.WriteAttributeString("RollDirection", thisData("IsBlindIn").ToString())
                writer.WriteAttributeString("Colour", "25")
                writer.WriteAttributeString("ValanceSize", valanceSize)
                writer.WriteAttributeString("ValancePosition", valancePosition)
                writer.WriteAttributeString("ValanceReturnSize", thisData("ReturnLengthValue").ToString())

                writer.WriteAttributeString("TotalItems", thisData("TotalItems").ToString())
                writer.WriteAttributeString("MarkUp", thisData("MarkUp").ToString())
                writer.WriteAttributeString("Notes", thisData("Notes").ToString())
                writer.WriteEndElement()
            End If

            If designName = "Panel Glide" Then
                Dim idhk As String = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "'")
                Dim webFabricId As String = thisData("FabricColourId").ToString()
                Dim boeFabricId As String = GetItemData("SELECT BoeId FROM FabricColours WHERE Id='" & webFabricId & "'")

                Dim layoutCode As String = thisData("LayoutCode").ToString()
                If layoutCode = "S" Then
                    layoutCode = thisData("LayoutCodeCustom").ToString()
                End If

                writer.WriteStartElement("OrderDetails")
                writer.WriteAttributeString("OrddID", thisData("Id").ToString())
                writer.WriteAttributeString("FKOrdID", thisData("HeaderId").ToString())
                writer.WriteAttributeString("Qty", thisData("Qty").ToString())
                writer.WriteAttributeString("Room", thisData("Room").ToString())
                writer.WriteAttributeString("Mounting", thisData("Mounting").ToString())
                writer.WriteAttributeString("BlindType", "Panel Glide")
                writer.WriteAttributeString("OrderType", blindName)
                writer.WriteAttributeString("TubeName", tubeName)
                writer.WriteAttributeString("IDHK", idhk)
                writer.WriteAttributeString("FabricID", boeFabricId)
                writer.WriteAttributeString("Width", thisData("Width").ToString())
                writer.WriteAttributeString("Drop", thisData("Drop").ToString())
                writer.WriteAttributeString("Baton", thisData("Batten").ToString())
                writer.WriteAttributeString("LayoutCode", layoutCode)
                writer.WriteAttributeString("Tracks", thisData("TrackType").ToString())
                writer.WriteAttributeString("Panel", thisData("PanelQty").ToString())
                writer.WriteAttributeString("WandLength", thisData("WandLengthValue").ToString())
                writer.WriteEndElement()
            End If

            If designName = "Pelmet" Then
                Dim kitId As String = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "'")
                Dim webFabricId As String = thisData("FabricColourId").ToString()
                Dim boeFabricId As String = GetItemData("SELECT BoeId FROM FabricColours WHERE Id='" & webFabricId & "'")

                If String.IsNullOrEmpty(kitId) OrElse String.IsNullOrEmpty(boeFabricId) Then Continue While

                Dim drop As String = String.Empty
                If blindName = "Pelmet 140mm" Then drop = "140"
                If blindName = "Pelmet 200mm" Then drop = "200"

                writer.WriteStartElement("OrderDetails")
                writer.WriteAttributeString("OrddID", thisData("Id").ToString())
                writer.WriteAttributeString("FKOrdID", thisData("HeaderId").ToString())
                writer.WriteAttributeString("Qty", thisData("Qty").ToString())
                writer.WriteAttributeString("Room", thisData("Room").ToString())
                writer.WriteAttributeString("BlindType", designName)
                writer.WriteAttributeString("OrderType", "Fabric Pelmet")
                writer.WriteAttributeString("IDHK", kitId)
                writer.WriteAttributeString("FabricID", boeFabricId)
                writer.WriteAttributeString("Width", thisData("Width").ToString())
                writer.WriteAttributeString("Drop", drop)
                writer.WriteAttributeString("Mounting", thisData("Mounting").ToString())
                writer.WriteAttributeString("Baton", thisData("Batten").ToString())
                writer.WriteAttributeString("ValancePosition", thisData("ReturnPosition").ToString())
                writer.WriteAttributeString("ValanceReturnSize", thisData("ReturnLengthValue").ToString())
                writer.WriteAttributeString("ControlLength2", thisData("ReturnLengthValueB").ToString())
                writer.WriteAttributeString("TotalItems", thisData("TotalItems").ToString())
                writer.WriteAttributeString("MarkUp", thisData("MarkUp").ToString())
                writer.WriteAttributeString("Notes", thisData("Notes").ToString())
                writer.WriteEndElement()
            End If

            If designName = "Roman Blind" Then
                Dim kitId As String = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "'")
                Dim webFabricId As String = thisData("FabricColourId").ToString()
                Dim boeFabricId As String = GetItemData("SELECT BoeId FROM FabricColours WHERE Id='" & webFabricId & "'")

                Dim chainId As String = thisData("ChainId").ToString()
                Dim boeChainId As String = String.Empty
                Dim mechanismeOption As String = String.Empty
                Dim controlColour As String = GetItemData("SELECT Name FROM Chains WHERE Id='" & chainId & "'")
                If controlName = "Chain" Then
                    boeChainId = GetItemData("SELECT BoeId FROM Chains WHERE Id='" & chainId & "'")
                    mechanismeOption = "Covered"
                End If

                If String.IsNullOrEmpty(kitId) Then Continue While
                If String.IsNullOrEmpty(boeFabricId) Then Continue While

                Dim valanceOption As String = thisData("ValanceOption").ToString()
                If valanceOption = "Cover Valance and Cord at Front" Then valanceOption = "Cover Valance"
                If valanceOption = "No Cover Valance and Cord at Back" Then valanceOption = "No Cover Valance"

                writer.WriteStartElement("OrderDetails")
                writer.WriteAttributeString("OrddID", thisData("Id").ToString())
                writer.WriteAttributeString("FKOrdID", thisData("HeaderId").ToString())
                writer.WriteAttributeString("Qty", thisData("Qty").ToString())
                writer.WriteAttributeString("Room", thisData("Room").ToString())
                writer.WriteAttributeString("Mounting", thisData("Mounting").ToString())

                writer.WriteAttributeString("BlindType", "Roman")
                writer.WriteAttributeString("OrderType", blindName)
                writer.WriteAttributeString("TubeName", tubeName)
                writer.WriteAttributeString("IDHK", kitId)
                writer.WriteAttributeString("FabricID", boeFabricId)
                writer.WriteAttributeString("IDChain", boeChainId)
                writer.WriteAttributeString("Width", thisData("Width").ToString())
                writer.WriteAttributeString("Drop", thisData("Drop").ToString())
                writer.WriteAttributeString("Baton", thisData("Batten").ToString())
                writer.WriteAttributeString("RomanHR", "Timber")
                writer.WriteAttributeString("ControlPosition", thisData("ControlPosition").ToString())
                writer.WriteAttributeString("RomanMechanismType", controlName)
                writer.WriteAttributeString("RomanChainColour", controlColour)
                writer.WriteAttributeString("RomanChainOption", mechanismeOption)
                writer.WriteAttributeString("RomanChainLength", thisData("ControlLengthValue").ToString())
                writer.WriteAttributeString("RomanValanceOption", valanceOption)

                writer.WriteAttributeString("TotalItems", thisData("TotalItems").ToString())
                writer.WriteAttributeString("MarkUp", thisData("MarkUp").ToString())
                writer.WriteAttributeString("Notes", thisData("Notes").ToString())
                writer.WriteEndElement()
            End If

            If designName = "Soft Roman" Then
                Dim kitId As String = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "'")
                Dim webFabricId As String = thisData("FabricColourId").ToString()
                Dim boeFabricId As String = GetItemData("SELECT BoeId FROM FabricColours WHERE Id='" & webFabricId & "'")

                Dim chainId As String = thisData("ChainId").ToString()
                Dim boeChainId As String = String.Empty
                Dim mechanismeOption As String = String.Empty
                Dim controlColour As String = GetItemData("SELECT Name FROM Chains WHERE Id='" & chainId & "'")
                If controlName = "Chain" Then
                    boeChainId = GetItemData("SELECT BoeId FROM Chains WHERE Id='" & chainId & "'")
                    mechanismeOption = "Covered"
                End If

                If String.IsNullOrEmpty(kitId) Then Continue While
                If String.IsNullOrEmpty(boeFabricId) Then Continue While

                Dim valanceOption As String = thisData("ValanceOption").ToString()
                If valanceOption = "Cover Valance and Cord at Front" Then valanceOption = "Cover Valance"
                If valanceOption = "No Cover Valance and Cord at Back" Then valanceOption = "No Cover Valance"

                writer.WriteStartElement("OrderDetails")
                writer.WriteAttributeString("OrddID", thisData("Id").ToString())
                writer.WriteAttributeString("FKOrdID", thisData("HeaderId").ToString())
                writer.WriteAttributeString("Qty", thisData("Qty").ToString())
                writer.WriteAttributeString("Room", thisData("Room").ToString())
                writer.WriteAttributeString("Mounting", thisData("Mounting").ToString())

                writer.WriteAttributeString("BlindType", "Soft Roman")
                writer.WriteAttributeString("OrderType", blindName)
                writer.WriteAttributeString("TubeName", tubeName)
                writer.WriteAttributeString("IDHK", kitId)
                writer.WriteAttributeString("FabricID", boeFabricId)
                writer.WriteAttributeString("IDChain", boeChainId)
                writer.WriteAttributeString("Width", thisData("Width").ToString())
                writer.WriteAttributeString("Drop", thisData("Drop").ToString())
                writer.WriteAttributeString("Baton", thisData("Batten").ToString())
                writer.WriteAttributeString("RomanHR", "Timber")
                writer.WriteAttributeString("ControlPosition", thisData("ControlPosition").ToString())
                writer.WriteAttributeString("RomanMechanismType", controlName)
                writer.WriteAttributeString("RomanChainColour", controlColour)
                writer.WriteAttributeString("RomanChainOption", mechanismeOption)
                writer.WriteAttributeString("RomanChainLength", thisData("ControlLengthValue").ToString())
                writer.WriteAttributeString("RomanValanceOption", valanceOption)

                writer.WriteAttributeString("TotalItems", thisData("TotalItems").ToString())
                writer.WriteAttributeString("MarkUp", thisData("MarkUp").ToString())
                writer.WriteAttributeString("Notes", thisData("Notes").ToString())
                writer.WriteEndElement()
            End If

            If designName = "Privacy Venetian" Then
                Dim venId As String = GetItemData("SELECT VenId FROM ProductKits WHERE ProductId='" & productId & "'")

                writer.WriteStartElement("OrderDetails")
                writer.WriteAttributeString("OrddID", thisData("Id").ToString())
                writer.WriteAttributeString("FKOrdID", thisData("HeaderId").ToString())
                writer.WriteAttributeString("Qty", thisData("Qty").ToString())
                writer.WriteAttributeString("Room", thisData("Room").ToString())
                writer.WriteAttributeString("Mounting", thisData("Mounting").ToString())

                writer.WriteAttributeString("BlindType", "Smart Privacy")
                writer.WriteAttributeString("OrderType", blindName)
                writer.WriteAttributeString("IDHK", venId)
                writer.WriteAttributeString("Width", thisData("Width").ToString())
                writer.WriteAttributeString("Drop", thisData("Drop").ToString())
                writer.WriteAttributeString("ControlPosition", thisData("ControlPosition").ToString())
                writer.WriteAttributeString("TilterPosition", thisData("TilterPosition").ToString())
                writer.WriteAttributeString("PullCordLength", thisData("ControlLengthValue").ToString())
                writer.WriteAttributeString("Supply", thisData("Supply").ToString())

                writer.WriteAttributeString("TotalItems", thisData("TotalItems").ToString())
                writer.WriteAttributeString("MarkUp", thisData("MarkUp").ToString())
                writer.WriteAttributeString("Notes", thisData("Notes").ToString())
                writer.WriteEndElement()
            End If

            If designName = "Venetian Blind" Then
                Dim subType As String = thisData("SubType").ToString()
                Dim tassel As String = thisData("Tassel").ToString()
                Dim venId As String = GetItemData("SELECT VenId FROM ProductKits WHERE ProductId='" & productId & "' AND BlindStatus='Semi Metal'")
                If tassel = "Gold" OrElse tassel = "Antique Brass" Then
                    venId = GetItemData("SELECT VenId FROM ProductKits WHERE ProductId='" & productId & "' AND BlindStatus='Metal'")
                End If
                Dim venIdB As String = String.Empty
                Dim venIdC As String = String.Empty
                If subType.Contains("2 on 1") Then venIdB = venId
                If subType.Contains("3 on 1") Then venIdB = venId : venIdC = venId

                Dim newSubType As String = "Single"
                If subType = "2 on 1 Left-Left" Then newSubType = "2 on 1 Venetian Left-Left"
                If subType = "2 on 1 Right-Right" Then newSubType = "2 on 1 Venetian Right-Right"
                If subType = "2 on 1 Left-Right" Then newSubType = "2 on 1 Venetian Left-Right"
                If subType = "3 on 1 Left-Left-Right" Then newSubType = "3 on 1 Venetian Left-Left-Right"
                If subType = "3 on 1 Left-Right-Right" Then newSubType = "3 on 1 Venetian Left-Right-Right"

                Dim valancePosition As String = thisData("ReturnPosition").ToString()
                If valancePosition = "Both Sides" Then valancePosition = "Right and Left"

                If String.IsNullOrEmpty(venId) Then Continue While

                writer.WriteStartElement("OrderDetails")
                writer.WriteAttributeString("OrddID", thisData("Id").ToString())
                writer.WriteAttributeString("FKOrdID", thisData("HeaderId").ToString())
                writer.WriteAttributeString("Qty", thisData("Qty").ToString())
                writer.WriteAttributeString("Room", thisData("Room").ToString())
                writer.WriteAttributeString("Mounting", thisData("Mounting").ToString())

                writer.WriteAttributeString("BlindType", "Venetian")
                writer.WriteAttributeString("OrderType", blindName)
                writer.WriteAttributeString("SubType", newSubType)
                writer.WriteAttributeString("IDHK", venId)
                writer.WriteAttributeString("IDHK2", venIdB)
                writer.WriteAttributeString("IDHK3", venIdC)
                writer.WriteAttributeString("Width", thisData("Width").ToString())
                writer.WriteAttributeString("Width2c", thisData("WidthB").ToString())
                writer.WriteAttributeString("Width_LSIdp3rd", thisData("WidthC").ToString())
                writer.WriteAttributeString("Drop", thisData("Drop").ToString())
                writer.WriteAttributeString("ControlPosition", thisData("ControlPosition").ToString())
                writer.WriteAttributeString("ControlPosition_LSDouble3rd", thisData("ControlPositionB").ToString())
                writer.WriteAttributeString("Additional", thisData("ControlPositionC").ToString())
                writer.WriteAttributeString("TilterPosition", thisData("TilterPosition").ToString())
                writer.WriteAttributeString("Additional2", thisData("TilterPositionB").ToString())
                writer.WriteAttributeString("Additional1", thisData("TilterPositionC").ToString())
                writer.WriteAttributeString("PullCordLength", thisData("ControlLengthValue").ToString())
                writer.WriteAttributeString("ChainLengthdb4", thisData("ControlLengthValueB").ToString())
                writer.WriteAttributeString("ControlLength", thisData("ControlLengthValueC").ToString())
                writer.WriteAttributeString("ControlColour", thisData("Tassel").ToString())
                writer.WriteAttributeString("ValanceType", thisData("ValanceType").ToString())
                writer.WriteAttributeString("ValanceSize", thisData("ValanceSizeValue").ToString())
                writer.WriteAttributeString("ValancePosition", valancePosition)
                writer.WriteAttributeString("ValanceReturnSize", thisData("ReturnLengthValue").ToString())
                writer.WriteAttributeString("Supply", thisData("Supply").ToString())
                writer.WriteAttributeString("WandLength", thisData("WandLengthValue").ToString())

                writer.WriteAttributeString("TotalItems", thisData("TotalItems").ToString())
                writer.WriteAttributeString("MarkUp", thisData("MarkUp").ToString())
                writer.WriteAttributeString("Notes", thisData("Notes").ToString())
                writer.WriteEndElement()
            End If

            If designName = "Vertical" Then
                Dim kitId As String = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "'")

                Dim webFabricId As String = thisData("FabricColourId").ToString()
                Dim boeFabricId As String = GetItemData("SELECT BoeId FROM FabricColours WHERE Id='" & webFabricId & "'")

                Dim chainId As String = thisData("ChainId").ToString()
                Dim boeChainId As String = String.Empty
                If controlName = "Chain" Then boeChainId = GetItemData("SELECT BoeId FROM Chains WHERE Id='" & chainId & "'")
                Dim controlLength As String = thisData("ControlLengthValue").ToString()
                If controlName = "Wand" Then controlLength = thisData("WandLengthValue").ToString()

                Dim orderType As String = "Vertical"
                If blindName = "Slat Only" Then orderType = "Vertical Slat Only"

                If String.IsNullOrEmpty(kitId) Then Continue While

                Dim stackPosition As String = String.Empty
                If Not String.IsNullOrEmpty(thisData("StackPosition")) Then
                    stackPosition = "Stack " & thisData("StackPosition").ToString()
                End If

                Dim sloping As String = "0"
                If thisData("Sloping").ToString() = "Yes" Then sloping = "1"

                writer.WriteStartElement("OrderDetails")
                writer.WriteAttributeString("OrddID", thisData("Id").ToString())
                writer.WriteAttributeString("FKOrdID", thisData("HeaderId").ToString())
                writer.WriteAttributeString("Qty", thisData("Qty").ToString())
                writer.WriteAttributeString("Room", thisData("Room").ToString())
                writer.WriteAttributeString("Mounting", thisData("Mounting").ToString())

                writer.WriteAttributeString("BlindType", orderType)
                writer.WriteAttributeString("OrderType", blindName)
                writer.WriteAttributeString("IDHK", kitId)
                writer.WriteAttributeString("Width", thisData("Width").ToString())
                writer.WriteAttributeString("Drop", thisData("Drop").ToString())
                writer.WriteAttributeString("ControlType", controlName)
                writer.WriteAttributeString("StackOption", stackPosition)
                writer.WriteAttributeString("ControlLength", controlLength)
                writer.WriteAttributeString("IDChain", boeChainId)
                writer.WriteAttributeString("FabricID", boeFabricId)
                writer.WriteAttributeString("FabricInsert", thisData("FabricInsert").ToString())
                writer.WriteAttributeString("BottomJoin", thisData("BottomJoining").ToString())
                writer.WriteAttributeString("Vertical_HeadrailColour", colourName)
                writer.WriteAttributeString("Vertical_WandCordColour", colourName)
                writer.WriteAttributeString("ControlPosition", thisData("ControlPosition").ToString())
                writer.WriteAttributeString("Extbracket", thisData("BracketExtension").ToString())
                writer.WriteAttributeString("RomanValanceOption", sloping)

                writer.WriteAttributeString("TotalItems", thisData("TotalItems").ToString())
                writer.WriteAttributeString("MarkUp", thisData("MarkUp").ToString())
                writer.WriteAttributeString("Notes", thisData("Notes").ToString())
                writer.WriteEndElement()
            End If

            If designName = "Saphora Drape" Then
                Dim kitId As String = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "'")

                Dim webFabricId As String = thisData("FabricColourId").ToString()
                Dim boeFabricId As String = GetItemData("SELECT BoeId FROM FabricColours WHERE Id='" & webFabricId & "'")

                Dim chainId As String = thisData("ChainId").ToString()
                Dim boeChainId As String = String.Empty
                If controlName = "Chain" Then
                    boeChainId = GetItemData("SELECT BoeId FROM Chains WHERE Id='" & chainId & "'")
                End If
                Dim controlLength As String = thisData("ControlLengthValue").ToString()
                If controlName = "Wand" Then
                    controlLength = thisData("WandLengthValue").ToString()
                End If

                If String.IsNullOrEmpty(kitId) Then Continue While
                If String.IsNullOrEmpty(boeFabricId) Then Continue While

                writer.WriteStartElement("OrderDetails")
                writer.WriteAttributeString("OrddID", thisData("Id").ToString())
                writer.WriteAttributeString("FKOrdID", thisData("HeaderId").ToString())
                writer.WriteAttributeString("Qty", thisData("Qty").ToString())
                writer.WriteAttributeString("Room", thisData("Room").ToString())
                writer.WriteAttributeString("Mounting", thisData("Mounting").ToString())

                writer.WriteAttributeString("BlindType", designName)
                writer.WriteAttributeString("OrderType", blindName)
                writer.WriteAttributeString("IDHK", kitId)
                writer.WriteAttributeString("Width", thisData("Width").ToString())
                writer.WriteAttributeString("Drop", thisData("Drop").ToString())
                writer.WriteAttributeString("ControlType", controlName)
                writer.WriteAttributeString("StackOption", thisData("StackPosition").ToString())
                writer.WriteAttributeString("ControlLength", controlLength)
                writer.WriteAttributeString("IDChain", boeChainId)
                writer.WriteAttributeString("FabricID", boeFabricId)
                writer.WriteAttributeString("FabricInsert", thisData("FabricInsert").ToString())
                writer.WriteAttributeString("BottomJoin", thisData("BottomJoining").ToString())
                writer.WriteAttributeString("Vertical_HeadrailColour", colourName)
                writer.WriteAttributeString("Vertical_WandCordColour", colourName)
                writer.WriteAttributeString("ControlPosition", thisData("ControlPosition").ToString())
                writer.WriteAttributeString("Extbracket", thisData("BracketExtension").ToString())

                writer.WriteAttributeString("TotalItems", thisData("TotalItems").ToString())
                writer.WriteAttributeString("MarkUp", thisData("MarkUp").ToString())
                writer.WriteAttributeString("Notes", thisData("Notes").ToString())
                writer.WriteEndElement()
            End If

            If designName = "Roller Blind" Then
                Dim webFabricId As String = thisData("FabricColourId").ToString()
                Dim boeFabricId As String = GetItemData("SELECT BoeId FROM FabricColours WHERE Id='" & webFabricId & "'")

                Dim chainId As String = thisData("ChainId").ToString()
                Dim boeChainId As String = GetItemData("SELECT BoeId FROM Chains WHERE Id='" & chainId & "'")

                Dim bottomColourId As String = thisData("BottomColourId").ToString()
                Dim boeBottomId As String = GetItemData("SELECT BoeId FROM BottomColours WHERE Id='" & bottomColourId & "'")

                Dim orderType As String = String.Empty
                If blindName = "Single Blind" Then orderType = "Regular Chain"
                If blindName = "Dual Blinds" Then orderType = "Double Bracket"
                If blindName = "Link 2 Blinds Dependent" Then orderType = "Link System Dependent 2 Blinds"
                If blindName = "Link 2 Blinds Independent" Then orderType = "Link System Independent 2 Blinds"
                If blindName = "Link 3 Blinds Dependent" Then orderType = "Link System Dependent 3 Blinds"
                If blindName = "Link 3 Blinds Independent with Dependent" Then orderType = "Link System Independent 3 Blinds"
                If blindName = "DB Link 2 Blinds Dependent" Then orderType = "Double and Link System Dependent"
                If blindName = "DB Link 2 Blinds Independent" Then orderType = "Double and Link System Independent"
                If blindName = "DB Link 3 Blinds Dependent" Then orderType = "Double and Link System Dependent"
                If blindName = "DB Link 3 Blinds Independent with Dependent" Then orderType = "Double and Link System Independent"

                Dim kitId As String = String.Empty
                Dim kitIdB As String = String.Empty
                Dim kitIdC As String = String.Empty
                Dim kitIdD As String = String.Empty
                Dim kitIdE As String = String.Empty
                Dim kitIdF As String = String.Empty
                Dim kitIdG As String = String.Empty
                Dim kitIdH As String = String.Empty

                Dim springAssist As String = thisData("SpringAssist").ToString()

                If blindName = "Single Blind" OrElse blindName = "Full Cassette" OrElse blindName = "Semi Cassette" OrElse blindName = "Wire Guide" Then
                    kitId = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "'")
                    Dim kitName As String = String.Empty

                    If tubeName = "Standard" AndAlso controlName = "Chain" Then
                        kitName = String.Format("{0} (LD)", productName)

                        Dim width As Integer = thisData("Width")
                        If width > 1810 Then kitName = String.Format("{0} (HD)", productName)

                        kitId = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "'")
                    End If

                    If tubeName = "Acmeda 49mm" AndAlso controlName = "Chain" Then
                        kitName = productName
                        If springAssist = "Yes" Then
                            kitName = String.Format("{0} (Spring Assist)", productName)
                        End If
                        kitId = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "'")
                    End If

                    If String.IsNullOrEmpty(kitId) Then Continue While
                    If String.IsNullOrEmpty(boeFabricId) Then Continue While

                    Dim flatOption As String = thisData("FlatOption").ToString()
                    If flatOption = "Fabric on Back" Then flatOption = "Fabric on back"
                    If flatOption = "Fabric on Front" Then flatOption = "Fabric on front"

                    writer.WriteStartElement("OrderDetails")
                    writer.WriteAttributeString("OrddID", thisData("Id").ToString())
                    writer.WriteAttributeString("FKOrdID", thisData("HeaderId").ToString())
                    writer.WriteAttributeString("Qty", thisData("Qty").ToString())
                    writer.WriteAttributeString("Room", thisData("Room").ToString())
                    writer.WriteAttributeString("Mounting", thisData("Mounting").ToString())
                    writer.WriteAttributeString("BlindType", "Roller")
                    writer.WriteAttributeString("OrderType", orderType)
                    writer.WriteAttributeString("IDHK", kitId)
                    writer.WriteAttributeString("IDChain", boeChainId)
                    writer.WriteAttributeString("RollerChainLength", thisData("ControlLengthValue").ToString())
                    writer.WriteAttributeString("Additional", thisData("ChainStopper").ToString())
                    writer.WriteAttributeString("FabricID", boeFabricId)
                    writer.WriteAttributeString("IDBottomRail", boeBottomId)
                    writer.WriteAttributeString("FlatBottomOp", flatOption)
                    writer.WriteAttributeString("ControlPosition", thisData("ControlPosition").ToString())
                    writer.WriteAttributeString("RollDirection", thisData("Roll").ToString())
                    writer.WriteAttributeString("Width", thisData("Width").ToString())
                    writer.WriteAttributeString("Drop", thisData("Drop").ToString())
                    writer.WriteAttributeString("Panel", thisData("BracketSize").ToString())
                    writer.WriteAttributeString("PelmetLayout", thisData("TopTrack").ToString())
                    writer.WriteAttributeString("Notes", thisData("Notes").ToString())
                    writer.WriteEndElement()
                End If

                If blindName = "Dual Blinds" Then
                    Dim webFabricIdB As String = thisData("FabricColourIdB").ToString()
                    Dim boeFabricIdB As String = GetItemData("SELECT BoeId FROM FabricColours WHERE Id='" & webFabricIdB & "'")

                    Dim chainIdB As String = thisData("ChainIdB").ToString()
                    Dim boeChainIdB As String = GetItemData("SELECT BoeId FROM Chains WHERE Id='" & chainIdB & "'")

                    Dim bottomColourIdB As String = thisData("BottomColourIdB").ToString()
                    Dim boeBottomIdB As String = GetItemData("SELECT BoeId FROM BottomColours WHERE Id='" & bottomColourIdB & "'")

                    kitId = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "'")
                    kitIdB = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "'")

                    Dim kitName As String = String.Empty

                    If tubeName = "Standard" AndAlso controlName = "Chain" Then
                        kitName = String.Format("{0} (LD)", productName)

                        Dim width As Integer = thisData("Width")
                        Dim widthB As Integer = thisData("WidthB")
                        If width > 1810 OrElse widthB > 1810 Then kitName = String.Format("{0} (HD)", productName)

                        kitId = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "'")
                        kitIdB = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "'")
                    End If

                    If tubeName = "Acmeda 49mm" AndAlso controlName = "Chain" Then
                        kitName = productName
                        If springAssist = "Yes" Then
                            kitName = String.Format("{0} (Spring Assist)", productName)
                        End If
                        kitId = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "'")
                        kitIdB = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "'")
                    End If

                    If String.IsNullOrEmpty(kitId) OrElse String.IsNullOrEmpty(kitIdB) Then Continue While
                    If String.IsNullOrEmpty(boeFabricId) Then Continue While
                    If String.IsNullOrEmpty(boeFabricIdB) Then Continue While

                    Dim flatOption As String = thisData("FlatOption").ToString()
                    Dim flatOptionB As String = thisData("FlatOptionB").ToString()

                    If flatOption = "Fabric on Back" Then flatOption = "Fabric on back"
                    If flatOption = "Fabric on Front" Then flatOption = "Fabric on front"

                    If flatOptionB = "Fabric on Back" Then flatOptionB = "Fabric on back"
                    If flatOptionB = "Fabric on Front" Then flatOptionB = "Fabric on front"

                    writer.WriteStartElement("OrderDetails")
                    writer.WriteAttributeString("OrddID", thisData("Id").ToString())
                    writer.WriteAttributeString("FKOrdID", thisData("HeaderId").ToString())
                    writer.WriteAttributeString("Qty", thisData("Qty").ToString())
                    writer.WriteAttributeString("Room", thisData("Room").ToString())
                    writer.WriteAttributeString("Mounting", thisData("Mounting").ToString())
                    writer.WriteAttributeString("Extbracket", "Yes")
                    writer.WriteAttributeString("BlindType", "Roller")
                    writer.WriteAttributeString("OrderType", orderType)
                    writer.WriteAttributeString("IDHK", kitId)
                    writer.WriteAttributeString("IDHK2", kitIdB)
                    writer.WriteAttributeString("IDChain", boeChainId)
                    writer.WriteAttributeString("RollerChainLength", thisData("ControlLengthValue").ToString())
                    writer.WriteAttributeString("Additional", thisData("ChainStopper").ToString())
                    If controlName = "Chain" Then
                        writer.WriteAttributeString("IDChain_DoubleBracket", boeChainIdB)
                        writer.WriteAttributeString("RolChainLength2c", thisData("ControlLengthValueB").ToString())
                        writer.WriteAttributeString("Additional1", thisData("ChainStopperB").ToString())
                    End If

                    writer.WriteAttributeString("FabricID", boeFabricId)
                    writer.WriteAttributeString("FabricID_DoubleBracket", boeFabricIdB)
                    writer.WriteAttributeString("IDBottomRail", boeBottomId)
                    writer.WriteAttributeString("FlatBottomOp", flatOption)
                    writer.WriteAttributeString("IDBottomRail_DoubleBracket", boeBottomIdB)
                    writer.WriteAttributeString("FlatBottomOp2a", flatOptionB)
                    writer.WriteAttributeString("ControlPosition", thisData("ControlPosition").ToString())
                    writer.WriteAttributeString("RollDirection", thisData("Roll").ToString())
                    writer.WriteAttributeString("ControlPosition_LSDouble3rd", thisData("ControlPositionB").ToString())
                    writer.WriteAttributeString("Roll2a", thisData("RollB").ToString())
                    writer.WriteAttributeString("Width", thisData("Width").ToString())
                    writer.WriteAttributeString("Drop", thisData("Drop").ToString())
                    writer.WriteAttributeString("Panel", thisData("BracketSize").ToString())
                    writer.WriteAttributeString("Notes", thisData("Notes").ToString())
                    writer.WriteAttributeString("ProductId", productId)
                    writer.WriteEndElement()
                End If

                If blindName = "Link 2 Blinds Dependent" Then
                    kitId = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND BlindStatus='Control'")
                    kitIdB = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND BlindStatus='End'")

                    Dim webFabricIdB As String = thisData("FabricColourIdB").ToString()
                    Dim boeFabricIdB As String = GetItemData("SELECT BoeId FROM FabricColours WHERE Id='" & webFabricIdB & "'")

                    Dim bottomColourIdB As String = thisData("BottomColourIdB").ToString()
                    Dim boeBottomIdB As String = GetItemData("SELECT BoeId FROM BottomColours WHERE Id='" & bottomColourIdB & "'")

                    Dim kitName As String = String.Empty

                    If tubeName = "Standard" AndAlso controlName = "Chain" Then
                        kitName = String.Format("{0} (LD)", productName)

                        Dim width As Integer = thisData("Width")
                        Dim widthB As Integer = thisData("WidthB")

                        If width > 1810 OrElse widthB > 1810 Then kitName = String.Format("{0} (HD)", productName)

                        kitId = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='Control'")
                        kitIdB = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='End'")
                    End If

                    If tubeName = "Acmeda 49mm" AndAlso controlName = "Chain" Then
                        kitName = productName
                        If springAssist = "Yes" Then
                            kitName = String.Format("{0} (Spring Assist)", productName)
                        End If
                        kitId = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='Control'")
                        kitIdB = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='End'")
                    End If

                    If String.IsNullOrEmpty(kitId) OrElse String.IsNullOrEmpty(kitIdB) Then Continue While
                    If String.IsNullOrEmpty(boeFabricId) Then Continue While
                    If String.IsNullOrEmpty(boeFabricIdB) Then Continue While

                    Dim flatOption As String = thisData("FlatOption").ToString()
                    Dim flatOptionB As String = thisData("FlatOptionB").ToString()

                    If flatOption = "Fabric on Back" Then flatOption = "Fabric on back"
                    If flatOption = "Fabric on Front" Then flatOption = "Fabric on front"

                    If flatOptionB = "Fabric on Back" Then flatOptionB = "Fabric on back"
                    If flatOptionB = "Fabric on Front" Then flatOptionB = "Fabric on front"

                    writer.WriteStartElement("OrderDetails")
                    writer.WriteAttributeString("OrddID", thisData("Id").ToString())
                    writer.WriteAttributeString("FKOrdID", thisData("HeaderId").ToString())
                    writer.WriteAttributeString("Qty", thisData("Qty").ToString())
                    writer.WriteAttributeString("Room", thisData("Room").ToString())
                    writer.WriteAttributeString("Mounting", thisData("Mounting").ToString())
                    writer.WriteAttributeString("BlindType", "Roller")
                    writer.WriteAttributeString("OrderType", orderType)
                    writer.WriteAttributeString("IDHK", kitId)
                    writer.WriteAttributeString("IDHK2", kitIdB)
                    writer.WriteAttributeString("IDChain", boeChainId)
                    writer.WriteAttributeString("RollerChainLength", thisData("ControlLengthValue").ToString())
                    writer.WriteAttributeString("Additional", thisData("ChainStopper").ToString())
                    writer.WriteAttributeString("FabricID", boeFabricId)
                    writer.WriteAttributeString("FabricID_LinkSys", boeFabricIdB)
                    writer.WriteAttributeString("IDBottomRail", boeBottomId)
                    writer.WriteAttributeString("FlatBottomOp", flatOption)
                    writer.WriteAttributeString("IDBottomRail_LinkSys", boeBottomIdB)
                    writer.WriteAttributeString("FlatBottomOp2b", flatOptionB)
                    writer.WriteAttributeString("ControlPosition", thisData("ControlPosition").ToString())
                    writer.WriteAttributeString("RollDirection", thisData("Roll").ToString())
                    writer.WriteAttributeString("Width", thisData("Width").ToString())
                    writer.WriteAttributeString("Width2b", thisData("WidthB").ToString())
                    writer.WriteAttributeString("Drop", thisData("Drop").ToString())
                    writer.WriteAttributeString("Drop2c", thisData("DropB").ToString())
                    writer.WriteAttributeString("Panel", thisData("BracketSize").ToString())
                    writer.WriteAttributeString("Notes", thisData("Notes").ToString())
                    writer.WriteEndElement()
                End If

                If blindName = "Link 3 Blinds Dependent" Then
                    kitId = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND BlindStatus='Control'")
                    kitIdB = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND BlindStatus='Middle'")
                    kitIdC = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND BlindStatus='End'")

                    Dim kitName As String = String.Empty

                    If tubeName = "Standard" AndAlso controlName = "Chain" Then
                        kitName = String.Format("{0} (LD)", productName)

                        Dim width As Integer = thisData("Width")
                        Dim widthB As Integer = thisData("WidthB")
                        Dim widthC As Integer = thisData("WidthC")

                        If width > 1810 OrElse widthB > 1810 OrElse widthC > 1810 Then kitName = String.Format("{0} (HD)", productName)

                        kitId = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='Control'")
                        kitIdB = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='Middle'")
                        kitIdC = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='End'")
                    End If

                    If tubeName = "Acmeda 49mm" AndAlso controlName = "Chain" Then
                        kitName = productName
                        If springAssist = "Yes" Then
                            kitName = String.Format("{0} (Spring Assist)", productName)
                        End If
                        kitId = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='Control'")
                        kitIdB = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='Middle'")
                        kitIdC = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='End'")
                    End If

                    If String.IsNullOrEmpty(kitId) OrElse String.IsNullOrEmpty(kitIdB) OrElse String.IsNullOrEmpty(kitIdC) Then Continue While

                    Dim webFabricIdB As String = thisData("FabricColourIdB").ToString()
                    Dim webFabricIdC As String = thisData("FabricColourIdC").ToString()

                    Dim boeFabricIdB As String = GetItemData("SELECT BoeId FROM FabricColours WHERE Id='" & webFabricIdB & "'")
                    Dim boeFabricIdC As String = GetItemData("SELECT BoeId FROM FabricColours WHERE Id='" & webFabricIdC & "'")

                    Dim bottomColourIdB As String = thisData("BottomColourIdB").ToString()
                    Dim bottomColourIdC As String = thisData("BottomColourIdC").ToString()

                    Dim boeBottomIdB As String = GetItemData("SELECT BoeId FROM BottomColours WHERE Id='" & bottomColourIdB & "'")
                    Dim boeBottomIdC As String = GetItemData("SELECT BoeId FROM BottomColours WHERE Id='" & bottomColourIdC & "'")

                    Dim flatOption As String = thisData("FlatOption").ToString()
                    Dim flatOptionB As String = thisData("FlatOptionB").ToString()
                    Dim flatOptionC As String = thisData("FlatOptionC").ToString()

                    If flatOption = "Fabric on Back" Then flatOption = "Fabric on back"
                    If flatOption = "Fabric on Front" Then flatOption = "Fabric on front"

                    If flatOptionB = "Fabric on Back" Then flatOptionB = "Fabric on back"
                    If flatOptionB = "Fabric on Front" Then flatOptionB = "Fabric on front"

                    If flatOptionC = "Fabric on Back" Then flatOptionC = "Fabric on back"
                    If flatOptionC = "Fabric on Front" Then flatOptionC = "Fabric on front"

                    writer.WriteStartElement("OrderDetails")
                    writer.WriteAttributeString("OrddID", thisData("Id").ToString())
                    writer.WriteAttributeString("FKOrdID", thisData("HeaderId").ToString())
                    writer.WriteAttributeString("Qty", thisData("Qty").ToString())
                    writer.WriteAttributeString("Room", thisData("Room").ToString())
                    writer.WriteAttributeString("Mounting", thisData("Mounting").ToString())
                    writer.WriteAttributeString("BlindType", "Roller")
                    writer.WriteAttributeString("OrderType", orderType)
                    writer.WriteAttributeString("IDHK", kitId)
                    writer.WriteAttributeString("IDHK2", kitIdB)
                    writer.WriteAttributeString("IDHK3", kitIdC)
                    writer.WriteAttributeString("IDChain", boeChainId)
                    writer.WriteAttributeString("RollerChainLength", thisData("ControlLengthValue").ToString())
                    writer.WriteAttributeString("Additional", thisData("ChainStopper").ToString())
                    writer.WriteAttributeString("FabricID", boeFabricId)
                    writer.WriteAttributeString("FabricID_LinkSys", boeFabricIdB)
                    writer.WriteAttributeString("FabricID_LS3rd", boeFabricIdC)
                    writer.WriteAttributeString("IDBottomRail", boeBottomId)
                    writer.WriteAttributeString("FlatBottomOp", flatOption)
                    writer.WriteAttributeString("IDBottomRail_LinkSys", boeBottomIdB)
                    writer.WriteAttributeString("FlatBottomOp2b", flatOptionB)
                    writer.WriteAttributeString("IDBottomRail_LinkSys_3rd", boeBottomIdC)
                    writer.WriteAttributeString("FlatBottomOp3b", flatOptionC)
                    writer.WriteAttributeString("ControlPosition", thisData("ControlPosition").ToString())
                    writer.WriteAttributeString("RollDirection", thisData("Roll").ToString())
                    writer.WriteAttributeString("Width", thisData("Width").ToString())
                    writer.WriteAttributeString("Width2b", thisData("WidthB").ToString())
                    writer.WriteAttributeString("Width_LS3rd", thisData("WidthC").ToString())
                    writer.WriteAttributeString("Drop", thisData("Drop").ToString())
                    writer.WriteAttributeString("Drop2c", thisData("DropB").ToString())
                    writer.WriteAttributeString("Drop3c", thisData("DropC").ToString())
                    writer.WriteAttributeString("Panel", thisData("BracketSize").ToString())
                    writer.WriteAttributeString("Notes", thisData("Notes").ToString())
                    writer.WriteEndElement()
                End If

                If blindName = "Link 2 Blinds Independent" Then
                    kitId = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND BlindStatus='Control'")
                    kitIdB = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND BlindStatus='End'")

                    Dim kitName As String = String.Empty

                    If tubeName = "Standard" AndAlso controlName = "Chain" Then
                        kitName = String.Format("{0} (LD)", productName)

                        Dim width As Integer = thisData("Width")
                        Dim widthB As Integer = thisData("WidthB")

                        If width > 1810 OrElse widthB > 1810 Then kitName = String.Format("{0} (HD)", productName)

                        kitId = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='Control'")
                        kitIdB = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='End'")
                    End If

                    If tubeName = "Acmeda 49mm" AndAlso controlName = "Chain" Then
                        kitName = productName
                        If springAssist = "Yes" Then
                            kitName = String.Format("{0} (Spring Assist)", productName)
                        End If
                        kitId = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='Control'")
                        kitIdB = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='End'")
                    End If

                    If String.IsNullOrEmpty(kitId) OrElse String.IsNullOrEmpty(kitIdB) Then Continue While

                    Dim webFabricIdB As String = thisData("FabricColourIdB").ToString()
                    Dim boeFabricIdB As String = GetItemData("SELECT BoeId FROM FabricColours WHERE Id='" & webFabricIdB & "'")

                    Dim chainIdB As String = thisData("ChainIdB").ToString()
                    Dim boeChainIdB As String = GetItemData("SELECT BoeId FROM Chains WHERE Id='" & chainIdB & "'")

                    Dim bottomColourIdB As String = thisData("BottomColourIdB").ToString()
                    Dim boeBottomIdB As String = GetItemData("SELECT BoeId FROM BottomColours WHERE Id='" & bottomColourIdB & "'")

                    Dim flatOption As String = thisData("FlatOption").ToString()
                    Dim flatOptionB As String = thisData("FlatOptionB").ToString()

                    If flatOption = "Fabric on Back" Then flatOption = "Fabric on back"
                    If flatOption = "Fabric on Front" Then flatOption = "Fabric on front"

                    If flatOptionB = "Fabric on Back" Then flatOptionB = "Fabric on back"
                    If flatOptionB = "Fabric on Front" Then flatOptionB = "Fabric on front"

                    writer.WriteStartElement("OrderDetails")
                    writer.WriteAttributeString("OrddID", thisData("Id").ToString())
                    writer.WriteAttributeString("FKOrdID", thisData("HeaderId").ToString())
                    writer.WriteAttributeString("Qty", thisData("Qty").ToString())
                    writer.WriteAttributeString("Room", thisData("Room").ToString())
                    writer.WriteAttributeString("Mounting", thisData("Mounting").ToString())
                    writer.WriteAttributeString("BlindType", "Roller")
                    writer.WriteAttributeString("OrderType", orderType)
                    writer.WriteAttributeString("IDHK", kitId)
                    writer.WriteAttributeString("IDHK2", kitIdB)
                    writer.WriteAttributeString("IDChain", boeChainId)
                    writer.WriteAttributeString("RollerChainLength", thisData("ControlLengthValue").ToString())
                    writer.WriteAttributeString("Additional", thisData("ChainStopper").ToString())
                    If controlName = "Chain" Then
                        writer.WriteAttributeString("IDChain_LinkSysIdp", boeChainIdB)
                        writer.WriteAttributeString("RolChainLength2c", thisData("ControlLengthValueB").ToString())
                        writer.WriteAttributeString("Additional1", thisData("ChainStopperB").ToString())
                    End If
                    writer.WriteAttributeString("FabricID", boeFabricId)
                    writer.WriteAttributeString("FabricID_LinkSysIdp", boeFabricIdB)
                    writer.WriteAttributeString("IDBottomRail", boeBottomId)
                    writer.WriteAttributeString("FlatBottomOp", flatOption)
                    writer.WriteAttributeString("IDBottomRail_LinkSysIdp", boeBottomIdB)
                    writer.WriteAttributeString("FlatBottomOp2c", flatOption)
                    writer.WriteAttributeString("ControlPosition", thisData("ControlPosition").ToString())
                    writer.WriteAttributeString("RollDirection", thisData("Roll").ToString())
                    writer.WriteAttributeString("Width", thisData("Width").ToString())
                    writer.WriteAttributeString("Width2c", thisData("WidthB").ToString())
                    writer.WriteAttributeString("Drop", thisData("Drop").ToString())
                    writer.WriteAttributeString("Drop2c", thisData("DropB").ToString())
                    writer.WriteAttributeString("Panel", thisData("BracketSize").ToString())
                    writer.WriteAttributeString("Notes", thisData("Notes").ToString())
                    writer.WriteEndElement()
                End If

                If blindName = "Link 3 Blinds Independent with Dependent" Then
                    kitId = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND BlindStatus='Control'")
                    kitIdB = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND BlindStatus='Middle'")
                    kitIdC = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND BlindStatus='End'")

                    Dim kitName As String = String.Empty

                    If tubeName = "Standard" AndAlso controlName = "Chain" Then
                        kitName = String.Format("{0} (LD)", productName)

                        Dim width As Integer = thisData("Width")
                        Dim widthB As Integer = thisData("WidthB")
                        Dim widthC As Integer = thisData("WidthC")

                        If width > 1810 OrElse widthB > 1810 OrElse widthC > 1810 Then kitName = String.Format("{0} (HD)", productName)

                        kitId = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='Control'")
                        kitIdB = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='Middle'")
                        kitIdC = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='End'")
                    End If

                    If tubeName = "Acmeda 49mm" AndAlso controlName = "Chain" Then
                        kitName = productName
                        If springAssist = "Yes" Then
                            kitName = String.Format("{0} (Spring Assist)", productName)
                        End If
                        kitId = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='Control'")
                        kitIdB = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='Middle'")
                        kitIdC = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='End'")
                    End If

                    If String.IsNullOrEmpty(kitId) OrElse String.IsNullOrEmpty(kitIdB) Then Continue While

                    Dim webFabricIdB As String = thisData("FabricColourIdB").ToString()
                    Dim webFabricIdC As String = thisData("FabricColourIdC").ToString()

                    Dim boeFabricIdB As String = GetItemData("SELECT BoeId FROM FabricColours WHERE Id='" & webFabricIdB & "'")
                    Dim boeFabricIdC As String = GetItemData("SELECT BoeId FROM FabricColours WHERE Id='" & webFabricIdC & "'")

                    Dim chainIdB As String = thisData("ChainIdB").ToString()
                    Dim chainIdC As String = thisData("ChainIdC").ToString()

                    Dim boeChainIdB As String = GetItemData("SELECT BoeId FROM Chains WHERE Id='" & chainIdB & "'")
                    Dim boeChainIdC As String = GetItemData("SELECT BoeId FROM Chains WHERE Id='" & chainIdC & "'")

                    Dim bottomColourIdB As String = thisData("BottomColourIdB").ToString()
                    Dim bottomColourIdC As String = thisData("BottomColourIdC").ToString()

                    Dim boeBottomIdB As String = GetItemData("SELECT BoeId FROM BottomColours WHERE Id='" & bottomColourIdB & "'")
                    Dim boeBottomIdC As String = GetItemData("SELECT BoeId FROM BottomColours WHERE Id='" & bottomColourIdC & "'")

                    Dim flatOption As String = thisData("FlatOption").ToString()
                    Dim flatOptionB As String = thisData("FlatOptionB").ToString()
                    Dim flatOptionC As String = thisData("FlatOptionC").ToString()

                    If flatOption = "Fabric on Back" Then flatOption = "Fabric on back"
                    If flatOption = "Fabric on Front" Then flatOption = "Fabric on front"

                    If flatOptionB = "Fabric on Back" Then flatOptionB = "Fabric on back"
                    If flatOptionB = "Fabric on Front" Then flatOptionB = "Fabric on front"

                    If flatOptionC = "Fabric on Back" Then flatOptionC = "Fabric on back"
                    If flatOptionC = "Fabric on Front" Then flatOptionC = "Fabric on front"

                    writer.WriteStartElement("OrderDetails")
                    writer.WriteAttributeString("OrddID", thisData("Id").ToString())
                    writer.WriteAttributeString("FKOrdID", thisData("HeaderId").ToString())
                    writer.WriteAttributeString("Qty", thisData("Qty").ToString())
                    writer.WriteAttributeString("Room", thisData("Room").ToString())
                    writer.WriteAttributeString("Mounting", thisData("Mounting").ToString())
                    writer.WriteAttributeString("BlindType", "Roller")
                    writer.WriteAttributeString("OrderType", orderType)
                    writer.WriteAttributeString("IDHK", kitId)
                    writer.WriteAttributeString("IDHK2", kitIdB)
                    writer.WriteAttributeString("IDHK3", kitIdC)
                    writer.WriteAttributeString("IDChain", boeChainId)
                    writer.WriteAttributeString("RollerChainLength", thisData("ControlLengthValue").ToString())
                    writer.WriteAttributeString("Additional", thisData("ChainStopper").ToString())
                    If controlName = "Chain" Then
                        writer.WriteAttributeString("IDChain_LSIdp3rd", boeChainIdC)
                        writer.WriteAttributeString("ChainLength_LSIdp3rd", thisData("ControlLengthValueC").ToString())
                        writer.WriteAttributeString("Additional2", thisData("ChainStopperC").ToString())
                    End If
                    writer.WriteAttributeString("FabricID", boeFabricId)
                    writer.WriteAttributeString("FabricID_LinkSysIdp", boeFabricIdB)
                    writer.WriteAttributeString("FabricID_LSIdp3rd", boeFabricIdC)
                    writer.WriteAttributeString("IDBottomRail", boeBottomId)
                    writer.WriteAttributeString("FlatBottomOp", flatOption)
                    writer.WriteAttributeString("IDBottomRail_LinkSysIdp", boeBottomIdB)
                    writer.WriteAttributeString("FlatBottomOp2c", flatOptionB)
                    writer.WriteAttributeString("IDBottomRail_LinkSysIdp_3rd", boeBottomIdC)
                    writer.WriteAttributeString("FlatBottomOp3c", flatOptionC)
                    writer.WriteAttributeString("ControlPosition", thisData("ControlPosition").ToString())
                    writer.WriteAttributeString("RollDirection", thisData("Roll").ToString())
                    writer.WriteAttributeString("Width", thisData("Width").ToString())
                    writer.WriteAttributeString("Width2c", thisData("WidthB").ToString())
                    writer.WriteAttributeString("Width_LSIdp3rd", thisData("WidthC").ToString())
                    writer.WriteAttributeString("Drop", thisData("Drop").ToString())
                    writer.WriteAttributeString("Drop2c", thisData("DropB").ToString())
                    writer.WriteAttributeString("Drop3c", thisData("DropC").ToString())
                    writer.WriteAttributeString("Panel", thisData("BracketSize").ToString())
                    writer.WriteAttributeString("Notes", thisData("Notes").ToString())
                    writer.WriteEndElement()
                End If

                If blindName = "DB Link 2 Blinds Dependent" Then
                    kitId = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND BlindStatus='Control'")
                    kitIdB = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND BlindStatus='End'")
                    kitIdC = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND BlindStatus='Control'")
                    kitIdD = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND BlindStatus='End'")

                    Dim kitName As String = String.Empty

                    If tubeName = "Standard" AndAlso controlName = "Chain" Then
                        kitName = String.Format("{0} (LD)", productName)

                        Dim width As Integer = thisData("Width")
                        Dim widthB As Integer = thisData("WidthB")

                        If width > 1810 OrElse widthB > 1810 Then kitName = String.Format("{0} (HD)", productName)

                        kitId = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='Control'")
                        kitIdB = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='End'")
                        kitIdC = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='Control'")
                        kitIdD = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='End'")
                    End If

                    If tubeName = "Acmeda 49mm" AndAlso controlName = "Chain" Then
                        kitName = productName
                        If springAssist = "Yes" Then
                            kitName = String.Format("{0} (Spring Assist)", productName)
                        End If
                        kitId = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='Control'")
                        kitIdB = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='End'")
                        kitIdC = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='Control'")
                        kitIdD = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='End'")
                    End If

                    If String.IsNullOrEmpty(kitId) OrElse String.IsNullOrEmpty(kitIdB) OrElse String.IsNullOrEmpty(kitIdC) OrElse String.IsNullOrEmpty(kitIdD) Then Continue While

                    Dim webFabricIdB As String = thisData("FabricColourIdB").ToString()
                    Dim webFabricIdC As String = thisData("FabricColourIdC").ToString()
                    Dim webFabricIdD As String = thisData("FabricColourIdD").ToString()

                    Dim boeFabricIdB As String = GetItemData("SELECT BoeId FROM FabricColours WHERE Id='" & webFabricIdB & "'")
                    Dim boeFabricIdC As String = GetItemData("SELECT BoeId FROM FabricColours WHERE Id='" & webFabricIdC & "'")
                    Dim boeFabricIdD As String = GetItemData("SELECT BoeId FROM FabricColours WHERE Id='" & webFabricIdD & "'")

                    Dim chainIdC As String = thisData("ChainIdC").ToString()
                    Dim boeChainIdC As String = GetItemData("SELECT BoeId FROM Chains WHERE Id='" & chainIdC & "'")

                    Dim bottomColourIdB As String = thisData("BottomColourIdB").ToString()
                    Dim bottomColourIdC As String = thisData("BottomColourIdC").ToString()
                    Dim bottomColourIdD As String = thisData("BottomColourIdD").ToString()

                    Dim boeBottomIdB As String = GetItemData("SELECT BoeId FROM BottomColours WHERE Id='" & bottomColourIdB & "'")
                    Dim boeBottomIdC As String = GetItemData("SELECT BoeId FROM BottomColours WHERE Id='" & bottomColourIdC & "'")
                    Dim boeBottomIdD As String = GetItemData("SELECT BoeId FROM BottomColours WHERE Id='" & bottomColourIdD & "'")

                    Dim flatOption As String = thisData("FlatOption").ToString()
                    Dim flatOptionB As String = thisData("FlatOptionB").ToString()
                    Dim flatOptionC As String = thisData("FlatOptionC").ToString()
                    Dim flatOptionD As String = thisData("FlatOptionD").ToString()

                    If flatOption = "Fabric on Back" Then flatOption = "Fabric on back"
                    If flatOption = "Fabric on Front" Then flatOption = "Fabric on front"

                    If flatOptionB = "Fabric on Back" Then flatOptionB = "Fabric on back"
                    If flatOptionB = "Fabric on Front" Then flatOptionB = "Fabric on front"

                    If flatOptionC = "Fabric on Back" Then flatOptionC = "Fabric on back"
                    If flatOptionC = "Fabric on Front" Then flatOptionC = "Fabric on front"

                    If flatOptionD = "Fabric on Back" Then flatOptionD = "Fabric on back"
                    If flatOptionD = "Fabric on Front" Then flatOptionD = "Fabric on front"

                    writer.WriteStartElement("OrderDetails")
                    writer.WriteAttributeString("OrddID", thisData("Id").ToString())
                    writer.WriteAttributeString("FKOrdID", thisData("HeaderId").ToString())
                    writer.WriteAttributeString("Qty", thisData("Qty").ToString())
                    writer.WriteAttributeString("Room", thisData("Room").ToString())
                    writer.WriteAttributeString("Mounting", thisData("Mounting").ToString())
                    writer.WriteAttributeString("BlindType", "Roller")
                    writer.WriteAttributeString("OrderType", orderType)
                    writer.WriteAttributeString("IDHK", kitId)
                    writer.WriteAttributeString("IDHK2", kitIdB)
                    writer.WriteAttributeString("IDHK3", kitIdC)
                    writer.WriteAttributeString("IDHK4", kitIdD)
                    writer.WriteAttributeString("NumOfBlind", thisData("TotalItems").ToString())

                    writer.WriteAttributeString("IDChain", boeChainId)
                    writer.WriteAttributeString("IDChain_LSDouble3rd", boeChainIdC)

                    writer.WriteAttributeString("RollerChainLength", thisData("ControlLengthValue").ToString())
                    writer.WriteAttributeString("Chain_LSDouble3rd", thisData("ControlLengthValueC").ToString())

                    writer.WriteAttributeString("Additional", thisData("ChainStopper").ToString())
                    writer.WriteAttributeString("Additional1", thisData("ChainStopperC").ToString())

                    writer.WriteAttributeString("FabricID", boeFabricId)
                    writer.WriteAttributeString("FabricID_LinkSys", boeFabricIdB)
                    writer.WriteAttributeString("FabricID_LSDouble3rd", boeFabricIdC)
                    writer.WriteAttributeString("FabricID_LSDouble4th", boeFabricIdD)
                    writer.WriteAttributeString("IDBottomRail", boeBottomId)
                    writer.WriteAttributeString("IDBottomRail_LinkSys", boeBottomIdB)
                    writer.WriteAttributeString("IDBottomRail_LSDouble3rd", boeBottomIdC)
                    writer.WriteAttributeString("IDBottomRail_LSDouble4th", boeBottomIdD)
                    writer.WriteAttributeString("FlatBottomOp", flatOption)
                    writer.WriteAttributeString("FlatBottomOp2b", flatOption)
                    writer.WriteAttributeString("FlatBottomOpDB3b", flatOptionC)
                    writer.WriteAttributeString("FlatBottomOp4b", flatOptionD)
                    writer.WriteAttributeString("ControlPosition", thisData("ControlPosition").ToString())
                    writer.WriteAttributeString("ControlPosition_LSDouble3rd", thisData("ControlPositionC").ToString())
                    writer.WriteAttributeString("RollDirection", thisData("Roll").ToString())
                    writer.WriteAttributeString("Roll_LSDouble3rd", thisData("RollC").ToString())
                    writer.WriteAttributeString("Width", thisData("Width").ToString())
                    writer.WriteAttributeString("Width2b", thisData("WidthB").ToString())
                    writer.WriteAttributeString("Drop", thisData("Drop").ToString())
                    writer.WriteAttributeString("Panel", thisData("BracketSize").ToString())
                    writer.WriteAttributeString("Notes", thisData("Notes").ToString())
                    writer.WriteEndElement()
                End If

                If blindName = "DB Link 2 Blinds Independent" Then
                    kitId = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND BlindStatus='Control'")
                    kitIdB = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND BlindStatus='End'")
                    kitIdC = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND BlindStatus='Control'")
                    kitIdD = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND BlindStatus='End'")

                    Dim kitName As String = String.Empty

                    If tubeName = "Standard" AndAlso controlName = "Chain" Then
                        kitName = String.Format("{0} (LD)", productName)

                        Dim width As Integer = thisData("Width")
                        Dim widthB As Integer = thisData("WidthB")

                        If width > 1810 OrElse widthB > 1810 Then kitName = String.Format("{0} (HD)", productName)

                        kitId = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='Control'")
                        kitIdB = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='End'")
                        kitIdC = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='Control'")
                        kitIdD = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='End'")
                    End If

                    If tubeName = "Acmeda 49mm" AndAlso controlName = "Chain" Then
                        kitName = productName
                        If springAssist = "Yes" Then
                            kitName = String.Format("{0} (Spring Assist)", productName)
                        End If
                        kitId = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='Control'")
                        kitIdB = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='End'")
                        kitIdC = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='Control'")
                        kitIdD = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='End'")
                    End If

                    If String.IsNullOrEmpty(kitId) OrElse String.IsNullOrEmpty(kitIdB) OrElse String.IsNullOrEmpty(kitIdC) OrElse String.IsNullOrEmpty(kitIdD) Then Continue While

                    Dim webFabricIdB As String = thisData("FabricColourIdB").ToString()
                    Dim webFabricIdC As String = thisData("FabricColourIdC").ToString()
                    Dim webFabricIdD As String = thisData("FabricColourIdD").ToString()

                    Dim boeFabricIdB As String = GetItemData("SELECT BoeId FROM FabricColours WHERE Id='" & webFabricIdB & "'")
                    Dim boeFabricIdC As String = GetItemData("SELECT BoeId FROM FabricColours WHERE Id='" & webFabricIdC & "'")
                    Dim boeFabricIdD As String = GetItemData("SELECT BoeId FROM FabricColours WHERE Id='" & webFabricIdD & "'")

                    Dim chainIdB As String = thisData("ChainIdB").ToString()
                    Dim chainIdC As String = thisData("ChainIdC").ToString()
                    Dim chainIdD As String = thisData("ChainIdD").ToString()

                    Dim boeChainIdB As String = GetItemData("SELECT BoeId FROM Chains WHERE Id='" & chainIdB & "'")
                    Dim boeChainIdC As String = GetItemData("SELECT BoeId FROM Chains WHERE Id='" & chainIdC & "'")
                    Dim boeChainIdD As String = GetItemData("SELECT BoeId FROM Chains WHERE Id='" & chainIdD & "'")

                    Dim bottomColourIdB As String = thisData("BottomColourIdB").ToString()
                    Dim bottomColourIdC As String = thisData("BottomColourIdC").ToString()
                    Dim bottomColourIdD As String = thisData("BottomColourIdD").ToString()

                    Dim boeBottomIdB As String = GetItemData("SELECT BoeId FROM BottomColours WHERE Id='" & bottomColourIdB & "'")
                    Dim boeBottomIdC As String = GetItemData("SELECT BoeId FROM BottomColours WHERE Id='" & bottomColourIdC & "'")
                    Dim boeBottomIdD As String = GetItemData("SELECT BoeId FROM BottomColours WHERE Id='" & bottomColourIdD & "'")

                    Dim flatOption As String = thisData("FlatOption").ToString()
                    Dim flatOptionB As String = thisData("FlatOptionB").ToString()
                    Dim flatOptionC As String = thisData("FlatOptionC").ToString()
                    Dim flatOptionD As String = thisData("FlatOptionD").ToString()

                    If flatOption = "Fabric on Back" Then flatOption = "Fabric on back"
                    If flatOption = "Fabric on Front" Then flatOption = "Fabric on front"

                    If flatOptionB = "Fabric on Back" Then flatOptionB = "Fabric on back"
                    If flatOptionB = "Fabric on Front" Then flatOptionB = "Fabric on front"

                    If flatOptionC = "Fabric on Back" Then flatOptionC = "Fabric on back"
                    If flatOptionC = "Fabric on Front" Then flatOptionC = "Fabric on front"

                    If flatOptionD = "Fabric on Back" Then flatOptionD = "Fabric on back"
                    If flatOptionD = "Fabric on Front" Then flatOptionD = "Fabric on front"

                    writer.WriteStartElement("OrderDetails")
                    writer.WriteAttributeString("OrddID", thisData("Id").ToString())
                    writer.WriteAttributeString("FKOrdID", thisData("HeaderId").ToString())
                    writer.WriteAttributeString("Qty", thisData("Qty").ToString())
                    writer.WriteAttributeString("Room", thisData("Room").ToString())
                    writer.WriteAttributeString("Mounting", thisData("Mounting").ToString())
                    writer.WriteAttributeString("BlindType", "Roller")
                    writer.WriteAttributeString("OrderType", orderType)
                    writer.WriteAttributeString("IDHK", kitId)
                    writer.WriteAttributeString("IDHK2", kitIdB)
                    writer.WriteAttributeString("IDHK3", kitIdC)
                    writer.WriteAttributeString("IDHK4", kitIdD)
                    writer.WriteAttributeString("NumOfBlind", thisData("TotalItems").ToString())

                    writer.WriteAttributeString("IDChain", boeChainId)
                    writer.WriteAttributeString("IDChain_LinkSysIdp", boeChainIdB)
                    writer.WriteAttributeString("IDChain_LSIdpDouble3rd", boeChainIdC)
                    writer.WriteAttributeString("IDChain_LSIdpDouble4th", boeChainIdD)

                    writer.WriteAttributeString("RollerChainLength", thisData("ControlLengthValue").ToString())
                    writer.WriteAttributeString("RolChainLength2c", thisData("ControlLengthValueB").ToString())
                    writer.WriteAttributeString("Chain_LSIdpDouble3rd", thisData("ControlLengthValueC").ToString())
                    writer.WriteAttributeString("Chain_LSIdpDouble4th", thisData("ControlLengthValueD").ToString())

                    writer.WriteAttributeString("Additional", thisData("ChainStopper").ToString())
                    writer.WriteAttributeString("Additional1", thisData("ChainStopperB").ToString())
                    writer.WriteAttributeString("Additional2", thisData("ChainStopperC").ToString())
                    writer.WriteAttributeString("Additional3", thisData("ChainStopperD").ToString())

                    writer.WriteAttributeString("FabricID", boeFabricId)
                    writer.WriteAttributeString("FabricID_LinkSysIdp", boeFabricIdB)
                    writer.WriteAttributeString("FabricID_LSIdpDouble3rd", boeFabricIdC)
                    writer.WriteAttributeString("FabricID_LSIdpDouble4th", boeFabricIdD)
                    writer.WriteAttributeString("IDBottomRail", boeBottomId)
                    writer.WriteAttributeString("IDBottomRail_LinkSysIdp", boeBottomIdB)
                    writer.WriteAttributeString("IDBottomRail_LSIdpDouble3rd", boeBottomIdC)
                    writer.WriteAttributeString("IDBottomRail_LSIdpDouble4th", boeBottomIdD)
                    writer.WriteAttributeString("FlatBottomOp", flatOption)
                    writer.WriteAttributeString("FlatBottomOp2c", flatOption)
                    writer.WriteAttributeString("FlatBottomOpDB3c", flatOptionC)
                    writer.WriteAttributeString("FlatBottomOp4c", flatOptionD)
                    writer.WriteAttributeString("ControlPosition", thisData("ControlPosition").ToString())
                    writer.WriteAttributeString("ControlPosition_LSDouble3rd", thisData("ControlPositionC").ToString())
                    writer.WriteAttributeString("RollDirection", thisData("Roll").ToString())
                    writer.WriteAttributeString("Roll_LSDouble3rd", thisData("RollC").ToString())
                    writer.WriteAttributeString("Width", thisData("Width").ToString())
                    writer.WriteAttributeString("Width2c", thisData("WidthB").ToString())
                    writer.WriteAttributeString("Drop", thisData("Drop").ToString())
                    writer.WriteAttributeString("Panel", thisData("BracketSize").ToString())
                    writer.WriteAttributeString("Notes", thisData("Notes").ToString())
                    writer.WriteEndElement()
                End If

                If blindName = "DB Link 3 Blinds Independent with Dependent" Then
                    kitId = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND BlindStatus='Control'")
                    kitIdB = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND BlindStatus='Middle'")
                    kitIdC = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND BlindStatus='End'")
                    kitIdD = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND BlindStatus='Control'")
                    kitIdE = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND BlindStatus='Middle'")
                    kitIdF = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND BlindStatus='End'")

                    Dim kitName As String = String.Empty

                    If tubeName = "Standard" AndAlso controlName = "Chain" Then
                        kitName = String.Format("{0} (LD)", productName)

                        Dim width As Integer = thisData("Width")
                        Dim widthB As Integer = thisData("WidthB")
                        Dim widthC As Integer = thisData("WidthC")

                        If width > 1810 OrElse widthB > 1810 Then kitName = String.Format("{0} (HD)", productName)

                        kitId = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='Control'")
                        kitIdB = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='Middle'")
                        kitIdC = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='End'")
                        kitIdD = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='Control'")
                        kitIdE = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='Middle'")
                        kitIdF = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='End'")
                    End If

                    If tubeName = "Acmeda 49mm" AndAlso controlName = "Chain" Then
                        kitName = productName
                        If springAssist = "Yes" Then
                            kitName = String.Format("{0} (Spring Assist)", productName)
                        End If
                        kitId = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='Control'")
                        kitIdB = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='Middle'")
                        kitIdC = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='End'")
                        kitIdD = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='Control'")
                        kitIdE = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='Middle'")
                        kitIdF = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "' AND Name='" & kitName & "' AND BlindStatus='End'")
                    End If

                    If String.IsNullOrEmpty(kitId) OrElse String.IsNullOrEmpty(kitIdB) OrElse String.IsNullOrEmpty(kitIdC) OrElse String.IsNullOrEmpty(kitIdD) OrElse String.IsNullOrEmpty(kitIdE) OrElse String.IsNullOrEmpty(kitIdF) Then Continue While

                    Dim webFabricIdB As String = thisData("FabricColourIdB").ToString()
                    Dim webFabricIdC As String = thisData("FabricColourIdC").ToString()
                    Dim webFabricIdD As String = thisData("FabricColourIdD").ToString()
                    Dim webFabricIdE As String = thisData("FabricColourIdE").ToString()
                    Dim webFabricIdF As String = thisData("FabricColourIdF").ToString()

                    Dim boeFabricIdB As String = GetItemData("SELECT BoeId FROM FabricColours WHERE Id='" & webFabricIdB & "'")
                    Dim boeFabricIdC As String = GetItemData("SELECT BoeId FROM FabricColours WHERE Id='" & webFabricIdC & "'")
                    Dim boeFabricIdD As String = GetItemData("SELECT BoeId FROM FabricColours WHERE Id='" & webFabricIdD & "'")
                    Dim boeFabricIdE As String = GetItemData("SELECT BoeId FROM FabricColours WHERE Id='" & webFabricIdE & "'")
                    Dim boeFabricIdF As String = GetItemData("SELECT BoeId FROM FabricColours WHERE Id='" & webFabricIdF & "'")

                    Dim chainIdC As String = thisData("ChainIdC").ToString()
                    Dim chainIdD As String = thisData("ChainIdD").ToString()
                    Dim chainIdF As String = thisData("ChainIdF").ToString()

                    Dim boeChainIdC As String = GetItemData("SELECT BoeId FROM Chains WHERE Id='" & chainIdC & "'")
                    Dim boeChainIdD As String = GetItemData("SELECT BoeId FROM Chains WHERE Id='" & chainIdD & "'")
                    Dim boeChainIdF As String = GetItemData("SELECT BoeId FROM Chains WHERE Id='" & chainIdF & "'")

                    Dim bottomColourIdB As String = thisData("BottomColourIdB").ToString()
                    Dim bottomColourIdC As String = thisData("BottomColourIdC").ToString()
                    Dim bottomColourIdD As String = thisData("BottomColourIdD").ToString()
                    Dim bottomColourIdE As String = thisData("BottomColourIdE").ToString()
                    Dim bottomColourIdF As String = thisData("BottomColourIdF").ToString()

                    Dim boeBottomIdB As String = GetItemData("SELECT BoeId FROM BottomColours WHERE Id='" & bottomColourIdB & "'")
                    Dim boeBottomIdC As String = GetItemData("SELECT BoeId FROM BottomColours WHERE Id='" & bottomColourIdC & "'")
                    Dim boeBottomIdD As String = GetItemData("SELECT BoeId FROM BottomColours WHERE Id='" & bottomColourIdD & "'")
                    Dim boeBottomIdE As String = GetItemData("SELECT BoeId FROM BottomColours WHERE Id='" & bottomColourIdE & "'")
                    Dim boeBottomIdF As String = GetItemData("SELECT BoeId FROM BottomColours WHERE Id='" & bottomColourIdF & "'")

                    Dim flatOption As String = thisData("FlatOption").ToString()
                    Dim flatOptionB As String = thisData("FlatOptionB").ToString()
                    Dim flatOptionC As String = thisData("FlatOptionC").ToString()
                    Dim flatOptionD As String = thisData("FlatOptionD").ToString()
                    Dim flatOptionE As String = thisData("FlatOptionE").ToString()
                    Dim flatOptionF As String = thisData("FlatOptionF").ToString()

                    If flatOption = "Fabric on Back" Then flatOption = "Fabric on back"
                    If flatOption = "Fabric on Front" Then flatOption = "Fabric on front"

                    If flatOptionB = "Fabric on Back" Then flatOptionB = "Fabric on back"
                    If flatOptionB = "Fabric on Front" Then flatOptionB = "Fabric on front"

                    If flatOptionC = "Fabric on Back" Then flatOptionC = "Fabric on back"
                    If flatOptionC = "Fabric on Front" Then flatOptionC = "Fabric on front"

                    If flatOptionD = "Fabric on Back" Then flatOptionD = "Fabric on back"
                    If flatOptionD = "Fabric on Front" Then flatOptionD = "Fabric on front"

                    If flatOptionE = "Fabric on Back" Then flatOptionE = "Fabric on back"
                    If flatOptionE = "Fabric on Front" Then flatOptionE = "Fabric on front"

                    If flatOptionF = "Fabric on Back" Then flatOptionF = "Fabric on back"
                    If flatOptionF = "Fabric on Front" Then flatOptionF = "Fabric on front"

                    writer.WriteStartElement("OrderDetails")
                    writer.WriteAttributeString("OrddID", thisData("Id").ToString())
                    writer.WriteAttributeString("FKOrdID", thisData("HeaderId").ToString())
                    writer.WriteAttributeString("Qty", thisData("Qty").ToString())
                    writer.WriteAttributeString("Room", thisData("Room").ToString())
                    writer.WriteAttributeString("Mounting", thisData("Mounting").ToString())
                    writer.WriteAttributeString("BlindType", "Roller")
                    writer.WriteAttributeString("OrderType", orderType)
                    writer.WriteAttributeString("IDHK", kitId)
                    writer.WriteAttributeString("IDHK2", kitIdB)
                    writer.WriteAttributeString("IDHK3", kitIdC)
                    writer.WriteAttributeString("IDHK4", kitIdD)
                    writer.WriteAttributeString("IDHK5", kitIdE)
                    writer.WriteAttributeString("IDHK6", kitIdF)
                    writer.WriteAttributeString("NumOfBlind", thisData("TotalItems").ToString())

                    writer.WriteAttributeString("IDChain", boeChainId)
                    writer.WriteAttributeString("IDChain_LSIdp3rd", boeChainIdC)
                    writer.WriteAttributeString("IDChainDBL", boeChainIdD)
                    writer.WriteAttributeString("IDChainDBR", boeChainIdF)

                    writer.WriteAttributeString("RollerChainLength", thisData("ControlLengthValue").ToString())
                    writer.WriteAttributeString("ChainLength_LSIdp3rd", thisData("ControlLengthValueC").ToString())
                    writer.WriteAttributeString("ChainLengthdb4", thisData("ControlLengthValueD").ToString())
                    writer.WriteAttributeString("ChainLengthdb6", thisData("ControlLengthValueF").ToString())

                    writer.WriteAttributeString("Additional", thisData("ChainStopper").ToString())
                    writer.WriteAttributeString("Additional1", thisData("ChainStopperC").ToString())
                    writer.WriteAttributeString("chainst4", thisData("ChainStopperD").ToString())
                    writer.WriteAttributeString("chainst6", thisData("ChainStopperE").ToString())

                    writer.WriteAttributeString("FabricID", boeFabricId)
                    writer.WriteAttributeString("FabricID_LinkSysIdp", boeFabricIdB)
                    writer.WriteAttributeString("FabricID_LSIdp3rd", boeFabricIdC)
                    writer.WriteAttributeString("FabricID_LSDouble4th", boeFabricIdD)
                    writer.WriteAttributeString("FabricID5", boeFabricIdE)
                    writer.WriteAttributeString("FabricID6", boeFabricIdF)

                    writer.WriteAttributeString("IDBottomRail", boeBottomId)
                    writer.WriteAttributeString("IDBottomRail_LinkSysIdp", boeBottomIdB)
                    writer.WriteAttributeString("IDBottomRail_LinkSysIdp_3rd", boeBottomIdC)
                    writer.WriteAttributeString("IDBottomRail_LSDouble4th", boeBottomIdD)
                    writer.WriteAttributeString("IDBottomRail5", boeBottomIdE)
                    writer.WriteAttributeString("IDBottomRail6", boeBottomIdF)

                    writer.WriteAttributeString("FlatBottomOp", flatOption)
                    writer.WriteAttributeString("FlatBottomOp2c", flatOption)
                    writer.WriteAttributeString("FlatBottomOp3c", flatOptionC)
                    writer.WriteAttributeString("FlatBottomdb4", flatOptionD)
                    writer.WriteAttributeString("FlatBottomop5", flatOptionE)
                    writer.WriteAttributeString("FlatBottomop6", flatOptionF)

                    writer.WriteAttributeString("ControlPosition", thisData("ControlPosition").ToString())
                    writer.WriteAttributeString("ControlPosition_LSDouble3rd", thisData("ControlPositionD").ToString())
                    writer.WriteAttributeString("RollDirection", thisData("Roll").ToString())
                    writer.WriteAttributeString("Roll_LSDouble3rd", thisData("RollD").ToString())
                    writer.WriteAttributeString("Width", thisData("Width").ToString())
                    writer.WriteAttributeString("Width2c", thisData("WidthB").ToString())
                    writer.WriteAttributeString("Width_LSIdp3rd", thisData("WidthC").ToString())
                    writer.WriteAttributeString("Drop", thisData("Drop").ToString())
                    writer.WriteAttributeString("Panel", thisData("BracketSize").ToString())
                    writer.WriteAttributeString("Notes", thisData("Notes").ToString())
                    writer.WriteEndElement()
                End If
            End If

            If designName = "Outdoor" Then
                Dim kitId As String = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "'")

                Dim webFabricId As String = thisData("FabricColourId").ToString()
                Dim boeFabricId As String = GetItemData("SELECT BoeId FROM FabricColours WHERE Id='" & webFabricId & "'")

                If String.IsNullOrEmpty(kitId) Then Continue While
                If String.IsNullOrEmpty(boeFabricId) Then Continue While

                writer.WriteStartElement("OrderDetails")
                writer.WriteAttributeString("OrddID", thisData("Id").ToString())
                writer.WriteAttributeString("FKOrdID", thisData("HeaderId").ToString())
                writer.WriteAttributeString("Qty", thisData("Qty").ToString())
                writer.WriteAttributeString("Room", thisData("Room").ToString())
                writer.WriteAttributeString("Mounting", thisData("Mounting").ToString())

                writer.WriteAttributeString("BlindType", designName)
                writer.WriteAttributeString("OrderType", blindName)
                writer.WriteAttributeString("IDHK", kitId)
                writer.WriteAttributeString("Width", thisData("Width").ToString())
                writer.WriteAttributeString("Drop", thisData("Drop").ToString())
                writer.WriteAttributeString("ControlType", controlName)
                writer.WriteAttributeString("ControlPosition", thisData("ControlPosition").ToString())
                writer.WriteAttributeString("ControlLength", thisData("ControlLength").ToString())
                writer.WriteAttributeString("FabricID", boeFabricId)

                writer.WriteAttributeString("TotalItems", thisData("TotalItems").ToString())
                writer.WriteAttributeString("MarkUp", thisData("MarkUp").ToString())
                writer.WriteAttributeString("Notes", thisData("Notes").ToString())
                writer.WriteEndElement()
            End If

            If designName = "Skyline Shutter Express" Then
                writer.WriteStartElement("OrderDetails")
                writer.WriteAttributeString("OrddID", thisData("Id").ToString())
                writer.WriteAttributeString("FKOrdID", thisData("HeaderId").ToString())
                writer.WriteAttributeString("BlindType", "Shutter")
                writer.WriteAttributeString("Qty", thisData("Qty").ToString())
                writer.WriteAttributeString("Room", thisData("Room").ToString())
                writer.WriteAttributeString("Mounting", thisData("Mounting").ToString())
                writer.WriteAttributeString("Width", thisData("Width").ToString())
                writer.WriteAttributeString("Drop", thisData("Drop").ToString())
                writer.WriteAttributeString("TotalItems", thisData("TotalItems").ToString())
                writer.WriteAttributeString("MarkUp", thisData("MarkUp").ToString())
                writer.WriteAttributeString("Notes", thisData("Notes").ToString())
                writer.WriteEndElement()
            End If

            If designName = "Sample" Then
                Dim kitId As String = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "'")

                If String.IsNullOrEmpty(kitId) Then Continue While
            End If

            If designName = "Roller Horizon" Then
                Dim webFabricId As String = thisData("FabricColourId").ToString()
                Dim boeFabricId As String = GetItemData("SELECT BoeId FROM FabricColours WHERE Id='" & webFabricId & "'")

                Dim chainId As String = thisData("ChainId").ToString()
                Dim boeChainId As String = GetItemData("SELECT BoeId FROM Chains WHERE Id='" & chainId & "'")

                Dim bottomColourId As String = thisData("BottomColourId").ToString()
                Dim boeBottomId As String = GetItemData("SELECT BoeId FROM BottomColours WHERE Id='" & bottomColourId & "'")

                Dim kitId As String = GetItemData("SELECT KitId FROM ProductKits WHERE ProductId='" & productId & "'")

                If String.IsNullOrEmpty(kitId) Then Continue While
                If String.IsNullOrEmpty(boeFabricId) Then Continue While

                writer.WriteStartElement("OrderDetails")
                writer.WriteAttributeString("OrddID", thisData("Id").ToString())
                writer.WriteAttributeString("FKOrdID", thisData("HeaderId").ToString())
                writer.WriteAttributeString("Qty", thisData("Qty").ToString())
                writer.WriteAttributeString("Room", thisData("Room").ToString())
                writer.WriteAttributeString("Mounting", thisData("Mounting").ToString())
                writer.WriteAttributeString("BlindType", "Roller")
                writer.WriteAttributeString("OrderType", "Regular Chain")
                writer.WriteAttributeString("IDHK", kitId)
                writer.WriteAttributeString("IDChain", boeChainId)
                writer.WriteAttributeString("RollerChainLength", thisData("ControlLengthValue").ToString())
                writer.WriteAttributeString("Additional", thisData("ChainStopper").ToString())
                writer.WriteAttributeString("FabricID", boeFabricId)
                writer.WriteAttributeString("IDBottomRail", boeBottomId)
                writer.WriteAttributeString("FlatBottomOp", String.Empty)
                writer.WriteAttributeString("ControlPosition", thisData("ControlPosition").ToString())
                writer.WriteAttributeString("RollDirection", thisData("Roll").ToString())
                writer.WriteAttributeString("Width", thisData("Width").ToString())
                writer.WriteAttributeString("Drop", thisData("Drop").ToString())
                writer.WriteAttributeString("Notes", thisData("Notes").ToString())
                writer.WriteEndElement()
            End If

            If designName = "Door" OrElse designName = "Window" Then
                Dim frameColour As String = thisData("FrameColour").ToString()
                If frameColour.Contains("Regular") Then Continue While

                writer.WriteStartElement("OrderDetails")
                writer.WriteAttributeString("OrddID", thisData("Id").ToString())
                writer.WriteAttributeString("FKOrdID", thisData("HeaderId").ToString())
                writer.WriteAttributeString("BlindType", designName)
                writer.WriteAttributeString("OrderType", blindName)
                writer.WriteAttributeString("Qty", thisData("Qty").ToString())
                writer.WriteAttributeString("Room", thisData("Room").ToString())
                writer.WriteAttributeString("Mounting", thisData("Mounting").ToString())
                writer.WriteAttributeString("Width", thisData("Width").ToString())
                writer.WriteAttributeString("Drop", thisData("Drop").ToString())
                writer.WriteAttributeString("TotalItems", thisData("TotalItems").ToString())
                writer.WriteAttributeString("MarkUp", thisData("MarkUp").ToString())
                writer.WriteAttributeString("Notes", thisData("Notes").ToString())
                writer.WriteEndElement()
            End If
        End While

        writer.WriteEndElement()
        writer.WriteEndDocument()
        writer.Close()
    End Sub
</script>
