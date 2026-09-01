Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Net
Imports System.Net.Mail

Public Class ProcessClass

    Dim myConn As String = ConfigurationManager.ConnectionStrings("DefaultConnection").ConnectionString

    Public Function GetDataRow(thisString As String) As DataRow
        Try
            Using thisConn As New SqlConnection(myConn)
                Using thisCmd As New SqlCommand(thisString, thisConn)
                    Using thisAdapter As New SqlDataAdapter(thisCmd)
                        Dim thisTable As New DataTable()
                        thisAdapter.Fill(thisTable)
                        If thisTable.Rows.Count > 0 Then
                            Return thisTable.Rows(0)
                        Else
                            Return Nothing
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Function GetDataRowSP(spName As String, params As List(Of SqlParameter)) As DataRow
        Try
            Using thisConn As New SqlConnection(myConn)
                Using thisCmd As New SqlCommand(spName, thisConn)
                    thisCmd.CommandType = CommandType.StoredProcedure
                    thisCmd.Parameters.AddRange(params.ToArray())
                    Using thisAdapter As New SqlDataAdapter(thisCmd)
                        Dim thisTable As New DataTable()
                        thisAdapter.Fill(thisTable)
                        If thisTable.Rows.Count > 0 Then
                            Return thisTable.Rows(0)
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
        End Try
        Return Nothing
    End Function

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

    Public Function GetDataTableSP(spName As String, params As List(Of SqlParameter)) As DataTable
        Dim thisTable As New DataTable()
        Try
            Using thisConn As New SqlConnection(myConn)
                Using thisCmd As New SqlCommand(spName, thisConn)
                    thisCmd.CommandType = CommandType.StoredProcedure
                    If params IsNot Nothing Then
                        If params.Count > 0 Then
                            thisCmd.Parameters.AddRange(params.ToArray())
                        End If
                    End If
                    Using thisAdapter As New SqlDataAdapter(thisCmd)
                        thisAdapter.Fill(thisTable)
                    End Using
                End Using
            End Using
        Catch ex As Exception
            thisTable = New DataTable()
        End Try
        Return thisTable
    End Function

    Public Function GetItemData(thisString As String) As String
        Dim result As String = String.Empty
        Try
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
        Catch ex As Exception
            result = String.Empty
        End Try
        Return result
    End Function

    Public Function GetItemData_Integer(thisString As String) As Integer
        Dim result As Integer = 0
        Try
            Using thisConn As New SqlConnection(myConn)
                thisConn.Open()
                Using thisCmd As New SqlCommand(thisString, thisConn)
                    Using rdResult = thisCmd.ExecuteReader
                        While rdResult.Read
                            result = rdResult.Item(0)
                        End While
                    End Using
                End Using
                thisConn.Close()
            End Using
        Catch ex As Exception
            result = 0
        End Try
        Return result
    End Function

    Public Function CreateId(thisString As String) As String
        Dim result As String = String.Empty
        Try
            Dim id As Integer = 0
            Using thisConn As New SqlConnection(myConn)
                thisConn.Open()
                Using thisCmd As New SqlCommand(thisString, thisConn)
                    Using rdResult As SqlDataReader = thisCmd.ExecuteReader()
                        If rdResult.Read() Then
                            Integer.TryParse(rdResult(0).ToString(), id)
                        End If
                    End Using
                End Using
            End Using
            result = (id + 1).ToString()
        Catch ex As Exception
            result = String.Empty
        End Try
        Return result
    End Function

    Public Sub Logs(data As Object())
        Try
            If data.Length = 4 Then
                Using thisConn As New SqlConnection(myConn)
                    Using thisCmd As New SqlCommand("sp_Logs_Insert", thisConn)
                        thisCmd.CommandType = CommandType.StoredProcedure
                        thisCmd.Parameters.AddWithValue("@Type", Convert.ToString(data(0)))
                        thisCmd.Parameters.AddWithValue("@DataId", Convert.ToString(data(1)))
                        thisCmd.Parameters.AddWithValue("@ActionBy", Convert.ToString(data(2)))
                        thisCmd.Parameters.AddWithValue("@Description", Convert.ToString(data(3)))
                        thisConn.Open()
                        thisCmd.ExecuteNonQuery()
                    End Using
                End Using
            End If
        Catch ex As Exception
        End Try
    End Sub

    Public Sub RefreshSalesData(companyId As String)
        Try
            If Not String.IsNullOrEmpty(companyId) Then
                Using thisConn As New SqlConnection(myConn)
                    Using thisCmd As New SqlCommand("sp_Sales_Refresh", thisConn)
                        thisCmd.CommandType = CommandType.StoredProcedure
                        thisCmd.Parameters.AddWithValue("@CompanyId", companyId)
                        thisConn.Open()
                        thisCmd.ExecuteNonQuery()
                    End Using
                End Using
            End If
        Catch ex As Exception
        End Try
    End Sub

    Public Sub ResetProformaOrder(headerId As String)
        Try
            If String.IsNullOrEmpty(headerId) Then Exit Sub

            Dim orderData As DataRow = GetDataRow("SELECT OrderHeaders.*, Customers.Name AS CustomerName, Customers.CompanyId AS CompanyId, Customers.Operator AS Operator, OrderHeaders.InvoiceNumber AS InvoiceNumber FROM OrderHeaders LEFT JOIN Customers ON OrderHeaders.CustomerId=Customers.Id WHERE OrderHeaders.Id='" & headerId & "'")
            If orderData Is Nothing Then Exit Sub

            Dim customerId As String = orderData("CustomerId").ToString()
            Dim orderId As String = orderData("OrderId").ToString()
            Dim orderNumber As String = orderData("OrderNumber").ToString()
            Dim orderName As String = orderData("OrderName").ToString()
            Dim invoiceNumber As String = orderData("InvoiceNumber").ToString()

            Dim customerName As String = orderData("CustomerName").ToString()

            Dim companyId As String = orderData("CompanyId").ToString()
            Dim companyName As String = GetItemData("SELECT Name FROM Companys WHERE Id='" & companyId & "'")
            If companyId = "3" Then companyName = "PT Bumi Indah Global"

            Dim mailData As DataRow = GetDataRow("SELECT * FROM Mailings WHERE CompanyId='" & companyId & "' AND Name='Reset Proforma Order' AND Active=1")
            If mailData Is Nothing Then Exit Sub

            Dim mailServer As String = mailData("Server").ToString()
            Dim mailHost As String = mailData("Host").ToString()
            Dim mailPort As Integer = mailData("Port")

            Dim mailAccount As String = mailData("Account").ToString()
            Dim mailPassword As String = mailData("Password").ToString()
            Dim mailAlias As String = mailData("Alias").ToString()
            Dim mailSubject As String = mailData("Subject").ToString()

            Dim mailTo As String = mailData("To").ToString()
            Dim mailCc As String = mailData("Cc").ToString()
            Dim mailBcc As String = mailData("Bcc").ToString()

            Dim mailNetworkCredentials As Boolean = mailData("NetworkCredentials")
            Dim mailDefaultCredentials As Boolean = mailData("DefaultCredentials")
            Dim mailEnableSSL As Boolean = mailData("EnableSSL")

            Dim customerMail As DataTable = GetDataTable("SELECT Email FROM CustomerContacts CROSS APPLY STRING_SPLIT(Tags, ',') AS thisArray WHERE CustomerId='" & customerId & "' AND thisArray.VALUE='Confirming'")
            If customerMail.Rows.Count = 0 Then Exit Sub

            Dim mailBody As String = String.Empty

            mailBody = "<span style='font-family: Cambria; font-size: 16px;'>"
            mailBody &= "<i>- THIS IS AN AUTOMATED EMAIL. KINDLY DO NOT REPLY WITHOUT COPYING OUR TEAM. -</i>"
            mailBody &= "<br /><br /><br />"
            mailBody &= "Dear Valued Customer,"
            mailBody &= "<br /><br />"
            mailBody &= "This order has been moved to <b>Pending Payment</b>.</u></b>."
            mailBody &= "<br /><br />"
            mailBody &= "Please note that if payment is received after a price change has taken effect, the order will be processed using our <b>latest pricing</b>."
            mailBody &= "<br /><br />"
            mailBody &= "If our pricing changes before your payment is received, the order status will be <b>automatically updated to Unsubmitted</b>."
            mailBody &= "<br />"
            mailBody &= "The order will be repriced according to our <b>current price list</b>, and you will be <b>required to review and resubmit the order</b> before processing can continue."
            mailBody &= "<br /><br />"
            mailBody &= "Thank you for your understanding."
            mailBody &= "</span>"

            mailBody &= "<br /><br /><br />"

            mailBody &= "<span style='font-family: Cambria; font-size:16px;'>Kind Regards,</span>"
            mailBody &= "<br /><br /><br />"
            mailBody &= "<span style='font-family: Cambria; font-size:16px; font-weight: bold;'>" & companyName.ToUpper() & "</span>"

            Dim myMail As New MailMessage()

            Dim subject As String = String.Format("{0} - {1} - {2} - Due Date Order # {3}", customerName, orderNumber, orderName, orderId)

            myMail.Subject = subject
            myMail.From = New MailAddress(mailServer, mailAlias)

            If customerMail.Rows.Count > 0 Then
                For i As Integer = 0 To customerMail.Rows.Count - 1
                    Dim thisEmail As String = customerMail.Rows(i)("Email").ToString()
                    myMail.To.Add(thisEmail)
                Next
            End If

            If Not String.IsNullOrEmpty(mailCc) Then
                For Each thisMail In mailCc.Split(";"c)
                    If Not String.IsNullOrEmpty(thisMail.Trim()) Then myMail.CC.Add(thisMail.Trim())
                Next
            End If

            If companyId = "2" Then
                Dim operatorEmail As String = GetItemData("SELECT ISNULL(STRING_AGG(Logins.Email, ';'), '') FROM Customers OUTER APPLY STRING_SPLIT(Customers.Operator, ',') operatorArray LEFT JOIN Logins ON Logins.Id = TRY_CAST(operatorArray.value AS INT) WHERE Customers.Id='" & customerId & "';")

                If Not String.IsNullOrEmpty(operatorEmail) Then
                    Dim emailList() As String = operatorEmail.Split(";"c)

                    For Each email As String In emailList
                        If Not String.IsNullOrWhiteSpace(email) Then
                            myMail.CC.Add(email.Trim())
                        End If
                    Next
                End If
            End If

            If Not String.IsNullOrEmpty(mailBcc) Then
                For Each thisMail In mailBcc.Split(";"c)
                    If Not String.IsNullOrEmpty(thisMail.Trim()) Then myMail.Bcc.Add(thisMail.Trim())
                Next
            End If

            myMail.IsBodyHtml = True
            myMail.Body = mailBody
            Dim smtpClient As New SmtpClient()
            smtpClient.Host = mailHost
            smtpClient.Port = mailPort
            smtpClient.EnableSsl = mailEnableSSL
            smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network
            smtpClient.Timeout = 120000

            If mailNetworkCredentials Then
                smtpClient.UseDefaultCredentials = False
                smtpClient.Credentials = New NetworkCredential(mailAccount, mailPassword)
            Else
                smtpClient.UseDefaultCredentials = mailDefaultCredentials
            End If

            smtpClient.Send(myMail)
        Catch ex As Exception
        End Try
    End Sub
End Class
