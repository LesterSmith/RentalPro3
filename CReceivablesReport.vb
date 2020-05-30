'****************************************
'* Purpose: Create Accounts Receivables Report
'*
'* Author:  Les Smith
'* Date Created: 06/16/2003 at 08:16:49
'* CopyRight:  HHI Software
'****************************************
'*
Imports System
Imports System.Text

Public Class CReceivablesReport
   Private oDA As New CDataAccess()

#Region " Public Methods "
   Public Sub PrintReceivables()
      ' 1) Get table of customers with outstanding invoices
      ' 2) Print the customer information like Name, Contact, Phone
      ' 3) print the open invoice headers
      Dim i As Integer
      Dim SQL As String = ""
      Dim dt As New DataTable()
      Dim pd As New StringBuilder()
      Dim iCustId As Integer
      Dim custtotal As Decimal
      Dim grandtotal As Decimal
      Dim s As String

      SQL = ""
      SQL &= "select c.CustomerId,c.CompanyName, c.PhoneNumber, i*  "
      SQL &= "from customers c,invoices i "
      SQL &= "where i.balancedue<>0 "
      SQL &= "and i.CustomerId=c.CustomerId "
      SQL &= "order by c.companyname "

      If oDA.SendQuery(SQL, dt, ConnectString) = 0 Then
         MsgBox("There are no customers with open balances.", MsgBoxStyle.Information)
         Exit Sub
      End If

      iCustId = dt.Rows(0).Item("customerid")

      For i = 0 To dt.Rows.Count - 1
         With dt.Rows(i)
            If .Item("customerid") <> iCustId Then

               pd.Append("Customer Total: " & FormatCurrency(custtotal) & vbCrLf & vbCrLf)
               grandtotal += custtotal
               custtotal = 0
            End If
            If custtotal = 0 Then
               pd.Append("Customer ID: " & .Item("c.customerid").ToString.PadRight(6) & _
                         "Name: " & .Item("CompanyName").ToString.PadRight(40) & _
                         "Phone: " & .Item("phonenumber") & vbCrLf)
               custtotal += .Item("i.balancedue")
               pd.Append(Space(5) & _
                        "Invoice ID".PadRight(12) & _
                        "Date".PadRight(10) & _
                        "Contact Name".PadRight(21) & _
                        "P O Number".PadRight(12) & _
                        "Balance".PadRight(12) & _
                        "Age/Days" & vbCrLf)
            End If

            pd.Append(Space(5) & _
                      .Item("invoiceid").ToString.PadLeft(10) & "  " & _
                      Format(.Item("invoicedate"), "MM/dd/yyyy").PadRight(10))
            s = .Item("contactname").ToString
            If s.Length > 20 Then
               s = s.Substring(0, 20)
            End If
            pd.Append(s.PadRight(21))
            s = .Item("ponumber")
            If s.Length > 10 Then
               s = s.Substring(0, 10)
            End If
            pd.Append(s.PadRight(12))
            pd.Append(FormatCurrency(.Item("balancedue")).PadLeft(10) & "  ")
            Dim iDays As Integer = DateDiff(DateInterval.Day, .Item("invoicedate"), Today)
            pd.Append(iDays.ToString.PadLeft(5) & vbCrLf)

         End With
      Next
   End Sub

#End Region

End Class
