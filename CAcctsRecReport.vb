'****************************************
'* Purpose: Print Accounts Receivables Report.
'*
'* Author:  Unregistered User
'* Date Created: 07/06/2003 at 09:25:05
'* CopyRight:  Unregistered Company
'****************************************
'*
Imports System.Text

Public Class CAcctsRecReport
   Private m_Preview As Boolean

#Region " Public Methods "
   Public Property Preview() As Boolean
      Get
         Return m_Preview
      End Get
      Set(ByVal Value As Boolean)
         m_Preview = Value
      End Set
   End Property

   ''' <summary>
   ''' Print the aged ar report with the following line.
   ''' Customer 30-60 60-90 90-120 120-180 >180
   ''' </summary>
   Public Sub PrintAgedARReport()
      Dim iLastCustomer As Integer = 0
      Dim dt As New DataTable()
      Dim i As Integer
      Dim sb As New StringBuilder()
      Dim SQL As String
      Dim oDA As New CDataAccess()
      Dim decInvoiceTotal As Decimal = 0
      Dim s As String
      Dim bNewCustomer As Boolean = True
      Dim oPS As CPrintStringNew
      Const SL = 5
      Dim invDate As Date
      Dim daysPastDue As Integer = 0
      Dim invTotal As Decimal = 0
      Dim gTotal As Decimal = 0
      Dim daysCurrent As Decimal = 0
      Dim days30to60 As Decimal = 0
      Dim days60to90 As Decimal = 0
      Dim days90to120 As Decimal = 0
      Dim days120to150 As Decimal = 0
      Dim days150to180 As Decimal = 0
      Dim daysOver150 As Decimal = 0
      Dim gtDaysCurrent As Decimal = 0
      Dim gtDays30to60 As Decimal = 0
      Dim gtDays60to90 As Decimal = 0
      Dim gtDays90to120 As Decimal = 0
      Dim gtDays120to150 As Decimal = 0
      Dim gtDaysOver150 As Decimal = 0
      Dim balDue As Decimal = 0
      Dim lastCustomerName As String = ""

      Try

         sb = New StringBuilder()
         SQL = "select I.balancedue,i.invoicedate,C.Companyname,c.customerid "
         SQL &= "from invoices i, customers c "
         SQL &= "where i.customerid = c.customerid "
         SQL &= "and i.balancedue > 0 "
         SQL &= "and i.status <> 'CheckedOut' "
         SQL &= "order by c.companyname,i.invoicedate"
         Dim iRows As Integer = oDA.SendQuery(SQL, dt, ConnectString)
         If iRows < 1 Then
            MsgBox("There are no receivables to print.", MsgBoxStyle.Information)
            Exit Sub
         End If
         ' ______________________________________________________________________
         ' 0000000001111111111222222222233333333334444444444555555555566666666667777777778888888888
         ' 1234567890123456789012345678901234567890123456789012345678901234567890123456890123456789
         ' Customer Name                   30-60    60-90    90-120  120-150  150-180  > 180
         '                                123,456  123,456  123,456  123,456  123,456 123,456
         '-----------------------------------------------------------------
         ' header1
         Dim columnHdr1 As String = _
            "Customer Name".PadRight(30) & _
            "      1-30  " & _
            "     30-60  " & _
            "    60-90  " & _
            "   90-120  " & _
            "  120-150  " & _
            "    > 150  " & vbCrLf


         With dt.Rows(0)
            iLastCustomer = dt.Rows(0).Item("customerid")
            lastCustomerName = .Item("companyname")
         End With

         For i = 0 To dt.Rows.Count - 1
            With dt.Rows(i)
               If iLastCustomer <> .Item("customerid") Then
                  sb.Append(lastCustomerName.PadRight(30))
                  sb.Append(Format(daysCurrent, "#,##0.00").PadLeft(11))
                  sb.Append(Format(days30to60, "#,##0.00").PadLeft(11))
                  sb.Append(Format(days60to90, "#,##0.00").PadLeft(11))
                  sb.Append(Format(days90to120, "#,##0.00").PadLeft(11))
                  sb.Append(Format(days120to150, "#,##0.00").PadLeft(11))
                  sb.Append(Format(daysOver150, "#,##0.00").PadLeft(11))
                  sb.Append(vbCrLf)
                  gtDaysCurrent += daysCurrent
                  daysCurrent = 0
                  gtDays30to60 += days30to60
                  days30to60 = 0
                  gtDays60to90 += days60to90
                  days60to90 = 0
                  gtDays90to120 += days90to120
                  days90to120 = 0
                  gtDays120to150 += days120to150
                  days120to150 = 0
                  gtDaysOver150 += daysOver150
                  daysOver150 = 0
                  iLastCustomer = .Item("customerid")
                  lastCustomerName = .Item("companyname")
               End If
               invDate = .Item("invoicedate")
               balDue = .Item("balancedue")
               If DateDiff(DateInterval.Day, invDate, Today) > 150 Then
                  daysOver150 += balDue
               ElseIf DateDiff(DateInterval.Day, invDate, Today) > 120 Then
                  days120to150 += balDue
               ElseIf DateDiff(DateInterval.Day, invDate, Today) > 90 Then
                  days90to120 += balDue
               ElseIf DateDiff(DateInterval.Day, invDate, Today) > 60 Then
                  days60to90 += balDue
               ElseIf DateDiff(DateInterval.Day, invDate, Today) > 30 Then
                  days30to60 += balDue
               Else
                  daysCurrent += balDue
               End If

            End With
         Next i
         If days30to60 > 0 Or _
            days60to90 > 0 Or _
            days90to120 > 0 Or _
            days120to150 > 0 Or _
            daysCurrent > 0 Or _
            daysOver150 > 0 Then
            sb.Append(lastCustomerName.PadRight(30))
            sb.Append(Format(daysCurrent, "#,##0.00").PadLeft(11))
            sb.Append(Format(days30to60, "#,##0.00").PadLeft(11))
            sb.Append(Format(days60to90, "#,##0.00").PadLeft(11))
            sb.Append(Format(days90to120, "#,##0.00").PadLeft(11))
            sb.Append(Format(days120to150, "#,##0.00").PadLeft(11))
            sb.Append(Format(daysOver150, "#,##0.00").PadLeft(11))
            sb.Append(vbCrLf)
         End If
         sb.Append(vbCrLf & "Total Receivables".PadRight(30))
         sb.Append(Format(gtDaysCurrent, "#,##0.00").PadLeft(11))
         sb.Append(Format(gtDays30to60, "#,##0.00").PadLeft(11))
         sb.Append(Format(gtDays60to90, "#,##0.00").PadLeft(11))
         sb.Append(Format(gtDays90to120, "#,##0.00").PadLeft(11))
         sb.Append(Format(gtDays120to150, "#,##0.00").PadLeft(11))
         sb.Append(Format(gtDaysOver150, "#,##0.00").PadLeft(11))
         sb.Append(vbCrLf)


         oPS = New CPrintStringNew()
         oPS.TitleFontStyle = "BI"
         oPS.TitleFontSize = REPORT_TITLE_FONT_SIZE
         If Preview Then
            oPS.PrintPreview(96, sb.ToString, _
            ReportName, _
            "Aged Accounts Receivables Report", _
            colHdr1:=columnHdr1)
         Else
            oPS.StartPrint(96, sb.ToString, _
               ReportName, _
                "Aged Accounts Receivables Report", _
               colHdr1:=columnHdr1)
         End If
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

   ''' <summary>
   ''' this method will get a recordset of unpaid invoices,
   ''' joined with the customer billing address
   ''' it will print each customers statment on a different page
   ''' </summary>
   Public Sub PrintARReport()
      Dim iLastCustomer As Integer = 0
      Dim dt As New DataTable()
      Dim i As Integer
      Dim sb As New StringBuilder()
      Dim SQL As String
      Dim oDA As New CDataAccess()
      Dim decInvoiceTotal As Decimal = 0
      Dim s As String
      Dim bNewCustomer As Boolean = True
      Dim oPS As CPrintStringNew
      Const SL = 5
      Dim invDate As Date
      Dim daysPastDue As Integer
      Dim invTotal As Decimal = 0
      Dim gTotal As Decimal = 0

      Try

         sb = New StringBuilder()
         SQL = "select I.*,C.Companyname, c.phonenumber "
         SQL &= "from invoices i, customers c "
         SQL &= "where i.customerid = c.customerid "
         SQL &= "and i.balancedue > 0 "
         SQL &= "and i.status <> 'CheckedOut' "
         SQL &= "order by c.customerid,i.invoiceid"
         Dim iRows As Integer = oDA.SendQuery(SQL, dt, ConnectString)
         If iRows < 1 Then
            MsgBox("There are no receivables to print.", MsgBoxStyle.Information)
            Exit Sub
         End If
         ' ______________________________________________________________________
         ' 0000000001111111111222222222233333333334444444444555555555566666666667
         ' 1234567890123456789012345678901234567890123456789012345678901234567890
         ' Invoice #     Invoice Date   P.O. Number  Picked Up By: Invoice Amount
         '-----------------------------------------------------------------
         ' header1
         Dim columnHdr1 As String = _
            "Customer Name".PadRight(28) & _
            "Customer ID".PadRight(14) & _
            "Phone Number" & vbCrLf
         Dim columnHdr2 As String = _
            "   Invoice #".PadRight(18) & _
            "Invoice Date".PadRight(13) & _
            "P.O. Number".PadRight(12) & _
            "Days Past Due".PadRight(17) & _
            "Invoice Amt"

         iLastCustomer = dt.Rows(0).Item("customerid")
         With dt.Rows(0)
            sb.Append(CType(MNS(.Item("CompanyName")), String).PadRight(28) & _
                      CType(.Item("CustomerID"), String).PadRight(14) & _
                      CType(MNS(.Item("Phonenumber")), String) & vbCrLf)
         End With

         For i = 0 To dt.Rows.Count - 1
            With dt.Rows(i)
               ' print invoice detail data
               If .Item("customerid") = iLastCustomer Then
                  sb.Append(Space(3) & _
                     CType(.Item("invoiceid"), String).PadRight(15) & _
                     Format(.Item("invoicedate"), "MM/dd/yyyy").PadRight(13) & _
                     CType(.Item("ponumber"), String).PadRight(12))
                  invDate = .Item("invoicedate")
                  daysPastDue = DateDiff(DateInterval.Day, invDate, Today)
                  If daysPastDue > 30 Then
                     daysPastDue -= 30
                  Else
                     daysPastDue = 0
                  End If

                  sb.Append(Space(5) & daysPastDue.ToString.PadRight(14 - 5))
                  sb.Append(FormatCurrency(.Item("balancedue")).PadLeft(13) & vbCrLf)
                  invTotal += .Item("balancedue")
                  gTotal += .Item("balancedue")
               Else
                  sb.Append(Space(49) & "Total:  " & _
                     FormatCurrency(invTotal).PadLeft(13) & vbCrLf & vbCrLf)
                  invTotal = 0

                  sb.Append(MNS(.Item("CompanyName")).PadRight(28) & _
                            CType(.Item("CustomerID"), String).PadRight(11) & _
                            CType(MNS(.Item("Phonenumber")), String) & vbCrLf)
                  iLastCustomer = dt.Rows(i).Item("customerid")
                  sb.Append(Space(3) & _
                     CType(.Item("invoiceid"), String).PadRight(15) & _
                     Format(.Item("invoicedate"), "MM/dd/yyyy").PadRight(13) & _
                     CType(.Item("ponumber"), String).PadRight(12))

                  invDate = .Item("invoicedate")

                  daysPastDue = DateDiff(DateInterval.Day, invDate, Today)

                  If daysPastDue > 30 Then
                     daysPastDue -= 30
                  Else
                     daysPastDue = 0
                  End If

                  sb.Append(Space(5) & daysPastDue.ToString.PadRight(14 - 5))
                  sb.Append(FormatCurrency(.Item("balancedue")).PadLeft(13) & vbCrLf)
                  invTotal += .Item("balancedue")
                  gTotal += .Item("balancedue")
               End If
            End With
         Next i
         If invTotal > 0 Then
            sb.Append(Space(49) & "Total:  " & _
               FormatCurrency(invTotal).PadLeft(13) & vbCrLf & vbCrLf)
         End If

         sb.Append(Space(43) & "Grand Total:  " & _
            FormatCurrency(gTotal).PadLeft(13) & vbCrLf & vbCrLf)

         oPS = New CPrintStringNew()
         oPS.TitleFontStyle = "BI"
         oPS.TitleFontSize = REPORT_TITLE_FONT_SIZE
         If Preview Then
            oPS.PrintPreview(80, sb.ToString, _
            ReportName, _
            "Accounts Receivables Report", _
            colHdr1:=columnHdr1, _
            colHdr2:=columnHdr2)
         Else
            oPS.StartPrint(80, sb.ToString, _
               ReportName, _
               "Accounts Receivables Report", _
               colHdr1:=columnHdr1, _
               colHdr2:=columnHdr2)
         End If
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

#End Region

End Class
