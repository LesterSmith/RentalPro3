''' Create the ReRent Report String and call CPrintString.
''' Imports System.Text
Imports System.Text

Public Class CPrintReRentReport
#Region " Class Level Variables "
   Private _Preview As Boolean
   Private sDate As Date
   Private eDate As Date


#End Region



#Region " Constructor "
   ''' <summary>
   ''' Class Consturctor.
   ''' </summary>
   ''' <param name = "Preview"></param>
   ''' <param name = "StartDate"></param>
   ''' <param name = "EndDate"></param>
   Public Sub New(ByVal Preview As Boolean, ByVal StartDate As Date, ByVal EndDate As Date)
      _Preview = Preview
      sDate = DateValue(StartDate)
      eDate = DateValue(EndDate)
   End Sub

#End Region


#Region " Public Methods "
   ''' <summary>
   ''' Create the ReRent Report String and call 
   ''' the CPrintString class to print it.
   ''' </summary>
   Public Sub ReRentPrint()
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
         SQL = "select d.invoiceid,d.quantity,d.priceperunit,d.equip_name, "
         SQL &= "d.rental_period,d.record_description,d.rented_date, "
         SQL &= "d.customer_id,c.companyname "
         SQL &= "from invoice_details d,customers c "
         SQL &= "where d.equip_id='" & RERENT & "' "
         SQL &= "and d.rented_date >= #" & sDate & "# "
         SQL &= "and d.rented_date < #" & DateAdd(DateInterval.Day, 1, eDate) & "# "
         SQL &= "and d.customer_id = c.customerid "
         SQL &= "order by rented_date"
         If oDA.SendQuery(SQL, dt, ConnectString) = 0 Then
            MsgBox("No data for selected period.", MsgBoxStyle.Information)
            Exit Sub
         End If

         Dim colHdr As String = _
            "Invoice".PadRight(8) & _
            "Equip Desc".PadRight(28) & _
            "Customer".PadRight(28) & _
            "Nbr".PadRight(4) & _
            "Period".PadRight(8) & _
            "P.O. Nbr".PadRight(11) & _
            "Rent Date".PadRight(11) & _
            "Revenue".PadRight(10)

         For i = 0 To dt.Rows.Count - 1
            Dim dr As DataRow = dt.Rows(i)
            With dr
               sb.Append(CType(dr("invoiceid"), String).PadRight(8))
               sb.Append(LS(CType(dr("equip_name"), String), 27).PadRight(28))
               sb.Append(LS(CType(dr("companyname"), String), 27).PadRight(28))
               sb.Append(Format(dr("quantity"), "##0").PadLeft(3) & " ")
               sb.Append(CType(dr("rental_period"), String).PadRight(8))
               sb.Append(LS(MNS(dr("record_description")), 10).PadRight(11))
               sb.Append(Format(dr("rented_date"), "MM/dd/yyyy").PadRight(11))
               sb.Append(FormatCurrency(MNSng(dr("quantity")) * MND(dr("priceperunit"))).PadLeft(10))
               sb.Append(vbCrLf)
               gTotal += MNSng(dr("quantity")) * MND(dr("priceperunit"))
            End With
         Next
         sb.Append(vbCrLf & Space(88) & "Total".PadRight(10) & FormatCurrency(gTotal).PadLeft(10))
         oPS = New CPrintStringNew()
         oPS.TitleFontStyle = "BI"
         oPS.TitleFontSize = REPORT_TITLE_FONT_SIZE
         If _Preview Then
            oPS.PrintPreview(120, sb.ToString, _
            ReportName, _
            "ReRent Tracking Report", _
            colHdr1:=colHdr)
         Else
            oPS.StartPrint(120, sb.ToString, _
               ReportName, _
               "ReRent Tracking Report", _
               colHdr1:=colHdr)
         End If
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

#End Region


End Class
