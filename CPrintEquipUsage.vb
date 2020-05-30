Imports System.Windows.Forms.Application
Public Class CPrintEquipUsage
#Region " Class Level Variables "
   Private oDA As New CDataAccess()
   Private SQL As String
   Private m_Preview As Boolean
   Private m_StDate As Date
   Private m_EndDate As Date


#End Region


#Region " Public Properties "
   Public Property Preview() As Boolean
      Get
         Return m_Preview
      End Get
      Set(ByVal Value As Boolean)
         m_Preview = Value
      End Set
   End Property

   Public Property StDate() As Date
      Get
         Return m_StDate
      End Get
      Set(ByVal Value As Date)
         m_StDate = Value
      End Set
   End Property

   Public Property EndDate() As Date
      Get
         Return m_EndDate
      End Get
      Set(ByVal Value As Date)
         m_EndDate = Value
      End Set
   End Property



#End Region



#Region " Public Methods "
   ' Print the equipment usage report.
   Public Sub PrintUsageReport()
      Dim ps As New System.Text.StringBuilder()
      Dim i As Integer
      Dim dt As New DataTable()
      Dim oCP As New CPrintStringNew()
      Dim decTotal As Decimal
      Dim grandTotal As Decimal
      Dim totalDays As Single
      Dim lastEquip As String

      ' d.quantity,d.priceperunit,d.equip_id,d.equip_name,d.rental_period
      ' d.invoiceid,d.customer_id
      ' i.status,i.invoiceid


      Try
         Dim colHdr As String = ""
         colHdr &= "Equip ID".PadRight(10)
         colHdr &= "Equip Name".PadRight(27)
         colHdr &= "Rent Date".PadRight(11)
         colHdr &= "Returned".PadRight(11)
         colHdr &= "Per".PadRight(4)
         colHdr &= "No.".PadRight(4)
         colHdr &= "Days".PadRight(4)
         colHdr &= "Revenue".PadLeft(12)

         Dim title As String = ReportName
         Dim subTitle As String = _
            "Equipment Usage Report: " & Format(Me.StDate, "M/d/yyyy") & _
            " To " & Format(Me.EndDate, "M/d/yyyy")
         SQL = "select d.quantity,d.priceperunit,d.equip_id, "
         SQL &= "d.equip_name,d.rental_period,d.rented_date,d.returned_date "
         SQL &= "from invoice_details d, invoices i "
         SQL &= "where i.status <> 'CheckedOut' "
         SQL &= "and i.invoiceid = d.invoiceid "
         SQL &= "and d.record_type = 15 "
         SQL &= "and (d.rental_period = '" & DAILY & "' "
         SQL &= "or d.rental_period = '" & HALF_DAY & "' "
         SQL &= "or d.rental_period = '" & WEEKLY & "' "
         SQL &= "or d.rental_period = '" & MONTHLY & "' "
         SQL &= "or d.rental_period = '" & WEEK_END & "') "
         SQL &= "and d.rented_date >= #" & Me.m_StDate & "# "
         SQL &= "and d.returned_date <= #" & Me.m_EndDate & "# "
         SQL &= "order by d.equip_id,d.rented_date "

         If oDA.SendQuery(SQL, dt, ConnectString) = 0 Then
            MsgBox("No equipment rented and returned during selected date range.", MsgBoxStyle.Exclamation)
            Exit Sub
         End If

         Dim eqName As String
         Dim per As String
         Dim qty As Integer
         Dim ppUnit As Decimal
         Dim days As Single
         Dim rentedDate As String
         Dim returnedDate As String

         lastEquip = MNS(dt.Rows(0).Item("equip_id"))

         For i = 0 To dt.Rows.Count - 1
            With dt.Rows(i)
               If lastEquip <> MNS(.Item("equip_id")) Then
                  ps.Append(vbCrLf & Space(59) & "Equip".PadRight(8) & Format(totalDays, "##0.0").PadLeft(5) & FormatCurrency(decTotal).PadLeft(12) & vbCrLf & vbCrLf)
                  lastEquip = MNS(.Item("equip_id"))
                  decTotal = 0
                  totalDays = 0
               End If

               ps.Append(MNS(.Item("equip_id")).PadRight(10))
               eqName = MNS(.Item("equip_name"))
               If eqName.Trim.Length > 25 Then
                  eqName = eqName.Substring(0, 25)
               End If
               ps.Append(eqName.PadRight(27))
               If IsDBNull(.Item("rented_date")) Then
                  rentedDate = ""
               Else
                  rentedDate = Format(.Item("rented_date"), "MM/dd/yyyy")
               End If
               If IsDBNull(.Item("returned_date")) Then
                  returnedDate = ""
               Else
                  returnedDate = Format(.Item("returned_date"), "MM/dd/yyyy")
               End If

               ps.Append(rentedDate.PadRight(11))
               ps.Append(returnedDate.PadRight(11))
               per = .Item("rental_period")
               qty = MNI(.Item("quantity"))
               ppUnit = MND(.Item("priceperunit"))
               Select Case per
                  Case DAILY : per = "DY" : days = qty
                  Case WEEKLY : per = "WK" : days = qty * 5
                  Case MONTHLY : per = "MO" : days = qty * 21.6
                  Case HALF_DAY : per = "HD" : days = qty * 0.5
                  Case WEEK_END : per = "WE" : days = qty * 1
               End Select
               ppUnit = qty * ppUnit
               ps.Append(per.PadRight(4))
               ps.Append(Format(qty, "0").PadRight(4))
               ps.Append(Format(days, "##0.0").PadLeft(5))
               ps.Append(FormatCurrency(ppUnit).PadLeft(12))
               ps.Append(vbCrLf)
               decTotal += ppUnit
               grandTotal += ppUnit
               totalDays += days
            End With
         Next

         If decTotal > 0 Then
            ps.Append(vbCrLf & Space(59) & "Equip".PadRight(8) & Format(totalDays, "##0.0").PadLeft(5) & FormatCurrency(decTotal).PadLeft(12) & vbCrLf)
         End If
         ps.Append(vbCrLf & Space(59) & "Total".PadRight(13) & FormatCurrency(grandTotal).PadLeft(12) & vbCrLf)

         oCP.TitleFontSize = 14
         oCP.TitleFontStyle = "BI"
         oCP.TitleFontSize = REPORT_TITLE_FONT_SIZE

         If Me.Preview Then
            oCP.PrintPreview(96, _
               ps.ToString, _
               title, _
               subTitle, _
               ColHdr1:=colHdr)
         Else
            oCP.StartPrint(96, _
               ps.ToString, _
               title, _
               subTitle, _
               ColHdr1:=colHdr)
         End If


      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

#End Region

End Class
