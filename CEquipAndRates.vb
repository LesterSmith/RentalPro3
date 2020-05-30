''' Print List of all equipment.
Imports System.Text
Public Class CEquipAndRates
#Region " Module Variables "
   Dim oPS As New CPrintStringNew()
   Private oDA As CDataAccess
   Private SQL As String
   Private sb As StringBuilder
   Private sb2 As StringBuilder
   Private oUtil As CUtilities
   Private m_Preview As Boolean
   Private _SortOrder As String = "NAME"

#End Region

#Region "Processing Members"
   ''' <summary>
   ''' Print the equipment list
   ''' </summary>
   Public Sub PrintEquipmentList()
      Dim dt As New DataTable()
      Dim i As Integer
      Dim dr As DataRow
      Dim s As String
      Dim tot As Decimal


      Try
         oDA = New CDataAccess()
         sb = New StringBuilder()
         SQL = "select a.Equip_id,a.equip_name, b.price_id,b.hourrate, "
         SQL &= "b.halfday,b.daily,b.weekly,b.weekend,b.monthly "
         SQL &= "from equipment a,rental_rates b "
         SQL &= "where a.price_id = b.price_id "
         SQL &= "order by "
         If _SortOrder = "ID" Then
            SQL &= "a.equip_id"
         ElseIf _SortOrder = "NAME" Then
            SQL &= "a.equip_name"
         Else
            SQL &= "a.equip_type_id,a.equip_name"
         End If

         If oDA.SendQuery(SQL, dt, ConnectString) = 0 Then
            MsgBox("No equipment found for selected period.", MsgBoxStyle.Critical)
            Exit Sub
         End If
         Dim colHdr As String = _
            "Equipment Name".PadRight(25) & _
            "Equip ID".PadRight(10) & _
            "Rate ID".PadRight(9) & _
            "Hour".PadRight(5) & _
            "Half Day".PadRight(9) & _
            "Daily".PadRight(7) & _
            "WeekEnd".PadRight(8) & _
            "Weekly".PadRight(8) & _
            "Monthly".PadRight(10)

         For i = 0 To dt.Rows.Count - 1
            dr = dt.Rows(i)
            s = LS(MNS(dr("equip_name")), 24)
            sb.Append(s.PadRight(25))
            sb.Append(MNS(dr("equip_id")).PadRight(10))
            sb.Append(MNI(dr("price_id")).ToString.PadRight(8))
            sb.Append(FormatNDS(MND(dr("hourrate"))).PadLeft(5))
            sb.Append(FormatNDS(MND(dr("halfday"))).PadLeft(9))
            sb.Append(FormatNDS(MND(dr("daily"))).PadLeft(7))
            sb.Append(FormatNDS(MND(dr("weekend"))).PadLeft(8))
            sb.Append(FormatNDS(MND(dr("weekly"))).PadLeft(8))
            sb.Append(FormatNDS(MND(dr("monthly"))).PadLeft(10))
            sb.Append(vbCrLf)
         Next
         oPS.TitleFontSize = 14
         oPS.TitleFontStyle = "BI"
         If Me.m_Preview Then
            oPS.PrintPreview(96, sb.ToString, ReportName, "Rental Equipment List & Rates", ColHdr1:=colHdr)
         Else
            oPS.StartPrint(96, sb.ToString, ReportName, "Rental Equipment List & Rates", ColHdr1:=colHdr)
         End If

      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

#End Region

#Region "Public Properties"

   Public Property Preview() As Boolean
      Get
         Return m_Preview
      End Get
      Set(ByVal Value As Boolean)
         m_Preview = Value
      End Set
   End Property

   Public Property SortOrder() As String
      Get
         Return _SortOrder
      End Get
      Set(ByVal Value As String)
         _SortOrder = Value
      End Set
   End Property
#End Region
End Class
