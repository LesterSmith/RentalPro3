''' Print List of all equipment.
Imports System.Text
Public Class CEquipmentList
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
         SQL = "select Equip_id,equip_name,serial_number,model_number,purchase_price,equip_type_id "
         SQL &= "from equipment "
         SQL &= "order by "
         If _SortOrder = "ID" Then
            SQL &= "equip_id"
         ElseIf _SortOrder = "NAME" Then
            SQL &= "equip_name"
         Else
            SQL &= "equip_type_id,equip_name"
         End If

         If oDA.SendQuery(SQL, dt, ConnectString) = 0 Then
            MsgBox("No equipment found for selected period.", MsgBoxStyle.Critical)
            Exit Sub
         End If
         Dim colHdr As String = _
            "Equipment Name".PadRight(25) & _
            "Equip ID".PadRight(10) & _
            "Serial Nbr".PadRight(20) & _
            "Model Nbr".PadRight(15) & _
            "Purchase".PadRight(10) & _
            "Cat".PadRight(4)

         For i = 0 To dt.Rows.Count - 1
            dr = dt.Rows(i)
            s = MNS(dr("equip_name"))
            If s.Length > 24 Then
               s = s.Substring(0, 24)
            End If
            sb.Append(s.PadRight(25))
            sb.Append(MNS(dr("equip_id")).PadRight(10))
            sb.Append(MNS(dr("serial_number")).PadRight(20))
            sb.Append(MNS(dr("model_number")).PadRight(15))
            sb.Append(FormatCurrency(MND(dr("purchase_price"))).PadLeft(10))
            sb.Append(MNI(dr("equip_type_id")).ToString.PadLeft(5))
            tot += MND(dr("purchase_price"))
            sb.Append(vbCrLf)
         Next
         sb.Append(vbCrLf & "     Total Value".PadRight(69) & FormatCurrency(tot).PadLeft(10))
         oPS.TitleFontSize = 14
         oPS.TitleFontStyle = "BI"
         If Me.m_Preview Then
            oPS.PrintPreview(80, sb.ToString, ReportName, "Rental Equipment List", ColHdr1:=colHdr)
         Else
            oPS.StartPrint(80, sb.ToString, ReportName, "Rental Equipment List", ColHdr1:=colHdr)
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
