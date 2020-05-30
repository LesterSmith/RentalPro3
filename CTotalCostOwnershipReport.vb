''' Print Total Cost of OwnerShip Report.
Imports System.Text
Imports System.Windows.Forms.Application

Public Class CTotalCostOwnershipReport
#Region " Class Level Variables "
   Private _Preview As Boolean
   Private _Sort As String
   Dim oDA As New CDataAccess()


#End Region

#Region " Constructors "
   ''' <summary>
   ''' Constructor
   ''' </summary>
   ''' <param name = "Preview"></param>
   ''' <param name = "sort"></param>
   Public Sub New(ByVal Preview As Boolean, ByVal sort As String)
      _Preview = Preview
      _Sort = sort
   End Sub

#End Region

#Region " Public Methods "
   ''' <summary>
   ''' Return rental income for requested equipment.
   ''' </summary>
   ''' <param name = "equipID"></param>
   ''' <returns>Decimal</returns>
   Private Function GetIncome(ByVal equipID As String) As Decimal
      ' get rental income
      Dim dt As New DataTable()
      Dim rentalIncome As Decimal
      Dim sql As String
      Dim i As Integer

      dt.Reset()
      sql = "select quantity,priceperunit "
      sql &= "from invoice_details "
      sql &= "where equip_id = '" & equipID & "' "
      If oDA.SendQuery(sql, dt, ConnectString) > 0 Then
         For i = 0 To dt.Rows.Count - 1
            With dt.Rows(0)
               rentalIncome += MND(.Item("quantity")) * _
                  MND(.Item("priceperunit"))
            End With
         Next
      End If
      Return rentalIncome
   End Function

   ''' <summary>
   ''' Create report print string and call CPrintStringNew
   ''' </summary>
   Public Sub PrintTCOReport()
      Dim lastID As String = String.Empty
      Dim lastName As String
      Dim dt As New DataTable()
      Dim i As Integer
      Dim sb As New StringBuilder()
      Dim SQL As String
      Dim oPS As CPrintStringNew
      Dim gTotal As Decimal = 0
      Dim purchase As Decimal
      Dim payments As Decimal
      Dim labor As Decimal
      Dim supplies As Decimal
      Dim damage As Decimal
      Dim income As Decimal
      Dim cost As Decimal
      Dim net As Decimal
      Dim purchaseGT As Decimal
      Dim paymentsGT As Decimal
      Dim laborGT As Decimal
      Dim suppliesGT As Decimal
      Dim damageGT As Decimal
      Dim incomeGT As Decimal
      Dim costGT As Decimal
      Dim netGT As Decimal

      Try

         sb = New StringBuilder()
         SQL = ""
         SQL &= "select c.equip_id,c.cost_type,c.cost_price,  "
         SQL &= "e.equip_name "
         SQL &= "from equip_cost c, equipment e "
         SQL &= "where e.equip_id=c.equip_id "
         SQL &= "order by "
         If _Sort = "ID" Then
            SQL &= "c.equip_id "
         Else
            SQL &= "e.equip_name,c.equip_id"
         End If

         If oDA.SendQuery(SQL, dt, ConnectString) = 0 Then
            MsgBox("No data for selected period.", MsgBoxStyle.Information)
            Exit Sub
         End If

         Dim colHdr As String = _
            "Equipment".PadRight(11) & _
            "Equip Name".PadRight(21) & _
            "Purchase".PadLeft(13) & _
            "Payments".PadLeft(13) & _
            "Bal Due".PadLeft(13) & _
            "Labor".PadLeft(13) & _
            "Supplies".PadLeft(13) & _
            "Damage".PadLeft(13) & _
            "Income".PadLeft(13) & _
            "Cost".PadLeft(13) & _
            "Net Gain".PadLeft(13)

         lastID = dt.Rows(0).Item("equip_id")
         lastName = dt.Rows(0).Item("equip_name")
         Dim haveData As Boolean
         Dim dr As DataRow = dt.Rows(i)

         For i = 0 To dt.Rows.Count - 1
            dr = dt.Rows(i)
            With dr
               If lastID <> MNS(dr("equip_id")) Then
                  ' print detail
                  income = Me.GetIncome(lastID)
                  incomeGT += income
                  cost = purchase + payments + labor + supplies + damage
                  costGT += cost
                  net = income - cost
                  netGT += net
                  sb.Append(lastID.PadRight(11))
                  sb.Append(LS(lastName, 20).PadRight(21))
                  sb.Append(FormatNDS(purchase).PadLeft(13))
                  sb.Append(FormatNDS(payments).PadLeft(13))
                  sb.Append(FormatNDS(purchase - payments).PadLeft(13))
                  sb.Append(FormatNDS(labor).PadLeft(13))
                  sb.Append(FormatNDS(supplies).PadLeft(13))
                  sb.Append(FormatNDS(damage).PadLeft(13))
                  sb.Append(FormatNDS(income).PadLeft(13))
                  sb.Append(FormatNDS(cost).PadLeft(13))
                  sb.Append(FormatNDS(income - cost).PadLeft(13) & vbCrLf)
                  lastID = MNS(dr("equip_id"))
                  lastName = MNS(dr("equip_name"))
                  purchaseGT += purchase
                  paymentsGT += payments
                  laborGT += labor
                  suppliesGT += supplies
                  damageGT += damage
                  purchase = 0
                  payments = 0
                  labor = 0
                  supplies = 0
                  damage = 0
                  income = 0
                  cost = 0
                  net = 0
               End If
               Select Case dr("cost_type")
                  Case 9 : purchase += dr("cost_Price")
                  Case 8 : payments += dr("cost_Price")
                  Case 1 : labor += dr("cost_Price")
                  Case 2 : supplies += dr("cost_Price")
                  Case 4 : damage += dr("cost_Price")
               End Select
            End With
         Next
         cost = purchase + payments + labor + supplies + damage
         If cost <> 0 Then
            income = Me.GetIncome(lastID)
            incomeGT += income
            costGT += cost
            net = income - cost
            netGT += net
            purchaseGT += purchase
            paymentsGT += payments
            laborGT += labor
            suppliesGT += supplies
            damageGT += damage
            sb.Append(lastID.PadRight(11))
            sb.Append(LS(lastName, 20).PadRight(21))
            sb.Append(FormatNDS(purchase).PadLeft(13))
            sb.Append(FormatNDS(payments).PadLeft(13))
            sb.Append(FormatNDS(purchase - payments).PadLeft(13))
            sb.Append(FormatNDS(labor).PadLeft(13))
            sb.Append(FormatNDS(supplies).PadLeft(13))
            sb.Append(FormatNDS(damage).PadLeft(13))
            sb.Append(FormatNDS(income).PadLeft(13))
            sb.Append(FormatNDS(cost).PadLeft(13))
            sb.Append(FormatNDS(income - cost).PadLeft(13) & vbCrLf)
         End If
         sb.Append(vbCrLf & "Total".PadRight(32) & _
             FormatCurrency(purchaseGT).PadLeft(13) & _
             FormatCurrency(paymentsGT).PadLeft(13) & _
             FormatCurrency(purchaseGT - paymentsGT).PadLeft(13) & _
             FormatCurrency(laborGT).PadLeft(13) & _
             FormatCurrency(suppliesGT).PadLeft(13) & _
             FormatCurrency(damageGT).PadLeft(13) & _
             FormatCurrency(incomeGT).PadLeft(13) & _
             FormatCurrency(costGT).PadLeft(13) & _
             FormatCurrency(incomeGT - costGT).PadLeft(13))
         sb.Append(vbCrLf)
         oPS = New CPrintStringNew()
         oPS.TitleFontStyle = "BI"
         oPS.TitleFontSize = REPORT_TITLE_FONT_SIZE
         If _Preview Then
            oPS.PrintPreview(120, sb.ToString, _
            ReportName, _
            "Total Cost of Ownership Report", _
            colHdr1:=colHdr, Landscape:=True)
         Else
            oPS.StartPrint(120, sb.ToString, _
               ReportName, _
               "Total Cost of Ownership Report", _
               colHdr1:=colHdr, Landscape:=True)
         End If
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try

   End Sub

#End Region



End Class
