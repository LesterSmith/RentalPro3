''' Purpose:
''' ''' Checks once a day for reminders on equipment
''' due to come in.
''' Author:  Les Smith
''' Date Created: 10/07/2003 at 08:59:16
''' CopyRight:  HHI Software
'''
Imports System.Text

Public Class CReminder
   Dim sql As String
   Dim oDA As New CDataAccess()
   Dim oCG As New CGrid()

   ''' <summary>
   ''' Check to see if we have run the reminder already.
   ''' </summary>
   Public Sub CheckReminder()
      Dim serr As String = String.Empty
      Dim dt As New DataTable("dt")
      Try
         sql = "select reminder_date from configuration"
         If oDA.SendQuery(sql, dt, ConnectString) > 0 Then
            If Not IsDBNull(dt.Rows(0).Item("reminder_date")) Then
               Dim dv As Date = DateValue(dt.Rows(0).Item("reminder_date"))
               If dv < Today Then
                  Call Reminder()
                  sql = "update configuration set reminder_date = #" & Now.ToString & "#"
                  If oDA.SendActionSql(sql, ConnectString, serr) < 1 Then
                     Throw New System.Exception(serr)
                  End If
               Else
                  Exit Sub
               End If
            End If
         End If
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub
   ''' <summary>
   ''' Return byref datatable of items due to be returned today.
   ''' </summary>
   ''' <param name = "dt"></param>
   ''' <returns>Integer</returns>
   Public Function GetEquipDueInToday(ByRef dg As DataGrid, ByVal frm As frmReminder) As Integer
      Dim dt As New DataTable()
      Try
         sql = "select Equip_id,equip_name,rented_date,available_date,companyname,phonenumber "
         sql &= "from equipment,customers "
         sql &= "where available = 'ON RENT' "
         sql &= "and available_date <= #" & Today & "# "
         'sql &= "and available_date < #" & DateAdd(DateInterval.Day, 1, Today) & "# "
         sql &= "and equipment.renting_company_id=customers.customerid "
         sql &= "order by equip_name"
         oCG.ClearDataTableForRebinding(dt)
         If oDA.SendQuery(sql, dt, ConnectString, "dt") > 0 Then
            Dim formats() As String = _
               {"", "60", "T", "L", _
                "", "150", "T", "L", _
                "MM/dd/yyyy hh:mm tt", "130", "T", "L", _
                "MM/dd/yyyy hh:mm tt", "130", "T", "L", _
                "", "150", "T", "L", _
                "", "100", "T", "L"}
            oCG.SetTablesStyle(dt, dg, formats)
            oCG.BindDataTableToGrid(dt, dg)
            oCG.DisableAddNew(dg, frm)
         End If

      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Function
   ''' <summary>
   ''' Fill passed grid with rerents due to rent today or past due renting.
   ''' </summary>
   ''' <param name = "dg"></param>
   ''' <param name = "frm"></param>
   ''' <returns>Integer</returns>
   Public Function GetRerentsDue(ByRef dg As DataGrid, ByRef frm As frmReminder) As Integer
      Dim dt As New DataTable()
      Dim SQL As String = ""
      SQL = ""
      SQL &= "select Equip_Name,Customer_Name,Customer_Phone as  "
      SQL &= "phone,Date_Needed, Nbr_Periods,Period "
      SQL &= "from rerents "
      SQL &= "where isnull(rented_date) "
      SQL &= "and datevalue(date_needed) <= #" & Today & "# "
      SQL &= "order by equip_name "
      If oDA.SendQuery(SQL, dt, ConnectString, "dt") > 0 Then
         Dim formats() As String = _
            {"", "150", "T", "L", _
             "", "150", "T", "L", _
             "", "130", "T", "L", _
             "MM/dd/yyyy hh:mm tt", "130", "T", "L", _
             "", "60", "T", "R", _
             "", "60", "T", "L"}
         oCG.SetTablesStyle(dt, dg, formats)
         oCG.BindDataTableToGrid(dt, dg)
         oCG.DisableAddNew(dg, frm)
      End If
   End Function


   ''' <summary>
   ''' Return byref datatable of items due to be returned today.
   ''' </summary>
   ''' <param name = "dt"></param>
   ''' <returns>Integer</returns>
   Public Function GetEquipReservedToday(ByRef dg As DataGrid, ByVal frm As frmReminder) As Integer
      Dim dt As New DataTable()
      Try
         sql = "select equip_name,customer_name,num_periods,res_period,phone,res_end_date,contact,phone "
         sql &= "from reservations "
         sql &= "where res_date >= #" & Today & "# "
         'sql &= "and res_date < #" & DateAdd(DateInterval.Day, 1, Today) & "# "
         sql &= "order by equip_name"
         oCG.ClearDataTableForRebinding(dt)
         If oDA.SendQuery(sql, dt, ConnectString, "dt") > 0 Then
            Dim formats() As String = _
               {"", "150", "T", "L", _
                "", "150", "T", "L", _
                "", "60", "T", "R", _
                "", "80", "T", "L", _
                "", "150", "T", "L", _
                "MM/dd/yyyy hh:mm tt", "130", "T", "L", _
                "", "120", "T", "L", _
                "", "100", "T", "L"}
            oCG.SetTablesStyle(dt, dg, formats)
            oCG.BindDataTableToGrid(dt, dg)
            oCG.DisableAddNew(dg, frm)
         End If

      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Function
   ''' <summary>
   ''' Return dg filled with open invoices over 17
   ''' days and less < 18 days
   ''' </summary>
   ''' <param name = "dg"></param>
   ''' <param name = "frm"></param>
   ''' <returns>Integer</returns>
   Public Function GetInvoicesOver17Days(ByRef dg As DataGrid, ByVal frm As frmReminder) As Integer
      Dim dt As New DataTable()
      'sql = "select d.invoiceid, d.rented_date, d.equip_id, "
      'sql &= "c.companyname,e.equip_name "
      'sql &= "from invoice_details d,customers c,equipment e "
      'sql &= "where e.available ='ON RENT' "
      'sql &= "and d.customer_id = c.customerid "
      'sql &= "and d.equip_id = e.equip_id "

      'sql &= "order by d.rented_date "
      Try
         Dim SQL As String = ""
         SQL = ""
         SQL &= "select i.invoiceid,i.invoicedate,c.companyname,c.contactname "
         SQL &= "from invoices i, customers c "
         SQL &= "where i.status ='CheckedOut' "
         SQL &= "and i.customerID=c.customerid "
         SQL &= "order by i.invoicedate "

         oCG.ClearDataTableForRebinding(dt)
         If oDA.SendQuery(SQL, dt, ConnectString, "dt") = 0 Then
            Return 0
         End If
         Dim i As Integer
         Dim dr As DataRow
         Dim diff As Integer

         For i = dt.Rows.Count - 1 To 0 Step -1
            dr = dt.Rows(i)
            diff = DateDiff(DateInterval.Hour, dr("invoicedate"), Now)
            If Not ((diff > (24 * MonthlyBreakDays))) Then 'And diff < (24 * (MonthlyBreakDays + 1))) Then
               dt.Rows(i).Delete()
            End If
         Next i
         Dim formats() As String = _
            {"", "60", "T", "L", _
             "MM/dd/yyyy hh:mm tt", "130", "T", "L", _
             "", "150", "T", "L", _
             "", "150", "T", "L"}
         oCG.SetTablesStyle(dt, dg, formats)
         oCG.BindDataTableToGrid(dt, dg)
         oCG.DisableAddNew(dg, frm)
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Function

   ''' <summary>
   ''' Return dg filled with open invoices over 17
   ''' days and less < 18 days
   ''' </summary>
   ''' <param name = "dg"></param>
   ''' <param name = "frm"></param>
   ''' <returns>Integer</returns>
   Public Function GetInvoicesOver3Days(ByRef dg As DataGrid, ByVal frm As frmReminder) As Integer
      Dim dt As New DataTable()
      Try
         sql = "select d.invoiceid, d.rented_date, d.equip_id, "
         sql &= "c.companyname,e.equip_name "
         sql &= "from invoice_details d,customers c,equipment e "
         sql &= "where e.available ='ON RENT' "
         sql &= "and d.customer_id = c.customerid "
         sql &= "and d.equip_id = e.equip_id "
         sql &= "order by d.rented_date "
         oCG.ClearDataTableForRebinding(dt)
         If oDA.SendQuery(sql, dt, ConnectString, "dt") = 0 Then
            Return 0
         End If
         Dim i As Integer
         Dim dr As DataRow
         Dim diff As Integer

         For i = dt.Rows.Count - 1 To 0 Step -1
            dr = dt.Rows(i)
            diff = DateDiff(DateInterval.Hour, dr("rented_date"), Now)
            If Not (diff > (24 * WeeklyBreakDays)) Then ' And diff < (24 * (WeeklyBreakDays + 1))) Then
               dt.Rows(i).Delete()
            End If
         Next i
         Dim formats() As String = _
            {"", "60", "T", "L", _
             "MM/dd/yyyy hh:mm tt", "130", "T", "L", _
             "", "60", "T", "L", _
             "", "150", "T", "L", _
             "", "150", "T", "L"}
         oCG.SetTablesStyle(dt, dg, formats)
         oCG.BindDataTableToGrid(dt, dg)
         oCG.DisableAddNew(dg, frm)
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Function




   ''' <summary>
   ''' Check all equipment to see if anything is due this morning.
   ''' Loop through the equipment table of rented equipment
   ''' List any items that are due to be returned today.
   ''' List any items that are over three days and less than 4 days,
   ''' they should be invoiced at a week.
   ''' List any items that are over 17 days and less than 18 days, they
   ''' should be invoiced for a month.
   ''' </summary>
   Private Sub Reminder()
      Dim dt As New DataTable()
      Try
         sql = "select d.invoiceid,d.rented_date, d.equip_id, "
         sql &= "c.companyname,e.equip_name "
         sql &= "from invoice_details d,customers c,equipment e "
         sql &= "where e.available ='ON RENT' "
         'sql &= "((datediff(DateInterval.Hour,d.rented_date,today) > " & 24 * 3.5 & " "
         'sql &= "and datediff(DateInterval.Hour,d.rentd_date,today) < " & 24 * 4.5 & ") "
         'sql &= "or (datediff(DateInterval.Hour,d.rentd_date,today) > " & 24 * 17 & " "
         'sql &= "and DateInterval.Hour,d.rentd_date,today) < " & 24 * 18 & ")) "
         sql &= "and d.customer_id = c.customerid "
         sql &= "and d.equip_id = e.equip_id "
         sql &= "order by d.rented_date "
         dt.Reset()
         If oDA.SendQuery(sql, dt, ConnectString) > 0 Then
            Dim sb As New StringBuilder()
            Dim i As Integer
            Dim dr As DataRow
            Dim diff As Integer
            Dim colHdr As String = _
               "Invoice".PadRight(10) & _
               "Customer".PadRight(25) & _
               "Equip ID".PadRight(11) & _
               "Equip Name".PadRight(30) & _
               "Hrs".PadRight(6) & _
               "Action"

            For i = 0 To dt.Rows.Count - 1
               dr = dt.Rows(i)
               diff = DateDiff(DateInterval.Hour, dr("rented_date"), Today)
               If (diff > (24 * 3.5) And diff < (24 * 4.5)) Or _
                 (diff > (24 * 17) And diff < (24 * 18)) Then
                  sb.Append(MNS(dr("invoiceid").PadRight(10)))
                  sb.Append(LS(MNS(dr("companyname")), 24).PadRight(25))
                  sb.Append(MNS(dr("equip_id")).PadRight(11))
                  sb.Append(LS(MNS(dr("equip_name")), 29).PadRight(30))
                  sb.Append(Format(dr("diff"), "0.0").PadLeft(5) & " ")
                  If dr("diff") > (24 * 3.5) And dr("diff") < (24 * 4.5) Then
                     sb.Append("> 3.5 Days")
                  ElseIf dr("diff") > (24 * 17) And dr("diff") < (24 * 18) Then
                     sb.Append("> 17 Days")
                  End If
                  sb.Append(vbCrLf)
               End If
            Next
            Dim o As New CPrintStringNew()
            o.TitleFontSize = 16
            o.TitleFontStyle = "BI"
            o.PrintPreview(96, sb.ToString, ReportName, "Daily Reminder to Pre-Invoice", colHdr1:=colHdr)
         End If
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub


End Class
