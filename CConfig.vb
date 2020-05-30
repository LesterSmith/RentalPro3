Public Class CConfig
    Private oDA As CDataAccess
#Region " Public Methods "
    ''' <summary>
    ''' Save the changes to the config table.
    ''' </summary>
    ''' <param name = "frm"></param>
    Public Sub SaveConfig(ByRef frm As frmSetup)
        Dim sql As New System.Text.StringBuilder()
        Dim sErr As String

        Try
            With frm
                sql.Append("update configuration set ")
                sql.Append("report_name = '" & .txtReportName.Text & "', ")
                sql.Append("corporate_name = '" & .txtCorporateName.Text & "', ")
                sql.Append("tax_rate = " & .txtTaxRate.Text & ", ")
                sql.Append("address1 = '" & .txtAddress1.Text & "', ")
                sql.Append("address2 = '" & .txtAddress2.Text & "', ")
                sql.Append("city  = '" & .txtCity.Text & "', ")
                sql.Append("state = '" & .txtState.Text & "', ")
                sql.Append("zip =  '" & .txtZip.Text & "', ")
                sql.Append("phone  = '" & .txtPhone.Text & "', ")
                sql.Append("fax = '" & .txtFax.Text & "', ")
                sql.Append("email = '" & .txtEmail.Text & "', ")
                sql.Append("emailserver = '" & .txtEmailServer.Text & "', ")
                sql.Append("emailbody = '" & .txtEmailBody.Text & "', ")
                sql.Append("use_deposits = " & .chkUseDeposits.Checked & ", ")
                sql.Append("use_hourly_rates = " & IIf(.chkUseHourlyRates.Checked, True, False) & ", ")
                sql.Append("accounting_basis = '" & .cbAcctBasis.Text & "',")
                sql.Append("print_employee_initials = " & IIf(.ckInitialsOnly.Checked, True, False) & ", ")
                sql.Append("hours_per_month = " & Val(.textHoursPerMonth.Text) & ", ")
                sql.Append("calc_by_month = " & IIf(.ckCalcByMonth.Checked, True, False) & ", ")
                sql.Append("days_per_month = " & Val(.textDaysPerMonth.Text) & ", ")
                sql.Append("auto_calc = " & IIf(.ckAutoCalc.Checked, True, False) & ",")
                sql.Append("grace_hours_half_day = '" & .textGraceHoursForHalfDay.Text & "', ")
                sql.Append("grace_hours_day = '" & .textGraceHrsForDay.Text & "', ")
                sql.Append("use_half_days = " & IIf(.ckUseHalfDays.Checked, True, False) & " ")
                sql.Append(", monthly_break_days = " & .textMonthlyBreakDays.Text & " ")
                sql.Append(", weekly_break_days = " & .textWeeklyBreakDays.Text & ", ")
                sql.Append("calc_best_rate = " & IIf(.chkCalcBestRates.Checked, True, False) & ", ")
                sql.Append("use_weekend_rates = " & IIf(.chkWeekendRates.Checked, True, False) & ", ")
                sql.Append("EmailServer = '" & .txtEmailServer.Text & "', ")
                sql.Append("EmailBody = '" & .txtEmailBody.Text & "', ")
                sql.Append("EmailSubject = '" & .txtEmailSubject.Text & "', ")
                sql.Append("EmailPort = " & .txtEmailPort.Text & ", ")
                sql.Append("EmailSSL = " & IIf(.chkEmailSSL.Checked, True, False) & " ")
            End With
            If oDA.SendActionSql(sql.ToString, ConnectString, sErr) = 0 Then
                Throw New Exception("Update of configuration parameters failed.")
            End If
            Me.GetConfig()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Fill the setup form textboxes.
    ''' </summary>
    ''' <param name = "frm"></param>
    Public Sub FillConfigTextBoxes(ByRef frm As frmSetup)
        Dim sql As String
        Dim dt As New DataTable()
        Try
            sql = "select * from configuration"
            If oDA.SendQuery(sql, dt, ConnectString) > 0 Then
                With dt.Rows(0)
                    Dim dr As DataRow = dt.Rows(0)
                    frm.txtReportName.Text = dr("report_name")
                    frm.txtCorporateName.Text = dr("corporate_name")
                    frm.txtTaxRate.Text = dr("tax_rate")
                    frm.txtAddress1.Text = dr("address1")
                    frm.txtAddress2.Text = MNS(dr("address2"))
                    frm.txtCity.Text = MNS(dr("city"))
                    frm.txtState.Text = MNS(dr("state"))
                    frm.txtZip.Text = MNS(dr("zip"))
                    frm.txtPhone.Text = MNS(dr("phone"))
                    frm.txtFax.Text = MNS(dr("fax"))
                    frm.txtEmail.Text = MNS(dr("email"))
                    frm.chkUseDeposits.Checked = dr("use_deposits")
                    frm.chkUseHourlyRates.Checked = dr("use_hourly_rates")
                    frm.cbAcctBasis.Text = MNS(dr("accounting_basis"))
                    frm.ckInitialsOnly.Checked = dr("print_employee_initials")
                    frm.ckCalcByMonth.Checked = dr("calc_by_month")
                    frm.textDaysPerMonth.Text = MNI(dr("days_per_month"))
                    frm.textHoursPerMonth.Text = MNI(dr("hours_per_month"))
                    frm.textGraceHoursForHalfDay.Text = MNI(dr("grace_hours_Half_day"))
                    frm.textGraceHrsForDay.Text = MNI(dr("grace_hours_day"))
                    frm.ckAutoCalc.Checked = dr("auto_calc")
                    frm.ckUseHalfDays.Checked = dr("use_half_days")
                    frm.textMonthlyBreakDays.Text = MNI(dr("monthly_break_days"))
                    frm.textWeeklyBreakDays.Text = MNI(dr("weekly_break_days"))
                    frm.chkCalcBestRates.Checked = dr("calc_best_rate")
                    frm.chkWeekendRates.Checked = dr("use_weekend_rates")
                    frm.txtEmailServer.Text = MNS(dr("emailserver"))
                    frm.txtEmailSubject.Text = MNS(dr("EmailSubject"))
                    frm.txtEmailPort.Text = MNS(dr("EmailPort"))
                    frm.chkEmailSSL.Checked = MNB(dr("EmailSSL"))
                    frm.txtEmailBody.Text = MNS(dr("EmailBody"))
                End With
            End If

        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Fill the configuration variables.
    ''' </summary>
    Public Sub GetConfig()

        Dim sql As String
        Dim dt As New DataTable()

        Try
            sql = "select * from configuration"
            If oDA.SendQuery(sql, dt, ConnectString) > 0 Then
                With dt.Rows(0)
                    ReportName = .Item("report_name")
                    CorporateName = .Item("corporate_name")
                    TaxRate = .Item("tax_rate")
                    Address1 = .Item("address1")
                    Address2 = MNS(.Item("address2"))
                    City = MNS(.Item("city"))
                    State = MNS(.Item("state"))
                    Zip = MNS(.Item("zip"))
                    Phone = MNS(.Item("phone"))
                    Fax = MNS(.Item("fax"))
                    EMail = MNS(.Item("email"))
                    UseDeposits = .Item("use_deposits")
                    UseHourlyRates = .Item("use_hourly_rates")
                    AccountingBasis = MNS(.Item("accounting_basis"))
                    PrintInitialsOnly = .Item("print_employee_initials")
                    HoursPerMonth = MNI(.Item("hours_per_month"))
                    DaysPerMonth = MNI(.Item("days_per_month"))
                    CalcByMonth = .Item("calc_by_month")
                    GraceHoursForHalfDayRent = MNI(.Item("grace_hours_Half_day"))
                    GraceHoursForDayRent = MNI(.Item("grace_hours_day"))
                    AutoCalcOn = .Item("auto_calc")
                    UseHalfDays = .Item("use_half_days")
                    MonthlyBreakDays = MNI(.Item("monthly_break_days"))
                    WeeklyBreakDays = MNI(.Item("weekly_break_days"))
                    CalcBestRate = .Item("calc_best_rate")
                    UseWeekEndRates = .Item("use_weekend_rates")
                    EmailServer = MNS(.Item("EmailServer"))
                    EmailBody = MNS(.Item("EmailBody"))
                    EmailSubject = MNS(.Item("EmailSubject"))
                    CutePDFFilePath = MNS(.Item("CutePDFFilePath"))
                    EmailSSL = MNI(.Item("EmailSSL"))
                    EmailPort = MNB(.Item("EmailPort"))
                End With
            End If

        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

#End Region




#Region " Constructor "
    Public Sub New()
        oDA = New CDataAccess()
    End Sub

#End Region

End Class
