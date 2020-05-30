Imports System.Windows.Forms.Application
Public Class frmSetup
   Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

   Public Sub New()
      MyBase.New()

      'This call is required by the Windows Form Designer.
      InitializeComponent()

      'Add any initialization after the InitializeComponent() call

   End Sub

   'Form overrides dispose to clean up the component list.
   Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
      If disposing Then
         If Not (components Is Nothing) Then
            components.Dispose()
         End If
      End If
      MyBase.Dispose(disposing)
   End Sub

   'Required by the Windows Form Designer
   Private components As System.ComponentModel.IContainer

   'NOTE: The following procedure is required by the Windows Form Designer
   'It can be modified using the Windows Form Designer.  
   'Do not modify it using the code editor.
   Friend WithEvents txtReportName As System.Windows.Forms.TextBox
   Friend WithEvents txtCorporateName As System.Windows.Forms.TextBox
   Friend WithEvents txtAddress1 As System.Windows.Forms.TextBox
   Friend WithEvents txtAddress2 As System.Windows.Forms.TextBox
   Friend WithEvents txtCity As System.Windows.Forms.TextBox
   Friend WithEvents txtState As System.Windows.Forms.TextBox
   Friend WithEvents txtZip As System.Windows.Forms.TextBox
   Friend WithEvents txtPhone As System.Windows.Forms.TextBox
   Friend WithEvents txtFax As System.Windows.Forms.TextBox
   Friend WithEvents txtEmail As System.Windows.Forms.TextBox
   Friend WithEvents txtTaxRate As System.Windows.Forms.TextBox
   Friend WithEvents cbAcctBasis As System.Windows.Forms.ComboBox
   Friend WithEvents chkUseDeposits As System.Windows.Forms.CheckBox
   Friend WithEvents chkUseHourlyRates As System.Windows.Forms.CheckBox
   Friend WithEvents Label1 As System.Windows.Forms.Label
   Friend WithEvents Label2 As System.Windows.Forms.Label
   Friend WithEvents Label3 As System.Windows.Forms.Label
   Friend WithEvents Label4 As System.Windows.Forms.Label
   Friend WithEvents Label5 As System.Windows.Forms.Label
   Friend WithEvents Label6 As System.Windows.Forms.Label
   Friend WithEvents Label7 As System.Windows.Forms.Label
   Friend WithEvents Label8 As System.Windows.Forms.Label
   Friend WithEvents Label9 As System.Windows.Forms.Label
   Friend WithEvents Label10 As System.Windows.Forms.Label
   Friend WithEvents Label11 As System.Windows.Forms.Label
   Friend WithEvents Label12 As System.Windows.Forms.Label
   Friend WithEvents btnSave As System.Windows.Forms.Button
   Friend WithEvents btnCancel As System.Windows.Forms.Button
   Friend WithEvents ckInitialsOnly As System.Windows.Forms.CheckBox
   Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
   Friend WithEvents ckCalcByMonth As System.Windows.Forms.CheckBox
   Friend WithEvents textDaysPerMonth As System.Windows.Forms.TextBox
   Friend WithEvents Label13 As System.Windows.Forms.Label
   Friend WithEvents Label14 As System.Windows.Forms.Label
   Friend WithEvents textHoursPerMonth As System.Windows.Forms.TextBox
   Friend WithEvents textGraceHrsForDay As System.Windows.Forms.TextBox
   Friend WithEvents Label15 As System.Windows.Forms.Label
   Friend WithEvents textGraceHoursForHalfDay As System.Windows.Forms.TextBox
   Friend WithEvents Label16 As System.Windows.Forms.Label
   Friend WithEvents ckUseHalfDays As System.Windows.Forms.CheckBox
   Friend WithEvents ckAutoCalc As System.Windows.Forms.CheckBox
   Friend WithEvents textMonthlyBreakDays As System.Windows.Forms.TextBox
   Friend WithEvents Label17 As System.Windows.Forms.Label
   Friend WithEvents Label18 As System.Windows.Forms.Label
   Friend WithEvents textWeeklyBreakDays As System.Windows.Forms.TextBox
    Friend WithEvents chkWeekendRates As System.Windows.Forms.CheckBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents txtEmailServer As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents txtEmailBody As System.Windows.Forms.TextBox
    Friend WithEvents txtCutePDFFilePath As System.Windows.Forms.TextBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents txtEmailSubject As System.Windows.Forms.TextBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents txtEmailPort As System.Windows.Forms.TextBox
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents chkEmailSSL As System.Windows.Forms.CheckBox
   Friend WithEvents chkCalcBestRates As System.Windows.Forms.CheckBox
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSetup))
        Me.txtReportName = New System.Windows.Forms.TextBox()
        Me.txtCorporateName = New System.Windows.Forms.TextBox()
        Me.txtAddress1 = New System.Windows.Forms.TextBox()
        Me.txtAddress2 = New System.Windows.Forms.TextBox()
        Me.txtCity = New System.Windows.Forms.TextBox()
        Me.txtState = New System.Windows.Forms.TextBox()
        Me.txtZip = New System.Windows.Forms.TextBox()
        Me.txtPhone = New System.Windows.Forms.TextBox()
        Me.txtFax = New System.Windows.Forms.TextBox()
        Me.txtEmail = New System.Windows.Forms.TextBox()
        Me.txtTaxRate = New System.Windows.Forms.TextBox()
        Me.cbAcctBasis = New System.Windows.Forms.ComboBox()
        Me.chkUseDeposits = New System.Windows.Forms.CheckBox()
        Me.chkUseHourlyRates = New System.Windows.Forms.CheckBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.ckInitialsOnly = New System.Windows.Forms.CheckBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ckCalcByMonth = New System.Windows.Forms.CheckBox()
        Me.textDaysPerMonth = New System.Windows.Forms.TextBox()
        Me.textHoursPerMonth = New System.Windows.Forms.TextBox()
        Me.textGraceHrsForDay = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.ckUseHalfDays = New System.Windows.Forms.CheckBox()
        Me.ckAutoCalc = New System.Windows.Forms.CheckBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.textGraceHoursForHalfDay = New System.Windows.Forms.TextBox()
        Me.textMonthlyBreakDays = New System.Windows.Forms.TextBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.textWeeklyBreakDays = New System.Windows.Forms.TextBox()
        Me.chkWeekendRates = New System.Windows.Forms.CheckBox()
        Me.chkCalcBestRates = New System.Windows.Forms.CheckBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.txtEmailServer = New System.Windows.Forms.TextBox()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.txtEmailBody = New System.Windows.Forms.TextBox()
        Me.txtCutePDFFilePath = New System.Windows.Forms.TextBox()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.txtEmailSubject = New System.Windows.Forms.TextBox()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.txtEmailPort = New System.Windows.Forms.TextBox()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.chkEmailSSL = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'txtReportName
        '
        Me.txtReportName.Location = New System.Drawing.Point(114, 8)
        Me.txtReportName.Name = "txtReportName"
        Me.txtReportName.Size = New System.Drawing.Size(144, 20)
        Me.txtReportName.TabIndex = 0
        '
        'txtCorporateName
        '
        Me.txtCorporateName.Location = New System.Drawing.Point(114, 34)
        Me.txtCorporateName.Name = "txtCorporateName"
        Me.txtCorporateName.Size = New System.Drawing.Size(160, 20)
        Me.txtCorporateName.TabIndex = 1
        '
        'txtAddress1
        '
        Me.txtAddress1.Location = New System.Drawing.Point(114, 60)
        Me.txtAddress1.Name = "txtAddress1"
        Me.txtAddress1.Size = New System.Drawing.Size(160, 20)
        Me.txtAddress1.TabIndex = 2
        '
        'txtAddress2
        '
        Me.txtAddress2.Location = New System.Drawing.Point(114, 86)
        Me.txtAddress2.Name = "txtAddress2"
        Me.txtAddress2.Size = New System.Drawing.Size(160, 20)
        Me.txtAddress2.TabIndex = 3
        '
        'txtCity
        '
        Me.txtCity.Location = New System.Drawing.Point(114, 112)
        Me.txtCity.Name = "txtCity"
        Me.txtCity.Size = New System.Drawing.Size(96, 20)
        Me.txtCity.TabIndex = 4
        '
        'txtState
        '
        Me.txtState.Location = New System.Drawing.Point(114, 138)
        Me.txtState.Name = "txtState"
        Me.txtState.Size = New System.Drawing.Size(88, 20)
        Me.txtState.TabIndex = 5
        '
        'txtZip
        '
        Me.txtZip.Location = New System.Drawing.Point(114, 164)
        Me.txtZip.Name = "txtZip"
        Me.txtZip.Size = New System.Drawing.Size(96, 20)
        Me.txtZip.TabIndex = 6
        '
        'txtPhone
        '
        Me.txtPhone.Location = New System.Drawing.Point(114, 190)
        Me.txtPhone.Name = "txtPhone"
        Me.txtPhone.Size = New System.Drawing.Size(96, 20)
        Me.txtPhone.TabIndex = 7
        '
        'txtFax
        '
        Me.txtFax.Location = New System.Drawing.Point(114, 216)
        Me.txtFax.Name = "txtFax"
        Me.txtFax.Size = New System.Drawing.Size(104, 20)
        Me.txtFax.TabIndex = 8
        '
        'txtEmail
        '
        Me.txtEmail.Location = New System.Drawing.Point(114, 242)
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.Size = New System.Drawing.Size(160, 20)
        Me.txtEmail.TabIndex = 9
        '
        'txtTaxRate
        '
        Me.txtTaxRate.Location = New System.Drawing.Point(113, 268)
        Me.txtTaxRate.Name = "txtTaxRate"
        Me.txtTaxRate.Size = New System.Drawing.Size(72, 20)
        Me.txtTaxRate.TabIndex = 10
        Me.ToolTip1.SetToolTip(Me.txtTaxRate, "Total Tax Rate to be charged")
        '
        'cbAcctBasis
        '
        Me.cbAcctBasis.Items.AddRange(New Object() {"CASH", "ACCURAL"})
        Me.cbAcctBasis.Location = New System.Drawing.Point(113, 294)
        Me.cbAcctBasis.Name = "cbAcctBasis"
        Me.cbAcctBasis.Size = New System.Drawing.Size(88, 21)
        Me.cbAcctBasis.TabIndex = 11
        Me.ToolTip1.SetToolTip(Me.cbAcctBasis, "Cash or Accrual basis")
        '
        'chkUseDeposits
        '
        Me.chkUseDeposits.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkUseDeposits.Location = New System.Drawing.Point(360, 115)
        Me.chkUseDeposits.Name = "chkUseDeposits"
        Me.chkUseDeposits.Size = New System.Drawing.Size(96, 16)
        Me.chkUseDeposits.TabIndex = 12
        Me.chkUseDeposits.Text = "Use Deposits"
        Me.ToolTip1.SetToolTip(Me.chkUseDeposits, "Check to use deposits")
        '
        'chkUseHourlyRates
        '
        Me.chkUseHourlyRates.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkUseHourlyRates.Location = New System.Drawing.Point(344, 136)
        Me.chkUseHourlyRates.Name = "chkUseHourlyRates"
        Me.chkUseHourlyRates.Size = New System.Drawing.Size(112, 16)
        Me.chkUseHourlyRates.TabIndex = 13
        Me.chkUseHourlyRates.Text = "Use Hourly Rates"
        Me.ToolTip1.SetToolTip(Me.chkUseHourlyRates, "Allow the use of hourly rates")
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(28, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(70, 13)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "Report Name"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(11, 34)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(84, 13)
        Me.Label2.TabIndex = 15
        Me.Label2.Text = "Corporate Name"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(47, 60)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(51, 13)
        Me.Label3.TabIndex = 16
        Me.Label3.Text = "Address1"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(47, 86)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(51, 13)
        Me.Label4.TabIndex = 17
        Me.Label4.Text = "Address2"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(75, 112)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(24, 13)
        Me.Label5.TabIndex = 18
        Me.Label5.Text = "City"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(68, 138)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(32, 13)
        Me.Label6.TabIndex = 19
        Me.Label6.Text = "State"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(79, 164)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(22, 13)
        Me.Label7.TabIndex = 20
        Me.Label7.Text = "Zip"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(62, 190)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(38, 13)
        Me.Label8.TabIndex = 21
        Me.Label8.Text = "Phone"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(76, 216)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(24, 13)
        Me.Label9.TabIndex = 22
        Me.Label9.Text = "Fax"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(66, 242)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(33, 13)
        Me.Label10.TabIndex = 23
        Me.Label10.Text = "EMail"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(50, 268)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(51, 13)
        Me.Label11.TabIndex = 24
        Me.Label11.Text = "Tax Rate"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(18, 294)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(89, 13)
        Me.Label12.TabIndex = 25
        Me.Label12.Text = "Accounting Basis"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(320, 508)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(74, 24)
        Me.btnSave.TabIndex = 26
        Me.btnSave.Text = "&Save"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(408, 508)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(64, 24)
        Me.btnCancel.TabIndex = 27
        Me.btnCancel.Text = "&Cancel"
        '
        'ckInitialsOnly
        '
        Me.ckInitialsOnly.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ckInitialsOnly.Checked = True
        Me.ckInitialsOnly.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ckInitialsOnly.Location = New System.Drawing.Point(312, 158)
        Me.ckInitialsOnly.Name = "ckInitialsOnly"
        Me.ckInitialsOnly.Size = New System.Drawing.Size(144, 16)
        Me.ckInitialsOnly.TabIndex = 28
        Me.ckInitialsOnly.Text = "Print Emp Initials Only"
        Me.ToolTip1.SetToolTip(Me.ckInitialsOnly, "Print initials only on invoice if checked, otherwise whole name")
        '
        'ckCalcByMonth
        '
        Me.ckCalcByMonth.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ckCalcByMonth.Location = New System.Drawing.Point(307, 174)
        Me.ckCalcByMonth.Name = "ckCalcByMonth"
        Me.ckCalcByMonth.Size = New System.Drawing.Size(149, 35)
        Me.ckCalcByMonth.TabIndex = 29
        Me.ckCalcByMonth.Text = "Calc by Month (Unchecked, use Days)"
        Me.ckCalcByMonth.TextAlign = System.Drawing.ContentAlignment.BottomRight
        Me.ToolTip1.SetToolTip(Me.ckCalcByMonth, "Check for Monthly Calc or uncheck for Days/Month calc")
        '
        'textDaysPerMonth
        '
        Me.textDaysPerMonth.Location = New System.Drawing.Point(443, 63)
        Me.textDaysPerMonth.MaxLength = 2
        Me.textDaysPerMonth.Name = "textDaysPerMonth"
        Me.textDaysPerMonth.Size = New System.Drawing.Size(20, 20)
        Me.textDaysPerMonth.TabIndex = 30
        Me.textDaysPerMonth.Text = "30"
        Me.ToolTip1.SetToolTip(Me.textDaysPerMonth, "Days for month usage if not calc by month")
        '
        'textHoursPerMonth
        '
        Me.textHoursPerMonth.Location = New System.Drawing.Point(443, 87)
        Me.textHoursPerMonth.MaxLength = 160
        Me.textHoursPerMonth.Name = "textHoursPerMonth"
        Me.textHoursPerMonth.Size = New System.Drawing.Size(28, 20)
        Me.textHoursPerMonth.TabIndex = 33
        Me.textHoursPerMonth.Text = "160"
        Me.ToolTip1.SetToolTip(Me.textHoursPerMonth, "Meter usage hours allowed in month")
        '
        'textGraceHrsForDay
        '
        Me.textGraceHrsForDay.Location = New System.Drawing.Point(443, 5)
        Me.textGraceHrsForDay.MaxLength = 2
        Me.textGraceHrsForDay.Name = "textGraceHrsForDay"
        Me.textGraceHrsForDay.Size = New System.Drawing.Size(20, 20)
        Me.textGraceHrsForDay.TabIndex = 34
        Me.textGraceHrsForDay.Text = "2"
        Me.ToolTip1.SetToolTip(Me.textGraceHrsForDay, "Number of hours check-in can be late")
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(304, 36)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(133, 13)
        Me.Label16.TabIndex = 37
        Me.Label16.Text = "Grace Hours for Half Day"
        Me.ToolTip1.SetToolTip(Me.Label16, "Number of hours check-in can be late")
        '
        'ckUseHalfDays
        '
        Me.ckUseHalfDays.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ckUseHalfDays.Location = New System.Drawing.Point(351, 213)
        Me.ckUseHalfDays.Name = "ckUseHalfDays"
        Me.ckUseHalfDays.Size = New System.Drawing.Size(104, 16)
        Me.ckUseHalfDays.TabIndex = 38
        Me.ckUseHalfDays.Text = "Use Half Days"
        Me.ToolTip1.SetToolTip(Me.ckUseHalfDays, "Use Half Day Charges (Unchecked use 4 hours)")
        '
        'ckAutoCalc
        '
        Me.ckAutoCalc.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ckAutoCalc.Checked = True
        Me.ckAutoCalc.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ckAutoCalc.Location = New System.Drawing.Point(383, 235)
        Me.ckAutoCalc.Name = "ckAutoCalc"
        Me.ckAutoCalc.Size = New System.Drawing.Size(72, 16)
        Me.ckAutoCalc.TabIndex = 39
        Me.ckAutoCalc.Text = "Auto Calc"
        Me.ToolTip1.SetToolTip(Me.ckAutoCalc, "Auto calculate the charges at check-in")
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(287, 66)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(151, 14)
        Me.Label13.TabIndex = 31
        Me.Label13.Text = "Days/Month (if Calc by Days)"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(368, 90)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(70, 13)
        Me.Label14.TabIndex = 32
        Me.Label14.Text = "Hours/Month"
        '
        'Label15
        '
        Me.Label15.Location = New System.Drawing.Point(304, 8)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(121, 13)
        Me.Label15.TabIndex = 35
        Me.Label15.Text = "Grace Hours for a Day"
        '
        'textGraceHoursForHalfDay
        '
        Me.textGraceHoursForHalfDay.Location = New System.Drawing.Point(443, 35)
        Me.textGraceHoursForHalfDay.MaxLength = 2
        Me.textGraceHoursForHalfDay.Name = "textGraceHoursForHalfDay"
        Me.textGraceHoursForHalfDay.Size = New System.Drawing.Size(20, 20)
        Me.textGraceHoursForHalfDay.TabIndex = 36
        Me.textGraceHoursForHalfDay.Text = "1"
        '
        'textMonthlyBreakDays
        '
        Me.textMonthlyBreakDays.Location = New System.Drawing.Point(440, 257)
        Me.textMonthlyBreakDays.MaxLength = 2
        Me.textMonthlyBreakDays.Name = "textMonthlyBreakDays"
        Me.textMonthlyBreakDays.Size = New System.Drawing.Size(24, 20)
        Me.textMonthlyBreakDays.TabIndex = 40
        Me.textMonthlyBreakDays.Tag = "(No Auto Formatting)"
        Me.textMonthlyBreakDays.Text = "17"
        Me.textMonthlyBreakDays.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(324, 259)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(102, 13)
        Me.Label17.TabIndex = 41
        Me.Label17.Text = "Monthly Break Days"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(319, 280)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(101, 13)
        Me.Label18.TabIndex = 43
        Me.Label18.Text = "Weekly Break Days"
        '
        'textWeeklyBreakDays
        '
        Me.textWeeklyBreakDays.Location = New System.Drawing.Point(439, 280)
        Me.textWeeklyBreakDays.MaxLength = 4
        Me.textWeeklyBreakDays.Name = "textWeeklyBreakDays"
        Me.textWeeklyBreakDays.Size = New System.Drawing.Size(27, 20)
        Me.textWeeklyBreakDays.TabIndex = 42
        Me.textWeeklyBreakDays.Tag = "(No Auto Formatting)"
        Me.textWeeklyBreakDays.Text = "3.0"
        Me.textWeeklyBreakDays.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'chkWeekendRates
        '
        Me.chkWeekendRates.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkWeekendRates.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.chkWeekendRates.Location = New System.Drawing.Point(312, 305)
        Me.chkWeekendRates.Name = "chkWeekendRates"
        Me.chkWeekendRates.Size = New System.Drawing.Size(140, 16)
        Me.chkWeekendRates.TabIndex = 44
        Me.chkWeekendRates.Text = "Use WeekEnd Rates"
        '
        'chkCalcBestRates
        '
        Me.chkCalcBestRates.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkCalcBestRates.Location = New System.Drawing.Point(347, 326)
        Me.chkCalcBestRates.Name = "chkCalcBestRates"
        Me.chkCalcBestRates.Size = New System.Drawing.Size(105, 18)
        Me.chkCalcBestRates.TabIndex = 45
        Me.chkCalcBestRates.Text = "Calc Best Rate"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(40, 322)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(66, 13)
        Me.Label19.TabIndex = 46
        Me.Label19.Text = "Email Server"
        '
        'txtEmailServer
        '
        Me.txtEmailServer.Location = New System.Drawing.Point(114, 322)
        Me.txtEmailServer.Name = "txtEmailServer"
        Me.txtEmailServer.Size = New System.Drawing.Size(122, 20)
        Me.txtEmailServer.TabIndex = 47
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(47, 407)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(59, 13)
        Me.Label21.TabIndex = 49
        Me.Label21.Text = "Email Body"
        '
        'txtEmailBody
        '
        Me.txtEmailBody.Location = New System.Drawing.Point(114, 405)
        Me.txtEmailBody.Multiline = True
        Me.txtEmailBody.Name = "txtEmailBody"
        Me.txtEmailBody.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtEmailBody.Size = New System.Drawing.Size(338, 94)
        Me.txtEmailBody.TabIndex = 50
        '
        'txtCutePDFFilePath
        '
        Me.txtCutePDFFilePath.Location = New System.Drawing.Point(113, 348)
        Me.txtCutePDFFilePath.Name = "txtCutePDFFilePath"
        Me.txtCutePDFFilePath.Size = New System.Drawing.Size(339, 20)
        Me.txtCutePDFFilePath.TabIndex = 52
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(25, 348)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(84, 13)
        Me.Label20.TabIndex = 51
        Me.Label20.Text = "File to Print Path"
        '
        'txtEmailSubject
        '
        Me.txtEmailSubject.Location = New System.Drawing.Point(114, 374)
        Me.txtEmailSubject.Name = "txtEmailSubject"
        Me.txtEmailSubject.Size = New System.Drawing.Size(122, 20)
        Me.txtEmailSubject.TabIndex = 54
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(40, 374)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(71, 13)
        Me.Label22.TabIndex = 53
        Me.Label22.Text = "Email Subject"
        '
        'txtEmailPort
        '
        Me.txtEmailPort.Location = New System.Drawing.Point(297, 372)
        Me.txtEmailPort.Name = "txtEmailPort"
        Me.txtEmailPort.Size = New System.Drawing.Size(32, 20)
        Me.txtEmailPort.TabIndex = 56
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(241, 372)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(54, 13)
        Me.Label23.TabIndex = 55
        Me.Label23.Text = "Email Port"
        '
        'chkEmailSSL
        '
        Me.chkEmailSSL.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkEmailSSL.Location = New System.Drawing.Point(375, 372)
        Me.chkEmailSSL.Name = "chkEmailSSL"
        Me.chkEmailSSL.Size = New System.Drawing.Size(76, 18)
        Me.chkEmailSSL.TabIndex = 57
        Me.chkEmailSSL.Text = "Email SSL"
        '
        'frmSetup
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(490, 537)
        Me.Controls.Add(Me.chkEmailSSL)
        Me.Controls.Add(Me.txtEmailPort)
        Me.Controls.Add(Me.Label23)
        Me.Controls.Add(Me.txtEmailSubject)
        Me.Controls.Add(Me.Label22)
        Me.Controls.Add(Me.txtCutePDFFilePath)
        Me.Controls.Add(Me.Label20)
        Me.Controls.Add(Me.txtEmailBody)
        Me.Controls.Add(Me.Label21)
        Me.Controls.Add(Me.txtEmailServer)
        Me.Controls.Add(Me.Label19)
        Me.Controls.Add(Me.chkCalcBestRates)
        Me.Controls.Add(Me.chkWeekendRates)
        Me.Controls.Add(Me.Label18)
        Me.Controls.Add(Me.textWeeklyBreakDays)
        Me.Controls.Add(Me.Label17)
        Me.Controls.Add(Me.textMonthlyBreakDays)
        Me.Controls.Add(Me.ckAutoCalc)
        Me.Controls.Add(Me.ckUseHalfDays)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.textGraceHoursForHalfDay)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.textGraceHrsForDay)
        Me.Controls.Add(Me.textHoursPerMonth)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.textDaysPerMonth)
        Me.Controls.Add(Me.ckCalcByMonth)
        Me.Controls.Add(Me.ckInitialsOnly)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.chkUseHourlyRates)
        Me.Controls.Add(Me.chkUseDeposits)
        Me.Controls.Add(Me.cbAcctBasis)
        Me.Controls.Add(Me.txtTaxRate)
        Me.Controls.Add(Me.txtEmail)
        Me.Controls.Add(Me.txtFax)
        Me.Controls.Add(Me.txtPhone)
        Me.Controls.Add(Me.txtZip)
        Me.Controls.Add(Me.txtState)
        Me.Controls.Add(Me.txtCity)
        Me.Controls.Add(Me.txtAddress2)
        Me.Controls.Add(Me.txtAddress1)
        Me.Controls.Add(Me.txtCorporateName)
        Me.Controls.Add(Me.txtReportName)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmSetup"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Setup Configuration"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region
   Dim oCF As New CConfig()

   Private Sub frmSetup_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
      oCF.FillConfigTextBoxes(Me)
   End Sub

   Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
      oCF.SaveConfig(Me)
      Me.Close()
      DoEvents()
   End Sub

   Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
      Me.Close()
      DoEvents()
   End Sub
   Private Sub textMonthlyBreakDays_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles textMonthlyBreakDays.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
   End Sub
   Private Sub textMonthlyBreakDays_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles textMonthlyBreakDays.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub textMonthlyBreakDays_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles textMonthlyBreakDays.Enter
      textMonthlyBreakDays.SelectionStart = 0
      textMonthlyBreakDays.SelectionLength = textMonthlyBreakDays.Text.Trim.Length
   End Sub
   Private Sub textWeeklyBreakDays_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles textWeeklyBreakDays.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
   End Sub
   Private Sub textWeeklyBreakDays_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles textWeeklyBreakDays.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub textWeeklyBreakDays_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles textWeeklyBreakDays.Enter
      textWeeklyBreakDays.SelectionStart = 0
      textWeeklyBreakDays.SelectionLength = textWeeklyBreakDays.Text.Trim.Length
   End Sub

    Private Sub Label17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label17.Click

    End Sub

    Private Sub textWeeklyBreakDays_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles textWeeklyBreakDays.TextChanged

    End Sub

    Private Sub Label18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label18.Click

    End Sub
End Class
