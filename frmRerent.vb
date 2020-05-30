Public Class frmRerent
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
   Public WithEvents _lblFieldLable_0 As System.Windows.Forms.Label
   Friend WithEvents btnAdd As System.Windows.Forms.Button
   Friend WithEvents btnClose As System.Windows.Forms.Button
   Friend WithEvents btnDelete As System.Windows.Forms.Button
   Friend WithEvents btnRent As System.Windows.Forms.Button
   Friend WithEvents btnSave As System.Windows.Forms.Button
   Friend WithEvents cbCustomers As System.Windows.Forms.ComboBox
   Friend WithEvents cbNbrPeriods As System.Windows.Forms.ComboBox
   Friend WithEvents chkShowAllItems As System.Windows.Forms.CheckBox
   Friend WithEvents dgRerent As System.Windows.Forms.DataGrid
   Friend WithEvents dtpStartRes As System.Windows.Forms.DateTimePicker
   Friend WithEvents Label1 As System.Windows.Forms.Label
   Friend WithEvents Label2 As System.Windows.Forms.Label
   Friend WithEvents Label3 As System.Windows.Forms.Label
   Friend WithEvents Label4 As System.Windows.Forms.Label
   Friend WithEvents Label6 As System.Windows.Forms.Label
   Friend WithEvents Label8 As System.Windows.Forms.Label
   Friend WithEvents Label9 As System.Windows.Forms.Label
   Friend WithEvents lblRecordID As System.Windows.Forms.Label
   Friend WithEvents textCost As System.Windows.Forms.TextBox
   Friend WithEvents textPhone As System.Windows.Forms.TextBox
   Friend WithEvents textPO As System.Windows.Forms.TextBox
   Friend WithEvents textReRentEquip As System.Windows.Forms.TextBox
   Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
   Friend WithEvents HelpProvider1 As System.Windows.Forms.HelpProvider
   Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
   Friend WithEvents mnuFile As System.Windows.Forms.MenuItem
   Friend WithEvents mnuRestoreItem As System.Windows.Forms.MenuItem
   Friend WithEvents mnuExit As System.Windows.Forms.MenuItem
   Public WithEvents txtDaily As System.Windows.Forms.TextBox
   Public WithEvents txtHalfDay As System.Windows.Forms.TextBox
   Public WithEvents optDaily As System.Windows.Forms.RadioButton
   Public WithEvents optHalfDay As System.Windows.Forms.RadioButton
   Public WithEvents txtWeekEnd As System.Windows.Forms.TextBox
   Public WithEvents optWeekEnd As System.Windows.Forms.RadioButton
   Public WithEvents txtMonthly As System.Windows.Forms.TextBox
   Public WithEvents txtWeekly As System.Windows.Forms.TextBox
   Public WithEvents optMonthly As System.Windows.Forms.RadioButton
   Public WithEvents optWeekly As System.Windows.Forms.RadioButton
   Public WithEvents txtHourly As System.Windows.Forms.TextBox
   Public WithEvents optHour As System.Windows.Forms.RadioButton
   Friend WithEvents mnuHelp As System.Windows.Forms.MenuItem
   Friend WithEvents mnuPreferences As System.Windows.Forms.MenuItem
   Friend WithEvents mnuCloseAfterRent As System.Windows.Forms.MenuItem
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
      Me.components = New System.ComponentModel.Container()
      Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmRerent))
      Me.btnAdd = New System.Windows.Forms.Button()
      Me.textPO = New System.Windows.Forms.TextBox()
      Me.Label9 = New System.Windows.Forms.Label()
      Me.cbNbrPeriods = New System.Windows.Forms.ComboBox()
      Me.Label8 = New System.Windows.Forms.Label()
      Me.textReRentEquip = New System.Windows.Forms.TextBox()
      Me.Label6 = New System.Windows.Forms.Label()
      Me.dgRerent = New System.Windows.Forms.DataGrid()
      Me.cbCustomers = New System.Windows.Forms.ComboBox()
      Me._lblFieldLable_0 = New System.Windows.Forms.Label()
      Me.Label1 = New System.Windows.Forms.Label()
      Me.dtpStartRes = New System.Windows.Forms.DateTimePicker()
      Me.Label2 = New System.Windows.Forms.Label()
      Me.textPhone = New System.Windows.Forms.TextBox()
      Me.btnSave = New System.Windows.Forms.Button()
      Me.btnDelete = New System.Windows.Forms.Button()
      Me.btnClose = New System.Windows.Forms.Button()
      Me.textCost = New System.Windows.Forms.TextBox()
      Me.Label3 = New System.Windows.Forms.Label()
      Me.btnRent = New System.Windows.Forms.Button()
      Me.Label4 = New System.Windows.Forms.Label()
      Me.lblRecordID = New System.Windows.Forms.Label()
      Me.chkShowAllItems = New System.Windows.Forms.CheckBox()
      Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
      Me.txtDaily = New System.Windows.Forms.TextBox()
      Me.txtHalfDay = New System.Windows.Forms.TextBox()
      Me.txtWeekEnd = New System.Windows.Forms.TextBox()
      Me.txtMonthly = New System.Windows.Forms.TextBox()
      Me.txtWeekly = New System.Windows.Forms.TextBox()
      Me.txtHourly = New System.Windows.Forms.TextBox()
      Me.HelpProvider1 = New System.Windows.Forms.HelpProvider()
      Me.MainMenu1 = New System.Windows.Forms.MainMenu()
      Me.mnuFile = New System.Windows.Forms.MenuItem()
      Me.mnuRestoreItem = New System.Windows.Forms.MenuItem()
      Me.mnuExit = New System.Windows.Forms.MenuItem()
      Me.mnuHelp = New System.Windows.Forms.MenuItem()
      Me.mnuPreferences = New System.Windows.Forms.MenuItem()
      Me.mnuCloseAfterRent = New System.Windows.Forms.MenuItem()
      Me.optDaily = New System.Windows.Forms.RadioButton()
      Me.optHalfDay = New System.Windows.Forms.RadioButton()
      Me.optWeekEnd = New System.Windows.Forms.RadioButton()
      Me.optMonthly = New System.Windows.Forms.RadioButton()
      Me.optWeekly = New System.Windows.Forms.RadioButton()
      Me.optHour = New System.Windows.Forms.RadioButton()
      CType(Me.dgRerent, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.SuspendLayout()
      '
      'btnAdd
      '
      Me.btnAdd.Location = New System.Drawing.Point(559, 292)
      Me.btnAdd.Name = "btnAdd"
      Me.btnAdd.Size = New System.Drawing.Size(96, 22)
      Me.btnAdd.TabIndex = 13
      Me.btnAdd.Text = "&Add New Item"
      '
      'textPO
      '
      Me.textPO.Location = New System.Drawing.Point(111, 379)
      Me.textPO.Name = "textPO"
      Me.textPO.Size = New System.Drawing.Size(111, 20)
      Me.textPO.TabIndex = 3
      Me.textPO.Tag = "(No Auto Formatting)"
      Me.textPO.Text = ""
      '
      'Label9
      '
      Me.Label9.AutoSize = True
      Me.Label9.Location = New System.Drawing.Point(15, 382)
      Me.Label9.Name = "Label9"
      Me.Label9.Size = New System.Drawing.Size(84, 13)
      Me.Label9.TabIndex = 24
      Me.Label9.Text = "Purchase Order"
      '
      'cbNbrPeriods
      '
      Me.cbNbrPeriods.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16"})
      Me.cbNbrPeriods.Location = New System.Drawing.Point(455, 440)
      Me.cbNbrPeriods.Name = "cbNbrPeriods"
      Me.cbNbrPeriods.Size = New System.Drawing.Size(79, 21)
      Me.cbNbrPeriods.TabIndex = 11
      Me.cbNbrPeriods.Text = "1"
      '
      'Label8
      '
      Me.Label8.AutoSize = True
      Me.Label8.Location = New System.Drawing.Point(376, 445)
      Me.Label8.Name = "Label8"
      Me.Label8.Size = New System.Drawing.Size(64, 13)
      Me.Label8.TabIndex = 22
      Me.Label8.Text = "Nbr Periods"
      '
      'textReRentEquip
      '
      Me.HelpProvider1.SetHelpString(Me.textReRentEquip, "Enter ReRent company and equip")
      Me.textReRentEquip.Location = New System.Drawing.Point(112, 302)
      Me.textReRentEquip.MaxLength = 29
      Me.textReRentEquip.Name = "textReRentEquip"
      Me.HelpProvider1.SetShowHelp(Me.textReRentEquip, True)
      Me.textReRentEquip.Size = New System.Drawing.Size(225, 20)
      Me.textReRentEquip.TabIndex = 0
      Me.textReRentEquip.Tag = "(No Auto Formatting)"
      Me.textReRentEquip.Text = ""
      '
      'Label6
      '
      Me.Label6.Location = New System.Drawing.Point(36, 300)
      Me.Label6.Name = "Label6"
      Me.Label6.Size = New System.Drawing.Size(68, 27)
      Me.Label6.TabIndex = 16
      Me.Label6.Text = "Equip Name && Company"
      '
      'dgRerent
      '
      Me.dgRerent.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                  Or System.Windows.Forms.AnchorStyles.Left) _
                  Or System.Windows.Forms.AnchorStyles.Right)
      Me.dgRerent.CaptionVisible = False
      Me.dgRerent.DataMember = ""
      Me.dgRerent.HeaderForeColor = System.Drawing.SystemColors.ControlText
      Me.dgRerent.Location = New System.Drawing.Point(4, 6)
      Me.dgRerent.Name = "dgRerent"
      Me.dgRerent.Size = New System.Drawing.Size(664, 283)
      Me.dgRerent.TabIndex = 26
      '
      'cbCustomers
      '
      Me.cbCustomers.Location = New System.Drawing.Point(111, 354)
      Me.cbCustomers.Name = "cbCustomers"
      Me.cbCustomers.Size = New System.Drawing.Size(252, 21)
      Me.cbCustomers.TabIndex = 2
      '
      '_lblFieldLable_0
      '
      Me._lblFieldLable_0.BackColor = System.Drawing.SystemColors.Control
      Me._lblFieldLable_0.Cursor = System.Windows.Forms.Cursors.Default
      Me._lblFieldLable_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me._lblFieldLable_0.ForeColor = System.Drawing.SystemColors.ControlText
      Me._lblFieldLable_0.Location = New System.Drawing.Point(39, 350)
      Me._lblFieldLable_0.Name = "_lblFieldLable_0"
      Me._lblFieldLable_0.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me._lblFieldLable_0.Size = New System.Drawing.Size(56, 27)
      Me._lblFieldLable_0.TabIndex = 28
      Me._lblFieldLable_0.Text = "Select Customer"
      '
      'Label1
      '
      Me.Label1.AutoSize = True
      Me.Label1.Location = New System.Drawing.Point(46, 334)
      Me.Label1.Name = "Label1"
      Me.Label1.Size = New System.Drawing.Size(55, 13)
      Me.Label1.TabIndex = 29
      Me.Label1.Text = "Rent Date"
      '
      'dtpStartRes
      '
      Me.dtpStartRes.CustomFormat = "MM/dd/yyyy hh:mm tt"
      Me.dtpStartRes.Format = System.Windows.Forms.DateTimePickerFormat.Custom
      Me.dtpStartRes.Location = New System.Drawing.Point(112, 330)
      Me.dtpStartRes.Name = "dtpStartRes"
      Me.dtpStartRes.Size = New System.Drawing.Size(197, 20)
      Me.dtpStartRes.TabIndex = 1
      '
      'Label2
      '
      Me.Label2.AutoSize = True
      Me.Label2.Location = New System.Drawing.Point(57, 404)
      Me.Label2.Name = "Label2"
      Me.Label2.Size = New System.Drawing.Size(37, 13)
      Me.Label2.TabIndex = 31
      Me.Label2.Text = "Phone"
      '
      'textPhone
      '
      Me.textPhone.Location = New System.Drawing.Point(111, 403)
      Me.textPhone.Name = "textPhone"
      Me.textPhone.Size = New System.Drawing.Size(111, 20)
      Me.textPhone.TabIndex = 4
      Me.textPhone.Tag = "(No Auto Formatting)"
      Me.textPhone.Text = ""
      '
      'btnSave
      '
      Me.btnSave.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
      Me.HelpProvider1.SetHelpString(Me.btnSave, "Save changes to existing or new item")
      Me.btnSave.Location = New System.Drawing.Point(559, 319)
      Me.btnSave.Name = "btnSave"
      Me.HelpProvider1.SetShowHelp(Me.btnSave, True)
      Me.btnSave.Size = New System.Drawing.Size(96, 40)
      Me.btnSave.TabIndex = 12
      Me.btnSave.Text = "&Save New or Changed Item"
      Me.ToolTip1.SetToolTip(Me.btnSave, "Save changes to new or changed record")
      '
      'btnDelete
      '
      Me.btnDelete.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
      Me.HelpProvider1.SetHelpString(Me.btnDelete, "Delete the Selected item")
      Me.btnDelete.Location = New System.Drawing.Point(559, 361)
      Me.btnDelete.Name = "btnDelete"
      Me.HelpProvider1.SetShowHelp(Me.btnDelete, True)
      Me.btnDelete.Size = New System.Drawing.Size(96, 22)
      Me.btnDelete.TabIndex = 14
      Me.btnDelete.Text = "&Delete"
      Me.ToolTip1.SetToolTip(Me.btnDelete, "Delete selected row")
      '
      'btnClose
      '
      Me.btnClose.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
      Me.btnClose.Location = New System.Drawing.Point(559, 389)
      Me.btnClose.Name = "btnClose"
      Me.btnClose.Size = New System.Drawing.Size(96, 22)
      Me.btnClose.TabIndex = 15
      Me.btnClose.Text = "&Close"
      '
      'textCost
      '
      Me.textCost.Location = New System.Drawing.Point(408, 469)
      Me.textCost.Name = "textCost"
      Me.textCost.Size = New System.Drawing.Size(78, 20)
      Me.textCost.TabIndex = 8
      Me.textCost.Tag = "$#,##0.00;($#,##0.00)"
      Me.textCost.Text = ""
      Me.textCost.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      Me.textCost.Visible = False
      '
      'Label3
      '
      Me.Label3.AutoSize = True
      Me.Label3.Location = New System.Drawing.Point(368, 477)
      Me.Label3.Name = "Label3"
      Me.Label3.Size = New System.Drawing.Size(27, 13)
      Me.Label3.TabIndex = 36
      Me.Label3.Text = "Cost"
      Me.Label3.Visible = False
      '
      'btnRent
      '
      Me.btnRent.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
      Me.HelpProvider1.SetHelpString(Me.btnRent, "Add the selected item to shopping cart and mark as rented")
      Me.btnRent.Location = New System.Drawing.Point(560, 422)
      Me.btnRent.Name = "btnRent"
      Me.HelpProvider1.SetShowHelp(Me.btnRent, True)
      Me.btnRent.Size = New System.Drawing.Size(96, 44)
      Me.btnRent.TabIndex = 16
      Me.btnRent.Text = "&Rent Checked Item"
      Me.ToolTip1.SetToolTip(Me.btnRent, "Rent the selected item")
      '
      'Label4
      '
      Me.Label4.AutoSize = True
      Me.Label4.Location = New System.Drawing.Point(41, 430)
      Me.Label4.Name = "Label4"
      Me.Label4.Size = New System.Drawing.Size(55, 13)
      Me.Label4.TabIndex = 39
      Me.Label4.Text = "Record ID"
      '
      'lblRecordID
      '
      Me.lblRecordID.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
      Me.lblRecordID.Location = New System.Drawing.Point(112, 427)
      Me.lblRecordID.Name = "lblRecordID"
      Me.lblRecordID.Size = New System.Drawing.Size(88, 19)
      Me.lblRecordID.TabIndex = 40
      Me.lblRecordID.TextAlign = System.Drawing.ContentAlignment.MiddleRight
      '
      'chkShowAllItems
      '
      Me.HelpProvider1.SetHelpString(Me.chkShowAllItems, "Check to show rented items in order to restore a cancelled rerent item")
      Me.chkShowAllItems.Location = New System.Drawing.Point(240, 440)
      Me.chkShowAllItems.Name = "chkShowAllItems"
      Me.HelpProvider1.SetShowHelp(Me.chkShowAllItems, True)
      Me.chkShowAllItems.Size = New System.Drawing.Size(99, 28)
      Me.chkShowAllItems.TabIndex = 41
      Me.chkShowAllItems.Text = "Show All Rented Items"
      Me.ToolTip1.SetToolTip(Me.chkShowAllItems, "When checked, shows items already rented")
      '
      'txtDaily
      '
      Me.txtDaily.AcceptsReturn = True
      Me.txtDaily.AutoSize = False
      Me.txtDaily.BackColor = System.Drawing.SystemColors.Window
      Me.txtDaily.Cursor = System.Windows.Forms.Cursors.IBeam
      Me.txtDaily.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.txtDaily.ForeColor = System.Drawing.SystemColors.WindowText
      Me.txtDaily.Location = New System.Drawing.Point(456, 344)
      Me.txtDaily.MaxLength = 0
      Me.txtDaily.Name = "txtDaily"
      Me.txtDaily.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.txtDaily.Size = New System.Drawing.Size(75, 19)
      Me.txtDaily.TabIndex = 7
      Me.txtDaily.Tag = "$#,##0.00;($#,##0.00)"
      Me.txtDaily.Text = "$0.00"
      Me.txtDaily.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      Me.ToolTip1.SetToolTip(Me.txtDaily, "Change this price manually to give a price break")
      '
      'txtHalfDay
      '
      Me.txtHalfDay.AcceptsReturn = True
      Me.txtHalfDay.AutoSize = False
      Me.txtHalfDay.BackColor = System.Drawing.SystemColors.Window
      Me.txtHalfDay.Cursor = System.Windows.Forms.Cursors.IBeam
      Me.txtHalfDay.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.txtHalfDay.ForeColor = System.Drawing.SystemColors.WindowText
      Me.txtHalfDay.Location = New System.Drawing.Point(456, 320)
      Me.txtHalfDay.MaxLength = 0
      Me.txtHalfDay.Name = "txtHalfDay"
      Me.txtHalfDay.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.txtHalfDay.Size = New System.Drawing.Size(75, 19)
      Me.txtHalfDay.TabIndex = 6
      Me.txtHalfDay.Tag = "$#,##0.00;($#,##0.00)"
      Me.txtHalfDay.Text = "$0.00"
      Me.txtHalfDay.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      Me.ToolTip1.SetToolTip(Me.txtHalfDay, "Change this price manually to give a price break")
      '
      'txtWeekEnd
      '
      Me.txtWeekEnd.AcceptsReturn = True
      Me.txtWeekEnd.AutoSize = False
      Me.txtWeekEnd.BackColor = System.Drawing.SystemColors.Window
      Me.txtWeekEnd.Cursor = System.Windows.Forms.Cursors.IBeam
      Me.txtWeekEnd.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.txtWeekEnd.ForeColor = System.Drawing.SystemColors.WindowText
      Me.txtWeekEnd.Location = New System.Drawing.Point(456, 416)
      Me.txtWeekEnd.MaxLength = 0
      Me.txtWeekEnd.Name = "txtWeekEnd"
      Me.txtWeekEnd.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.txtWeekEnd.Size = New System.Drawing.Size(75, 19)
      Me.txtWeekEnd.TabIndex = 10
      Me.txtWeekEnd.Tag = "$#,##0.00;($#,##0.00)"
      Me.txtWeekEnd.Text = "$0.00"
      Me.txtWeekEnd.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      Me.ToolTip1.SetToolTip(Me.txtWeekEnd, "Change this price manually to give a price break")
      Me.txtWeekEnd.Visible = False
      '
      'txtMonthly
      '
      Me.txtMonthly.AcceptsReturn = True
      Me.txtMonthly.AutoSize = False
      Me.txtMonthly.BackColor = System.Drawing.SystemColors.Window
      Me.txtMonthly.Cursor = System.Windows.Forms.Cursors.IBeam
      Me.txtMonthly.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.txtMonthly.ForeColor = System.Drawing.SystemColors.WindowText
      Me.txtMonthly.Location = New System.Drawing.Point(456, 392)
      Me.txtMonthly.MaxLength = 0
      Me.txtMonthly.Name = "txtMonthly"
      Me.txtMonthly.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.txtMonthly.Size = New System.Drawing.Size(75, 19)
      Me.txtMonthly.TabIndex = 9
      Me.txtMonthly.Tag = "$#,##0.00;($#,##0.00)"
      Me.txtMonthly.Text = "$0.00"
      Me.txtMonthly.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      Me.ToolTip1.SetToolTip(Me.txtMonthly, "Change this price manually to give a price break")
      '
      'txtWeekly
      '
      Me.txtWeekly.AcceptsReturn = True
      Me.txtWeekly.AutoSize = False
      Me.txtWeekly.BackColor = System.Drawing.SystemColors.Window
      Me.txtWeekly.Cursor = System.Windows.Forms.Cursors.IBeam
      Me.txtWeekly.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.txtWeekly.ForeColor = System.Drawing.SystemColors.WindowText
      Me.txtWeekly.Location = New System.Drawing.Point(456, 368)
      Me.txtWeekly.MaxLength = 0
      Me.txtWeekly.Name = "txtWeekly"
      Me.txtWeekly.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.txtWeekly.Size = New System.Drawing.Size(75, 19)
      Me.txtWeekly.TabIndex = 8
      Me.txtWeekly.Tag = "$#,##0.00;($#,##0.00)"
      Me.txtWeekly.Text = "$0.00"
      Me.txtWeekly.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      Me.ToolTip1.SetToolTip(Me.txtWeekly, "Change this price manually to give a price break")
      '
      'txtHourly
      '
      Me.txtHourly.AcceptsReturn = True
      Me.txtHourly.AutoSize = False
      Me.txtHourly.BackColor = System.Drawing.SystemColors.Window
      Me.txtHourly.Cursor = System.Windows.Forms.Cursors.IBeam
      Me.txtHourly.Enabled = False
      Me.txtHourly.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.txtHourly.ForeColor = System.Drawing.SystemColors.WindowText
      Me.txtHourly.Location = New System.Drawing.Point(456, 296)
      Me.txtHourly.MaxLength = 0
      Me.txtHourly.Name = "txtHourly"
      Me.txtHourly.ReadOnly = True
      Me.txtHourly.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.txtHourly.Size = New System.Drawing.Size(75, 19)
      Me.txtHourly.TabIndex = 5
      Me.txtHourly.Tag = "$#,##0.00;($#,##0.00)"
      Me.txtHourly.Text = "$0.00"
      Me.txtHourly.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      Me.ToolTip1.SetToolTip(Me.txtHourly, "Change this price manually to give a price break")
      Me.txtHourly.Visible = False
      '
      'MainMenu1
      '
      Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFile, Me.mnuHelp, Me.mnuPreferences})
      '
      'mnuFile
      '
      Me.mnuFile.Index = 0
      Me.mnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuRestoreItem, Me.mnuExit})
      Me.mnuFile.Text = "&File"
      '
      'mnuRestoreItem
      '
      Me.mnuRestoreItem.Index = 0
      Me.mnuRestoreItem.Text = "&Restore Selected Item"
      '
      'mnuExit
      '
      Me.mnuExit.Index = 1
      Me.mnuExit.Shortcut = System.Windows.Forms.Shortcut.CtrlE
      Me.mnuExit.Text = "&Exit"
      '
      'mnuHelp
      '
      Me.mnuHelp.Index = 1
      Me.mnuHelp.Text = "&Help!"
      '
      'mnuPreferences
      '
      Me.mnuPreferences.Index = 2
      Me.mnuPreferences.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuCloseAfterRent})
      Me.mnuPreferences.Text = "&Preferences"
      '
      'mnuCloseAfterRent
      '
      Me.mnuCloseAfterRent.Checked = True
      Me.mnuCloseAfterRent.Index = 0
      Me.mnuCloseAfterRent.Text = "Close After Rent"
      '
      'optDaily
      '
      Me.optDaily.BackColor = System.Drawing.SystemColors.Control
      Me.optDaily.Checked = True
      Me.optDaily.Cursor = System.Windows.Forms.Cursors.Default
      Me.optDaily.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.optDaily.ForeColor = System.Drawing.SystemColors.ControlText
      Me.optDaily.Location = New System.Drawing.Point(371, 344)
      Me.optDaily.Name = "optDaily"
      Me.optDaily.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.optDaily.Size = New System.Drawing.Size(77, 21)
      Me.optDaily.TabIndex = 45
      Me.optDaily.TabStop = True
      Me.optDaily.Text = "&Daily"
      '
      'optHalfDay
      '
      Me.optHalfDay.BackColor = System.Drawing.SystemColors.Control
      Me.optHalfDay.Cursor = System.Windows.Forms.Cursors.Default
      Me.optHalfDay.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.optHalfDay.ForeColor = System.Drawing.SystemColors.ControlText
      Me.optHalfDay.Location = New System.Drawing.Point(371, 318)
      Me.optHalfDay.Name = "optHalfDay"
      Me.optHalfDay.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.optHalfDay.Size = New System.Drawing.Size(70, 25)
      Me.optHalfDay.TabIndex = 44
      Me.optHalfDay.Text = "&Half Day/ 4 Hrs"
      '
      'optWeekEnd
      '
      Me.optWeekEnd.BackColor = System.Drawing.SystemColors.Control
      Me.optWeekEnd.Cursor = System.Windows.Forms.Cursors.Default
      Me.optWeekEnd.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.optWeekEnd.ForeColor = System.Drawing.SystemColors.ControlText
      Me.optWeekEnd.Location = New System.Drawing.Point(371, 416)
      Me.optWeekEnd.Name = "optWeekEnd"
      Me.optWeekEnd.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.optWeekEnd.Size = New System.Drawing.Size(80, 21)
      Me.optWeekEnd.TabIndex = 50
      Me.optWeekEnd.Text = "Week &End"
      Me.optWeekEnd.Visible = False
      '
      'optMonthly
      '
      Me.optMonthly.BackColor = System.Drawing.SystemColors.Control
      Me.optMonthly.Cursor = System.Windows.Forms.Cursors.Default
      Me.optMonthly.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.optMonthly.ForeColor = System.Drawing.SystemColors.ControlText
      Me.optMonthly.Location = New System.Drawing.Point(371, 392)
      Me.optMonthly.Name = "optMonthly"
      Me.optMonthly.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.optMonthly.Size = New System.Drawing.Size(71, 17)
      Me.optMonthly.TabIndex = 49
      Me.optMonthly.Text = "&Monthly"
      '
      'optWeekly
      '
      Me.optWeekly.BackColor = System.Drawing.SystemColors.Control
      Me.optWeekly.Cursor = System.Windows.Forms.Cursors.Default
      Me.optWeekly.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.optWeekly.ForeColor = System.Drawing.SystemColors.ControlText
      Me.optWeekly.Location = New System.Drawing.Point(371, 368)
      Me.optWeekly.Name = "optWeekly"
      Me.optWeekly.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.optWeekly.Size = New System.Drawing.Size(65, 17)
      Me.optWeekly.TabIndex = 48
      Me.optWeekly.Text = "&Weekly"
      '
      'optHour
      '
      Me.optHour.BackColor = System.Drawing.SystemColors.Control
      Me.optHour.Cursor = System.Windows.Forms.Cursors.Default
      Me.optHour.Enabled = False
      Me.optHour.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.optHour.ForeColor = System.Drawing.SystemColors.ControlText
      Me.optHour.Location = New System.Drawing.Point(372, 296)
      Me.optHour.Name = "optHour"
      Me.optHour.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.optHour.Size = New System.Drawing.Size(53, 15)
      Me.optHour.TabIndex = 54
      Me.optHour.Text = "Hou&rly"
      Me.optHour.Visible = False
      '
      'frmRerent
      '
      Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
      Me.ClientSize = New System.Drawing.Size(672, 473)
      Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtHourly, Me.optHour, Me.txtWeekEnd, Me.optWeekEnd, Me.txtMonthly, Me.txtWeekly, Me.optMonthly, Me.optWeekly, Me.txtDaily, Me.txtHalfDay, Me.optDaily, Me.optHalfDay, Me.chkShowAllItems, Me.lblRecordID, Me.Label4, Me.Label3, Me.Label2, Me.Label1, Me.Label9, Me.Label8, Me.btnRent, Me.textCost, Me.btnClose, Me.btnDelete, Me.btnSave, Me.textPhone, Me.dtpStartRes, Me.cbCustomers, Me._lblFieldLable_0, Me.dgRerent, Me.textPO, Me.cbNbrPeriods, Me.textReRentEquip, Me.Label6, Me.btnAdd})
      Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
      Me.Menu = Me.MainMenu1
      Me.MinimizeBox = False
      Me.Name = "frmRerent"
      Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "Rerent Equipment"
      CType(Me.dgRerent, System.ComponentModel.ISupportInitialize).EndInit()
      Me.ResumeLayout(False)

   End Sub

#End Region

#Region "Private Variables"
   Private msAddorEdit As String
   Private mbDirty As Boolean
   Private oDA As New CDataAccess()
   Private oCG As New CGrid()
   Private miHitRow As Integer
   Private dtRerent As DataTable
   Dim bChecked As Boolean
#End Region

#Region "Button Events"
   Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
      Dim SQL As String
      Dim dt As New DataTable()


      Try
         If mbDirty Then
            If MsgBox("You have unsaved changes; do you want to add without saving your changes?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
               Exit Sub
            End If
         End If
         SQL = "select max(unique_id) from rerents"
         If oDA.SendQuery(SQL, dt, ConnectString, "dt") = 0 Then
            Throw New System.Exception("Can't retreive unique_id from rerents table.")
         End If

         ClearTextBoxes()
         Me.lblRecordID.Text = MNI(dt.Rows(0).Item(0)) + 1
         msAddorEdit = "A"
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub


   ''' <summary>
   ''' Clear the textboxes for an add.
   ''' </summary>
   Private Sub ClearTextBoxes()
      With Me
         .textCost.Text = FormatCurrency(0)
         .textPhone.Text = String.Empty
         .textPO.Text = String.Empty
         .txtHourly.Text = FormatCurrency(0)
         .txtHalfDay.Text = FormatCurrency(0)
         .txtDaily.Text = FormatCurrency(0)
         .txtMonthly.Text = FormatCurrency(0)
         .txtWeekly.Text = FormatCurrency(0)
         .txtWeekEnd.Text = FormatCurrency(0)
         .textPhone.Text = String.Empty
         .cbNbrPeriods.Text = 1
         .optDaily.Checked = True
         .cbCustomers.Text = String.Empty
         .textReRentEquip.Text = String.Empty
      End With
   End Sub
   ''' <summary>
   ''' Load text boxes from selected grid row.
   ''' </summary>
   Private Sub LoadTextBoxes()

      Try
         Dim dr As DataRow = dtRerent.Rows(miHitRow)
         With Me
            'Me.textCost.Text = FormatCurrency(MNSng(dr("our_cost")))
            Me.textPhone.Text = MNS(dr("customer_phone"))
            Me.textPO.Text = MNS(dr("po_number"))
            Me.txtHourly.Text = FormatCurrency(dr("hourrate"))
            Me.txtHalfDay.Text = FormatCurrency(dr("halfday"))
            Me.txtDaily.Text = FormatCurrency(dr("daily"))
            Me.txtWeekly.Text = FormatCurrency(dr("weekly"))
            Me.txtMonthly.Text = FormatCurrency(dr("monthly"))
            Me.txtWeekEnd.Text = FormatCurrency(dr("weekend"))

            Select Case MNS(dr("period"))
               Case HOURLY
                  .txtHourly.Text = FormatCurrency(MNSng(dr("cust_price")))
                  .optHour.Checked = True
               Case HALF_DAY
                  .txtHalfDay.Text = FormatCurrency(MNSng(dr("cust_price")))
                  .optHalfDay.Checked = True
               Case DAILY
                  .txtDaily.Text = FormatCurrency(MNSng(dr("cust_price")))
                  .optDaily.Checked = True
               Case WEEKLY
                  .txtWeekly.Text = FormatCurrency(MNSng(dr("cust_price")))
                  .optWeekly.Checked = True
               Case MONTHLY
                  .txtMonthly.Text = FormatCurrency(MNSng(dr("cust_price")))
                  .optMonthly.Checked = True
               Case WEEK_END
                  .txtWeekEnd.Text = FormatCurrency(MNSng(dr("cust_price")))
                  .optWeekEnd.Checked = True
            End Select

            Me.cbNbrPeriods.Text = MNI(dr("nbr_periods"))
            Me.cbCustomers.Text = MNS(dr("customer_name"))
            Me.textReRentEquip.Text = MNS(dr("equip_name"))
            If Not IsDBNull(dr("date_needed")) Then
               Me.dtpStartRes.Value = dr("date_needed")
            Else
               Me.dtpStartRes.Value = Now
            End If
            Me.lblRecordID.Text = dr("unique_id")
         End With
      Catch ex As System.Exception
         'StructuredErrorHandler(ex)
      End Try
   End Sub

   Private Sub dgRerent_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgRerent.MouseUp
      Try
         Dim pt = New Point(e.X, e.Y)
         Dim hti As DataGrid.HitTestInfo = Me.dgRerent.HitTest(pt)
         Me.dgRerent.Select(hti.Row)
         oCG.SelectCkBoxRow(dtRerent, dgRerent, e, "Rent", bChecked)
         miHitRow = hti.Row
         Me.LoadTextBoxes()
         mbDirty = False
         msAddorEdit = "E"
      Catch ex As System.Exception
         'StructuredErrorHandler(ex)
      End Try
   End Sub


   Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
      Dim SQL As String
      Dim dt As New DataTable()
      Dim oDA As New CDataAccess()
      Dim iRows As Integer

      Try
         Dim sErr As String = ""

         If MsgBox("Are you sure you want to delete the selected row?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
            Exit Sub
         End If
         SQL = "delete from rerents "
         SQL &= "where unique_id= " & Me.lblRecordID.Text & " "
         iRows = oDA.SendActionSql(SQL, ConnectString, sErr)
         If iRows = 0 Then
            MsgBox("Delete of equipment item failed.  " & Chr(10) & sErr, MsgBoxStyle.Critical)
            Exit Sub
         End If
         Me.LoadTheGrid()
         miHitRow = 0
         LoadTextBoxes()
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub
   Private Sub LoadCustomerCombo()
      Dim dt As New DataTable()
      Dim i As Integer
      Dim sql As String = "select companyname from customers "
      sql &= "order by companyname"
      oDA.SendQuery(sql, dt, ConnectString)
      Me.cbCustomers.Items.Clear()

      For i = 0 To dt.Rows.Count - 1
         With dt.Rows(i)
            Me.cbCustomers.Items.Add(.Item("companyname"))
         End With
      Next
   End Sub
   ''' <summary>
   ''' Load the grid with unrented items.  But if ShowAll = true
   ''' Show the rented items also
   ''' </summary>
   ''' <param name = "Optional ShowAll"></param>
   Private Sub LoadTheGrid(Optional ByVal ShowAll As Boolean = False)
      Dim SQL As String

      Try
         dtRerent = New DataTable()
         oCG.InitializeDatatableForStyles(dtRerent)
         oCG.BindDataTableToGrid(dtRerent, Me.dgRerent)
         SQL = ""
         SQL &= "select equip_name, customer_name, date_needed, "
         SQL &= "customer_phone, nbr_periods, period, po_number, "
         SQL &= "customer_price as Cust_Price, rented_date, unique_id "
         SQL &= ",hourrate,halfday,daily,weekly,monthly,weekend,newprices "
         SQL &= "from rerents "
         If Not ShowAll Then
            SQL &= "where isnull(rented_date) "
         Else
            SQL &= "where isnull(returned_date) "
         End If
         SQL &= "order by date_needed "

         oDA.SendQuery(SQL, dtRerent, ConnectString, "dt")
         If dtRerent.Rows.Count > 0 Then
            Dim Formats() As String = _
               {"", "150", "T", "L", _
               "", "150", "T", "L", _
               "M/d/yyyy hh:mm tt", "100", "T", "L", _
               "", "100", "T", "L", _
               "", "60", "T", "R", _
               "", "60", "T", "L", _
               "", "60", "T", "L", _
               "$#,##0.00", "60", "T", "R", _
               "M/d/yyyy hh:mm tt", "60", "T", "R", _
               "", "60", "T", "R"}
            oCG.SetTablesStyle(dtRerent, Me.dgRerent, Formats)
            oCG.BindDataTableToGrid(dtRerent, Me.dgRerent)
            oCG.DisableAddNew(Me.dgRerent, Me)
            miHitRow = 0
         End If
         mbDirty = False
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

   Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
      If mbDirty Then
         If MsgBox("You have unsaved changes; do you want to close without saving your changes?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
            Exit Sub
         End If
      End If

      Me.Close()
      System.Windows.Forms.Application.DoEvents()
   End Sub

#End Region

#Region "Control Events"
   Private Sub txtHalfDay_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtHalfDay.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
      e.Handled = CkKeyPressNumeric(Asc(Val(e.KeyChar)), txtHalfDay) = 0
   End Sub
   Private Sub txtHalfDay_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtHalfDay.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub txtHalfDay_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHalfDay.Enter
      txtHalfDay.Text = UnFmt_T_B(txtHalfDay)
      txtHalfDay.SelectionStart = 0
      txtHalfDay.SelectionLength = txtHalfDay.Text.Trim.Length
   End Sub
   Private Sub txtHalfDay_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHalfDay.Leave
      txtHalfDay.Text = Fmt_T_B(txtHalfDay)
   End Sub
   Private Sub txtHourly_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtHourly.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
      e.Handled = CkKeyPressNumeric(Asc(Val(e.KeyChar)), txtHourly) = 0
   End Sub
   Private Sub txtHourly_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtHourly.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub txtHourly_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHourly.Enter
      txtHourly.Text = UnFmt_T_B(txtHourly)
      txtHourly.SelectionStart = 0
      txtHourly.SelectionLength = txtHourly.Text.Trim.Length
   End Sub
   Private Sub txtHourly_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHourly.Leave
      txtHourly.Text = Fmt_T_B(txtHourly)
   End Sub
   Private Sub txtDaily_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDaily.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
      e.Handled = CkKeyPressNumeric(Asc(Val(e.KeyChar)), txtDaily) = 0
   End Sub
   Private Sub txtDaily_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDaily.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub txtDaily_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDaily.Enter
      txtDaily.Text = UnFmt_T_B(txtDaily)
      txtDaily.SelectionStart = 0
      txtDaily.SelectionLength = txtDaily.Text.Trim.Length
   End Sub
   Private Sub txtDaily_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDaily.Leave
      txtDaily.Text = Fmt_T_B(txtDaily)
   End Sub
   Private Sub txtMonthly_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMonthly.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
      e.Handled = CkKeyPressNumeric(Asc(Val(e.KeyChar)), txtMonthly) = 0
   End Sub
   Private Sub txtMonthly_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtMonthly.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub txtMonthly_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMonthly.Enter
      txtMonthly.Text = UnFmt_T_B(txtMonthly)
      txtMonthly.SelectionStart = 0
      txtMonthly.SelectionLength = txtMonthly.Text.Trim.Length
   End Sub
   Private Sub txtMonthly_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMonthly.Leave
      txtMonthly.Text = Fmt_T_B(txtMonthly)
   End Sub
   Private Sub txtWeekly_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtWeekly.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
      e.Handled = CkKeyPressNumeric(Asc(Val(e.KeyChar)), txtWeekly) = 0
   End Sub
   Private Sub txtWeekly_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtWeekly.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub txtWeekly_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtWeekly.Enter
      txtWeekly.Text = UnFmt_T_B(txtWeekly)
      txtWeekly.SelectionStart = 0
      txtWeekly.SelectionLength = txtWeekly.Text.Trim.Length
   End Sub
   Private Sub txtWeekly_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtWeekly.Leave
      txtWeekly.Text = Fmt_T_B(txtWeekly)
   End Sub
   Private Sub txtWeekEnd_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtWeekEnd.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
      e.Handled = CkKeyPressNumeric(Asc(Val(e.KeyChar)), txtWeekEnd) = 0
   End Sub
   Private Sub txtWeekEnd_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtWeekEnd.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub txtWeekEnd_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtWeekEnd.Enter
      txtWeekEnd.Text = UnFmt_T_B(txtWeekEnd)
      txtWeekEnd.SelectionStart = 0
      txtWeekEnd.SelectionLength = txtWeekEnd.Text.Trim.Length
   End Sub
   Private Sub txtWeekEnd_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtWeekEnd.Leave
      txtWeekEnd.Text = Fmt_T_B(txtWeekEnd)
   End Sub

   Private Sub mnuHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuHelp.Click
      Dim sTxt As String = ""
      sTxt &= "To change an existing item, click on the desired row in the "
      sTxt &= "grid, make the required changes in the text boxes and click "
      sTxt &= "the Save Button." & vbCrLf & vbCrLf
      sTxt &= "To add a newly secured reservation from a Rerent Company, "
      sTxt &= "follow these steps:" & Chr(13) & Chr(10) & vbCrLf
      sTxt &= "1) Click the Add button." & Chr(13) & Chr(10) & vbCrLf
      sTxt &= "2) Enter all required parameters, choosing a period option "
      sTxt &= "and its corresponding price." & Chr(13) & Chr(10) & vbCrLf
      sTxt &= "3) If you enter at least Daily, Weekly, and Monthly prices, "
      sTxt &= "AutoCalc can compute the Rerent pricing automatically upon "
      sTxt &= "return and they will be printed on the checkout invoice." & Chr(13) & Chr(10) & vbCrLf
      sTxt &= "To ReRent a previously scheduled (secured) item, follow "
      sTxt &= "these steps:" & Chr(13) & Chr(10) & vbCrLf
      sTxt &= "1) Click on the desired item in the left margin of the "
      sTxt &= "grid." & Chr(13) & Chr(10) & vbCrLf
      sTxt &= "2) Click the Rent Select Equip Button." & Chr(13) & Chr(10) & vbCrLf
      sTxt &= "3) Answer Yes to the message if you have selected the "
      sTxt &= "correct equipment, otherwise answer no." & Chr(13) & Chr(10) & vbCrLf

      Dim f As New frmHelp()
      f.CannedMessage = sTxt
      f.ShowDialog()
   End Sub

   Private Sub mnuCloseAfterRent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCloseAfterRent.Click
      Me.mnuCloseAfterRent.Checked = Not Me.mnuCloseAfterRent.Checked
      SaveSetting(RENTALPRO, SETTINGS, "CLOSERERENT", Me.mnuCloseAfterRent.Checked)
   End Sub
   Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
      Dim SQL As String
      Dim sErr As String
      Dim price As Decimal
      Dim period As String
      Dim newPrices As Boolean = False

      Try
         If UnFormat(Me.txtDaily.Text) = 0 And _
            UnFormat(Me.txtWeekly.Text) = 0 And _
            UnFormat(Me.txtMonthly.Text) = 0 And _
            UnFormat(Me.txtHalfDay.Text) = 0 And _
            UnFormat(Me.txtHourly.Text) = 0 And _
            UnFormat(Me.txtWeekEnd.Text) = 0 Then
            MsgBox("You must enter at least the price of the period option that you selected.", MsgBoxStyle.Information)
            Exit Sub
         End If

         With Me
            Select Case True
               Case .optHour.Checked
                  If UnFormat(.txtHourly.Text) = 0 Then
                     GoTo norates
                  End If
                  period = HOURLY
                  price = UnFormat(.txtHourly.Text)
               Case .optDaily.Checked
                  If UnFormat(.txtDaily.Text) = 0 Then
                     GoTo norates
                  End If
                  period = DAILY
                  price = UnFormat(.txtDaily.Text)
               Case .optWeekly.Checked
                  If UnFormat(.txtWeekly.Text) = 0 Then
                     GoTo norates
                  End If
                  period = WEEKLY
                  price = UnFormat(.txtWeekly.Text)
               Case .optMonthly.Checked
                  If UnFormat(.txtMonthly.Text) = 0 Then
                     GoTo norates
                  End If
                  period = MONTHLY
                  price = UnFormat(.txtMonthly.Text)
               Case .optHalfDay.Checked
                  If UnFormat(.txtHalfDay.Text) = 0 Then
                     GoTo norates
                  End If
                  period = HALF_DAY
                  price = UnFormat(.txtHalfDay.Text)
               Case .optWeekEnd.Checked
                  If UnFormat(.txtWeekEnd.Text) = 0 Then
                     GoTo norates
                  End If
                  period = WEEK_END
                  price = UnFormat(.txtWeekEnd.Text)
            End Select

            If Me.textPhone.Text.Length = 0 Or _
               Me.cbCustomers.Text.Length = 0 Or _
               Me.cbNbrPeriods.Text.Length = 0 Or _
               Me.textPhone.Text.Length = 0 Or _
               Me.textPO.Text.Length = 0 Or _
               Me.textReRentEquip.Text.Length = 0 _
               Then
               MsgBox("You must fill all parameters.", MsgBoxStyle.Exclamation)
               Exit Sub
            End If
            If UnFormat(Me.txtDaily.Text) > 0 And _
               UnFormat(.txtWeekly.Text) > 0 And _
               UnFormat(.txtMonthly.Text) > 0 Then
               newPrices = True
            End If
            If msAddorEdit = "A" Then
               SQL = "insert into rerents "
               SQL &= "(equip_name,nbr_periods,period,date_needed,customer_name, "
               SQL &= "po_number,customer_price,customer_phone, " 'our_cost) "
               SQL &= "hourrate,halfday,daily,weekly,monthly,weekend,newprices)"
               SQL &= "values("
               SQL &= "'" & Replace(.textReRentEquip.Text, "'", "''") & "', "
               SQL &= .cbNbrPeriods.Text & ", "
               SQL &= "'" & period & "', "
               SQL &= "#" & .dtpStartRes.Value.ToString & "#, "
               SQL &= "'" & Replace(.cbCustomers.Text, "'", "''") & "', "
               SQL &= "'" & .textPO.Text & "', "
               SQL &= price & ","
               SQL &= "'" & .textPhone.Text & "', "
               'SQL &= UnFormat(.textCost.Text)
               SQL &= UnFormat(.txtHourly.Text) & ", "
               SQL &= UnFormat(.txtHalfDay.Text) & ", "
               SQL &= UnFormat(.txtDaily.Text) & ", "
               SQL &= UnFormat(.txtWeekly.Text) & ", "
               SQL &= UnFormat(.txtMonthly.Text) & ", "
               SQL &= UnFormat(.txtWeekEnd.Text) & ", "
               SQL &= newPrices & " "
               SQL &= ")"
            Else
               SQL = "update rerents "
               SQL &= "set equip_name = '" & .textReRentEquip.Text & "', "
               SQL &= "nbr_periods = " & .cbNbrPeriods.Text & ", "
               SQL &= "period = '" & period & "', "
               SQL &= "date_needed = #" & .dtpStartRes.Value & "#, "
               SQL &= "customer_name '" & Replace(.cbCustomers.Text, "'", "''") & "', "
               SQL &= "po_number = '" & .textPO.Text & "', "
               SQL &= "customer_price = " & price & ", "
               SQL &= "customer_phone = '" & .textPhone.Text & "', "
               'SQL &= "our_cost = " & UnFormat(.textCost.Text) & ", "
               SQL &= "hourrate = " & UnFormat(.txtHourly.Text) & ", "
               SQL &= "halfday = " & UnFormat(.txtHalfDay.Text) & ", "
               SQL &= "daily= " & UnFormat(.txtDaily.Text) & ", "
               SQL &= "weekly = " & UnFormat(.txtWeekly.Text) & ", "
               SQL &= "monthly=" & UnFormat(.txtMonthly.Text) & ", "
               SQL &= "weekend=" & UnFormat(.txtWeekEnd.Text) & ", "
               SQL &= "newprices=" & newPrices & " "
               SQL &= "where unique_id = " & .lblRecordID.Text
            End If
         End With
         If oDA.SendActionSql(SQL, ConnectString, sErr) < 1 Then
            MsgBox("Update failed: " & Chr(10) & sErr, MsgBoxStyle.Critical)
         End If
         msAddorEdit = "E"
         miHitRow = 0
         mbDirty = False
         'If Me.chkRefresh.Checked Then
         LoadTheGrid()
         LoadTextBoxes()
         Exit Sub
NoRates:
         MsgBox("You must enter a cost for the selected period.", MsgBoxStyle.Information)

      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try

   End Sub
   Private Sub textCost_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles textCost.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
      e.Handled = CkKeyPressNumeric(Asc(Val(e.KeyChar)), textCost) = 0
   End Sub
   Private Sub textCost_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles textCost.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub textCost_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles textCost.Enter
      textCost.Text = UnFmt_T_B(textCost)
      textCost.SelectionStart = 0
      textCost.SelectionLength = textCost.Text.Trim.Length
   End Sub
   Private Sub textCost_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles textCost.Leave
      textCost.Text = Fmt_T_B(textCost)
   End Sub
   Private Function CkKeyPressNumeric(ByVal riKeyAscii As Integer, ByVal roTB As TextBox) As Integer
      Dim liKeyReturn As Integer
      ' allow 0-9,., Back, Del,-,Ins, and / if in tag format
      On Error Resume Next
      CkKeyPressNumeric = riKeyAscii
      If riKeyAscii = Keys.Back Or _
         riKeyAscii = Keys.Insert Or _
         riKeyAscii = Keys.Delete Or _
         riKeyAscii = 46 Or _
         (riKeyAscii >= Keys.D0 And riKeyAscii <= Keys.D9) Or _
         riKeyAscii = 45 Or _
         riKeyAscii = 46 Or _
         (InStr(roTB.Tag, "/") > 0 And riKeyAscii = Keys.Divide) _
         Then
         If roTB.SelectionLength = 0 Then
            If InStr(roTB.Text, ".") > 0 Then
               If Len(Mid(roTB.Text, InStr(roTB.Text, ".") + 1)) > 1 Then
                  SendKeys.Send("{TAB}")
                  CkKeyPressNumeric = 0
               End If
            End If
         Else
            roTB.Text = ""
         End If
         Exit Function
      End If
      CkKeyPressNumeric = 0
   End Function

   Public Function UnFmt_T_B(ByVal roTB As TextBox) As Object
      On Error Resume Next
      UnFmt_T_B = Val(Replace(Replace(Replace(Replace(Replace(roTB.Text, "$", ""), ",", ""), ")", ""), "(", ""), "%", ""))
      If InStr(roTB.Text, "%") Then
         UnFmt_T_B = UnFmt_T_B / 100
      End If
      If InStr(roTB.Text, "(") > 0 And InStr(roTB.Text, ")") > 0 Then
         UnFmt_T_B = UnFmt_T_B * -1
      End If
   End Function

   Public Function Fmt_T_B(ByVal roTB As TextBox) As String
      On Error Resume Next
      If InStr(1, roTB.Tag, ";", 1) > 0 Then
         If InStr(roTB.Text, "-") > 0 Or (InStr(roTB.Text, "(") > 0 And InStr(roTB.Text, ")") > 0) Then
            Fmt_T_B = Format$(Math.Abs(Val(roTB.Text)), Mid$(roTB.Tag, InStr(roTB.Tag, ";") + 1))
         Else
            Fmt_T_B = Format$(Math.Abs(Val(roTB.Text)), Microsoft.VisualBasic.Left(roTB.Tag, InStr(roTB.Tag, ";") - 1))
         End If
      ElseIf InStr(1, roTB.Tag, "%", 1) > 0 Then
         Fmt_T_B = Format$(roTB.Text, roTB.Tag)
      Else
         Fmt_T_B = Format$(roTB.Text, roTB.Tag)
      End If
   End Function

   Public Function Fmt_D_F(ByVal rsTxt As Object, ByVal roTB As TextBox) As String
      On Error Resume Next

      If InStr(1, roTB.Tag, ";", 1) > 0 Then
         If InStr(rsTxt, "-") Then

            Fmt_D_F = Format$(Replace(rsTxt, "-", ""), Mid$(roTB.Tag, InStr(roTB.Tag, ";") + 1))
         Else
            Fmt_D_F = Format$(rsTxt, Microsoft.VisualBasic.Left(roTB.Tag, InStr(roTB.Tag, ";") - 1))
         End If
      ElseIf InStr(1, roTB.Tag, "%", 1) > 0 Then
         Fmt_D_F = Format$(rsTxt, roTB.Tag)
      Else
         Fmt_D_F = Format$(rsTxt, roTB.Tag)
      End If
   End Function
   Private Sub textReRentEquip_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles textReRentEquip.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
   End Sub
   Private Sub textReRentEquip_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles textReRentEquip.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub textReRentEquip_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles textReRentEquip.Enter
      textReRentEquip.SelectionStart = 0
      textReRentEquip.SelectionLength = textReRentEquip.Text.Trim.Length
   End Sub
   Private Sub textPO_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles textPO.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
   End Sub
   Private Sub textPO_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles textPO.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub textPO_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles textPO.Enter
      textPO.SelectionStart = 0
      textPO.SelectionLength = textPO.Text.Trim.Length
   End Sub
   Private Sub textPhone_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles textPhone.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
   End Sub
   Private Sub textPhone_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles textPhone.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub textPhone_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles textPhone.Enter
      textPhone.SelectionStart = 0
      textPhone.SelectionLength = textPhone.Text.Trim.Length
   End Sub

   ''' <summary>
   ''' Place the selected item on rent.
   ''' </summary>
   Public Function RentTheItem()
      Dim sql As String
      Dim period As String
      Dim price As Decimal
      Dim equipName As String = Me.textReRentEquip.Text
      Dim newPrices As Boolean = False
      With Me
         If UnFormat(.txtDaily.Text) > 0 And _
           UnFormat(.txtWeekly.Text) > 0 And _
           UnFormat(.txtMonthly.Text) > 0 Then
            newPrices = True
         End If
         Select Case True
            Case .optHour.Checked
               If UnFormat(.txtHourly.Text) = 0 Then
                  GoTo norates
               End If
               period = HOURLY
               price = UnFormat(.txtHourly.Text)
            Case .optDaily.Checked
               If UnFormat(.txtDaily.Text) = 0 Then
                  GoTo norates
               End If
               period = DAILY
               price = UnFormat(.txtDaily.Text)
            Case .optWeekly.Checked
               If UnFormat(.txtWeekly.Text) = 0 Then
                  GoTo norates
               End If
               period = WEEKLY
               price = UnFormat(.txtWeekly.Text)
            Case .optMonthly.Checked
               If UnFormat(.txtMonthly.Text) = 0 Then
                  GoTo norates
               End If
               period = MONTHLY
               price = UnFormat(.txtMonthly.Text)
            Case .optHalfDay.Checked
               If UnFormat(.txtHalfDay.Text) = 0 Then
                  GoTo norates
               End If
               period = HALF_DAY
               price = UnFormat(.txtHalfDay.Text)
            Case .optWeekEnd.Checked
               If UnFormat(.txtWeekEnd.Text) = 0 Then
                  GoTo norates
               End If
               period = WEEK_END
               price = UnFormat(.txtWeekEnd.Text)
         End Select
      End With

      If equipName.Length > 29 Then
         equipName = equipName.Substring(0, 29)
      End If
      sql = "Insert into TempItems (ItemID,ItemName,ItemCount, "
      sql &= "ItemPeriod,ItemPrice,ItemExtendedPrice,ItemDeposit,"
      sql &= "rentorsale,meter_required,hour_meter,user_id,rerent_id"
      sql &= ",hourrate,halfday,daily,weekly,monthly,weekend,newprices ) "
      sql &= "Values("
      sql &= "'" & "ReRent" & "', "
      sql &= "'" & Replace(equipName, "'", "''") & "', "
      sql &= Val(Me.cbNbrPeriods.Text) & ", "
      Dim iPeriods As Integer = Val(Me.cbNbrPeriods.Text)

      Dim decPrice As Decimal = price
      Dim sCurrTP As String = period
      sql &= "'" & sCurrTP & "', "
      sql &= (decPrice) & ", " ' price
      decPrice *= iPeriods
      sql &= (decPrice) & ", " ' and extended price are the same
      sql &= "0, " ' deposit
      sql &= "'" & Me.textPO.Text & "', "
      sql &= False & ", "
      sql &= "0" ' hours
      sql &= ",'" & UserName & "',"
      sql &= Me.lblRecordID.Text & ", "
      sql &= UnFormat(Me.txtHourly.Text) & ", "
      sql &= UnFormat(Me.txtHalfDay.Text) & ", "
      sql &= UnFormat(Me.txtDaily.Text) & ", "
      sql &= UnFormat(Me.txtWeekly.Text) & ", "
      sql &= UnFormat(Me.txtMonthly.Text) & ", "
      sql &= UnFormat(Me.txtWeekEnd.Text) & ", "
      sql &= newPrices & " "
      sql &= ")"
      ' ensure that this item is being rented  for the same 
      ' time period as any other items
      Dim sql2 As String
      Dim dt As New DataTable()
      sql2 = "select * from tempitems where (rentorsale = '" & RENT & "' "
      sql2 &= "or itemid = '" & RERENT & "') "
      sql2 &= "and user_id = '" & UserName & "'"
      dt.Reset()
      Dim i As Integer
      Dim sLastTP As String
      If oDA.SendQuery(sql2, dt, ConnectString) > 0 Then
         For i = 0 To dt.Rows.Count - 1
            With dt.Rows(i)
               sLastTP = MNS(.Item("itemperiod"))
               If sLastTP.Trim <> sCurrTP Then
                  MsgBox("The time period for all rented items must be the same for all items on an invoice.", MsgBoxStyle.Exclamation)
                  Return False
               End If
            End With
         Next
      End If

      Dim sErr As String
      If oDA.SendActionSql(sql, ConnectString, sErr) <> 1 Then
         Throw New System.Exception("Database error, unable to create temp invoice item.")
      End If
      Return True
      Exit Function
NoRates:
      MsgBox("You must enter a cost for the selected period.", MsgBoxStyle.Information)
      Return False
   End Function

   ''' <summary>
   ''' Place rerent item on rent.
   ''' 1) Add row to tempitems for the rerent equipment.
   ''' 2) Mark the selected item as rented which will 
   '''    remove it from the grid for viewing purposes.
   ''' </summary>
   ''' <param name = "sender"></param>
   ''' <param name = "e"></param>
   Private Sub btnRent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRent.Click

      Dim sql As String
      Dim serr As String
      Dim i As Integer
      Dim dr As DataRow

      Try
         Dim sMsg As String
         Dim iRV As Integer
         sMsg = "Are you sure you want to rent: " & Me.textReRentEquip.Text & "?"
         sMsg &= "" & Chr(10)
         iRV = MsgBox(sMsg, CType(36, Microsoft.VisualBasic.MsgBoxStyle), "Confirm Rental")

         If iRV = 6 Then
            ' Yes Code goes here
         Else
            ' No code goes here
            Exit Sub
         End If
         If Not RentTheItem() Then
            Exit Sub
         End If
         sql = "update rerents set rented_date = #" & Now.ToString & "# where unique_id = " & Me.lblRecordID.Text
         If oDA.SendActionSql(sql, ConnectString, serr) <> 1 Then
            Throw New System.Exception("Update failed to mark rerent as rented: " & Me.lblRecordID.Text & Chr(10) & serr)
         End If
         LoadTheGrid()
         LoadTextBoxes()
         If Me.mnuCloseAfterRent.Checked Then
            Me.Close()
            System.Windows.Forms.Application.DoEvents()
         End If
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub
   ''' <summary>
   ''' Clear the Rented Date so that the item will not be rented.
   ''' </summary>
   ''' <param name = "sender"></param>
   ''' <param name = "e"></param>
   Private Sub Restore()
      Dim sql As String
      Dim serr As String
      Dim i As Integer
      Dim dr As DataRow

      Try
         Dim sMsg As String
         Dim iRV As Integer
         sMsg = "Are you sure you want to unrent: " & Me.textReRentEquip.Text & "?"
         sMsg &= "" & Chr(10)
         iRV = MsgBox(sMsg, CType(36, Microsoft.VisualBasic.MsgBoxStyle), "Confirm Marking as Not Rented")

         If iRV = 6 Then
            ' Yes Code goes here
         Else
            ' No code goes here
            Exit Sub
         End If
         sql = "update rerents set rented_date = Null where unique_id = " & Me.lblRecordID.Text
         If oDA.SendActionSql(sql, ConnectString, serr) <> 1 Then
            Throw New System.Exception("Update failed to mark rerent as unrented: " & Me.lblRecordID.Text & Chr(10) & serr)
         End If
         LoadTheGrid()
         LoadTextBoxes()
         'System.Windows.Forms.Application.DoEvents()
         'Me.Close()
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub
   Private Sub frmRerent_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
      With Me
         .mnuCloseAfterRent.Checked = GetSetting(RENTALPRO, SETTINGS, "CLOSERERENT", True)

         If modMain.UseHourlyRates Then
            .optHour.Visible = True
            .txtHourly.Visible = True
         End If
         If modMain.UseWeekEndRates Then
            .optWeekEnd.Visible = True
            .txtWeekEnd.Visible = True
         End If
      End With
      LoadTheGrid()
      LoadTextBoxes()
      LoadCustomerCombo()
   End Sub
   Private Sub chkShowAllItems_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowAllItems.CheckedChanged
      With Me.chkShowAllItems
         LoadTheGrid(ShowAll:=.Checked)
         LoadTextBoxes()
      End With
   End Sub
   Private Sub mnuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExit.Click
      Me.Close()
      System.Windows.Forms.Application.DoEvents()
   End Sub
   Private Sub mnuRestoreItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRestoreItem.Click
      Restore()
   End Sub
#End Region

End Class
