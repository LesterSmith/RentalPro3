Imports System.Windows.Forms.Application

Public Class frmReservations
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
   Public WithEvents btnClose2 As System.Windows.Forms.Button
   Friend WithEvents btnPrint As System.Windows.Forms.Button
   Friend WithEvents cbCust As System.Windows.Forms.ComboBox
   Friend WithEvents cbNbrPeriods As System.Windows.Forms.ComboBox
   Friend WithEvents cbPeriod As System.Windows.Forms.ComboBox
   Public WithEvents cmdAdd As System.Windows.Forms.Button
   Public WithEvents cmdClose As System.Windows.Forms.Button
   Public WithEvents cmdDelete As System.Windows.Forms.Button
   Friend WithEvents dgEquip As System.Windows.Forms.DataGrid
   Friend WithEvents dgReservations As System.Windows.Forms.DataGrid
   Friend WithEvents dtpStartRes As System.Windows.Forms.DateTimePicker
   Friend WithEvents grpAddRes As System.Windows.Forms.GroupBox
   Friend WithEvents Label1 As System.Windows.Forms.Label
   Friend WithEvents Label2 As System.Windows.Forms.Label
   Friend WithEvents Label3 As System.Windows.Forms.Label
   Friend WithEvents Label4 As System.Windows.Forms.Label
   Friend WithEvents Label5 As System.Windows.Forms.Label
   Friend WithEvents Label6 As System.Windows.Forms.Label
   Public WithEvents lblCustId As System.Windows.Forms.Label
   Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
   Friend WithEvents tabCurrRes As System.Windows.Forms.TabPage
   Friend WithEvents tabMakeRes As System.Windows.Forms.TabPage
   Friend WithEvents txtContact As System.Windows.Forms.TextBox
   Public WithEvents txtCustomerID As System.Windows.Forms.TextBox
   Friend WithEvents txtNotes As System.Windows.Forms.TextBox
   Friend WithEvents txtPhone As System.Windows.Forms.TextBox
   Friend WithEvents lblEName As System.Windows.Forms.Label
   Friend WithEvents lblEquipName As System.Windows.Forms.Label
   Friend WithEvents lblNPeriods As System.Windows.Forms.Label
   Friend WithEvents lblResPer As System.Windows.Forms.Label
   Friend WithEvents lblResDate As System.Windows.Forms.Label
   Friend WithEvents lblEDate As System.Windows.Forms.Label
   Friend WithEvents lblNumberPeriods As System.Windows.Forms.Label
   Friend WithEvents lblReservationPeriod As System.Windows.Forms.Label
   Friend WithEvents lblReservationDate As System.Windows.Forms.Label
   Friend WithEvents lblEndDate As System.Windows.Forms.Label
   Friend WithEvents lblCustName As System.Windows.Forms.Label
   Friend WithEvents lblCont As System.Windows.Forms.Label
   Friend WithEvents lblPNbr As System.Windows.Forms.Label
   Friend WithEvents lblCustomerName As System.Windows.Forms.Label
   Friend WithEvents lblContactName As System.Windows.Forms.Label
   Friend WithEvents lblPhoneNumber As System.Windows.Forms.Label
   Friend WithEvents Label10 As System.Windows.Forms.Label
   Friend WithEvents textResNotes As System.Windows.Forms.TextBox
   Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
   Friend WithEvents chkCashCustomer As System.Windows.Forms.CheckBox
   Friend WithEvents btnAddnClose As System.Windows.Forms.Button
   Friend WithEvents btnAddCustomer As System.Windows.Forms.Button
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
      Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmReservations))
      Me.cmdAdd = New System.Windows.Forms.Button()
      Me.cmdDelete = New System.Windows.Forms.Button()
      Me.cmdClose = New System.Windows.Forms.Button()
      Me.dgReservations = New System.Windows.Forms.DataGrid()
      Me.dtpStartRes = New System.Windows.Forms.DateTimePicker()
      Me.Label1 = New System.Windows.Forms.Label()
      Me.cbNbrPeriods = New System.Windows.Forms.ComboBox()
      Me.Label2 = New System.Windows.Forms.Label()
      Me.cbPeriod = New System.Windows.Forms.ComboBox()
      Me.Label3 = New System.Windows.Forms.Label()
      Me.grpAddRes = New System.Windows.Forms.GroupBox()
      Me.TabControl1 = New System.Windows.Forms.TabControl()
      Me.tabCurrRes = New System.Windows.Forms.TabPage()
      Me.textResNotes = New System.Windows.Forms.TextBox()
      Me.Label10 = New System.Windows.Forms.Label()
      Me.lblPhoneNumber = New System.Windows.Forms.Label()
      Me.lblContactName = New System.Windows.Forms.Label()
      Me.lblCustomerName = New System.Windows.Forms.Label()
      Me.lblPNbr = New System.Windows.Forms.Label()
      Me.lblCont = New System.Windows.Forms.Label()
      Me.lblCustName = New System.Windows.Forms.Label()
      Me.lblEndDate = New System.Windows.Forms.Label()
      Me.lblReservationDate = New System.Windows.Forms.Label()
      Me.lblReservationPeriod = New System.Windows.Forms.Label()
      Me.lblNumberPeriods = New System.Windows.Forms.Label()
      Me.lblEDate = New System.Windows.Forms.Label()
      Me.lblResDate = New System.Windows.Forms.Label()
      Me.lblResPer = New System.Windows.Forms.Label()
      Me.lblNPeriods = New System.Windows.Forms.Label()
      Me.lblEquipName = New System.Windows.Forms.Label()
      Me.lblEName = New System.Windows.Forms.Label()
      Me.btnPrint = New System.Windows.Forms.Button()
      Me.btnClose2 = New System.Windows.Forms.Button()
      Me.tabMakeRes = New System.Windows.Forms.TabPage()
      Me.btnAddnClose = New System.Windows.Forms.Button()
      Me.GroupBox1 = New System.Windows.Forms.GroupBox()
      Me.Label4 = New System.Windows.Forms.Label()
      Me.Label5 = New System.Windows.Forms.Label()
      Me._lblFieldLable_0 = New System.Windows.Forms.Label()
      Me.txtContact = New System.Windows.Forms.TextBox()
      Me.txtCustomerID = New System.Windows.Forms.TextBox()
      Me.lblCustId = New System.Windows.Forms.Label()
      Me.cbCust = New System.Windows.Forms.ComboBox()
      Me.txtPhone = New System.Windows.Forms.TextBox()
      Me.Label6 = New System.Windows.Forms.Label()
      Me.txtNotes = New System.Windows.Forms.TextBox()
      Me.chkCashCustomer = New System.Windows.Forms.CheckBox()
      Me.dgEquip = New System.Windows.Forms.DataGrid()
      Me.btnAddCustomer = New System.Windows.Forms.Button()
      CType(Me.dgReservations, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.grpAddRes.SuspendLayout()
      Me.TabControl1.SuspendLayout()
      Me.tabCurrRes.SuspendLayout()
      Me.tabMakeRes.SuspendLayout()
      Me.GroupBox1.SuspendLayout()
      CType(Me.dgEquip, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.SuspendLayout()
      '
      'cmdAdd
      '
      Me.cmdAdd.BackColor = System.Drawing.SystemColors.Control
      Me.cmdAdd.Cursor = System.Windows.Forms.Cursors.Default
      Me.cmdAdd.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.cmdAdd.ForeColor = System.Drawing.SystemColors.ControlText
      Me.cmdAdd.Location = New System.Drawing.Point(408, 372)
      Me.cmdAdd.Name = "cmdAdd"
      Me.cmdAdd.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.cmdAdd.Size = New System.Drawing.Size(57, 24)
      Me.cmdAdd.TabIndex = 1
      Me.cmdAdd.Text = "&Add"
      '
      'cmdDelete
      '
      Me.cmdDelete.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
      Me.cmdDelete.BackColor = System.Drawing.SystemColors.Control
      Me.cmdDelete.Cursor = System.Windows.Forms.Cursors.Default
      Me.cmdDelete.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.cmdDelete.ForeColor = System.Drawing.SystemColors.ControlText
      Me.cmdDelete.Location = New System.Drawing.Point(483, 315)
      Me.cmdDelete.Name = "cmdDelete"
      Me.cmdDelete.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.cmdDelete.Size = New System.Drawing.Size(57, 24)
      Me.cmdDelete.TabIndex = 26
      Me.cmdDelete.Text = "&Remove"
      '
      'cmdClose
      '
      Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
      Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
      Me.cmdClose.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
      Me.cmdClose.Location = New System.Drawing.Point(480, 372)
      Me.cmdClose.Name = "cmdClose"
      Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.cmdClose.Size = New System.Drawing.Size(57, 24)
      Me.cmdClose.TabIndex = 2
      Me.cmdClose.Text = "&Close"
      '
      'dgReservations
      '
      Me.dgReservations.AllowSorting = False
      Me.dgReservations.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                  Or System.Windows.Forms.AnchorStyles.Left) _
                  Or System.Windows.Forms.AnchorStyles.Right)
      Me.dgReservations.DataMember = ""
      Me.dgReservations.HeaderForeColor = System.Drawing.SystemColors.ControlText
      Me.dgReservations.Location = New System.Drawing.Point(7, 4)
      Me.dgReservations.Name = "dgReservations"
      Me.dgReservations.Size = New System.Drawing.Size(544, 254)
      Me.dgReservations.TabIndex = 23
      '
      'dtpStartRes
      '
      Me.dtpStartRes.CustomFormat = "MM/dd/yyyy HH:mm tt"
      Me.dtpStartRes.Format = System.Windows.Forms.DateTimePickerFormat.Custom
      Me.dtpStartRes.Location = New System.Drawing.Point(11, 35)
      Me.dtpStartRes.Name = "dtpStartRes"
      Me.dtpStartRes.Size = New System.Drawing.Size(197, 20)
      Me.dtpStartRes.TabIndex = 0
      '
      'Label1
      '
      Me.Label1.AutoSize = True
      Me.Label1.Location = New System.Drawing.Point(3, 19)
      Me.Label1.Name = "Label1"
      Me.Label1.Size = New System.Drawing.Size(92, 13)
      Me.Label1.TabIndex = 29
      Me.Label1.Text = "Reservation Start"
      '
      'cbNbrPeriods
      '
      Me.cbNbrPeriods.Items.AddRange(New Object() {"1", "2", "3", "4", "5"})
      Me.cbNbrPeriods.Location = New System.Drawing.Point(13, 80)
      Me.cbNbrPeriods.Name = "cbNbrPeriods"
      Me.cbNbrPeriods.Size = New System.Drawing.Size(64, 21)
      Me.cbNbrPeriods.TabIndex = 1
      Me.cbNbrPeriods.Text = "1"
      '
      'Label2
      '
      Me.Label2.AutoSize = True
      Me.Label2.Location = New System.Drawing.Point(13, 64)
      Me.Label2.Name = "Label2"
      Me.Label2.Size = New System.Drawing.Size(86, 13)
      Me.Label2.TabIndex = 31
      Me.Label2.Text = "Number Periods"
      '
      'cbPeriod
      '
      Me.cbPeriod.Items.AddRange(New Object() {"Daily", "Hourly", "Half Day", "Monthly", "Weekly", "Week End"})
      Me.cbPeriod.Location = New System.Drawing.Point(128, 80)
      Me.cbPeriod.Name = "cbPeriod"
      Me.cbPeriod.Size = New System.Drawing.Size(72, 21)
      Me.cbPeriod.TabIndex = 2
      Me.cbPeriod.Text = "Daily"
      '
      'Label3
      '
      Me.Label3.AutoSize = True
      Me.Label3.Location = New System.Drawing.Point(128, 64)
      Me.Label3.Name = "Label3"
      Me.Label3.Size = New System.Drawing.Size(37, 13)
      Me.Label3.TabIndex = 33
      Me.Label3.Text = "Period"
      '
      'grpAddRes
      '
      Me.grpAddRes.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label3, Me.dtpStartRes, Me.cbNbrPeriods, Me.Label2, Me.Label1, Me.cbPeriod})
      Me.grpAddRes.Location = New System.Drawing.Point(267, 8)
      Me.grpAddRes.Name = "grpAddRes"
      Me.grpAddRes.Size = New System.Drawing.Size(285, 112)
      Me.grpAddRes.TabIndex = 34
      Me.grpAddRes.TabStop = False
      Me.grpAddRes.Text = "Add Reservation Data"
      '
      'TabControl1
      '
      Me.TabControl1.Controls.AddRange(New System.Windows.Forms.Control() {Me.tabCurrRes, Me.tabMakeRes})
      Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
      Me.TabControl1.Name = "TabControl1"
      Me.TabControl1.SelectedIndex = 0
      Me.TabControl1.Size = New System.Drawing.Size(568, 430)
      Me.TabControl1.TabIndex = 35
      '
      'tabCurrRes
      '
      Me.tabCurrRes.Controls.AddRange(New System.Windows.Forms.Control() {Me.textResNotes, Me.Label10, Me.lblPhoneNumber, Me.lblContactName, Me.lblCustomerName, Me.lblPNbr, Me.lblCont, Me.lblCustName, Me.lblEndDate, Me.lblReservationDate, Me.lblReservationPeriod, Me.lblNumberPeriods, Me.lblEDate, Me.lblResDate, Me.lblResPer, Me.lblNPeriods, Me.lblEquipName, Me.lblEName, Me.btnPrint, Me.btnClose2, Me.dgReservations, Me.cmdDelete})
      Me.tabCurrRes.Location = New System.Drawing.Point(4, 22)
      Me.tabCurrRes.Name = "tabCurrRes"
      Me.tabCurrRes.Size = New System.Drawing.Size(560, 404)
      Me.tabCurrRes.TabIndex = 0
      Me.tabCurrRes.Text = "Currently Reserved"
      '
      'textResNotes
      '
      Me.textResNotes.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
      Me.textResNotes.Location = New System.Drawing.Point(272, 335)
      Me.textResNotes.Multiline = True
      Me.textResNotes.Name = "textResNotes"
      Me.textResNotes.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
      Me.textResNotes.Size = New System.Drawing.Size(184, 59)
      Me.textResNotes.TabIndex = 56
      Me.textResNotes.Text = ""
      '
      'Label10
      '
      Me.Label10.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
      Me.Label10.AutoSize = True
      Me.Label10.Location = New System.Drawing.Point(232, 340)
      Me.Label10.Name = "Label10"
      Me.Label10.Size = New System.Drawing.Size(31, 13)
      Me.Label10.TabIndex = 55
      Me.Label10.Text = "Note:"
      '
      'lblPhoneNumber
      '
      Me.lblPhoneNumber.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
      Me.lblPhoneNumber.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
      Me.lblPhoneNumber.Location = New System.Drawing.Point(274, 310)
      Me.lblPhoneNumber.Name = "lblPhoneNumber"
      Me.lblPhoneNumber.Size = New System.Drawing.Size(182, 16)
      Me.lblPhoneNumber.TabIndex = 52
      '
      'lblContactName
      '
      Me.lblContactName.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
      Me.lblContactName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
      Me.lblContactName.Location = New System.Drawing.Point(274, 290)
      Me.lblContactName.Name = "lblContactName"
      Me.lblContactName.Size = New System.Drawing.Size(182, 16)
      Me.lblContactName.TabIndex = 51
      '
      'lblCustomerName
      '
      Me.lblCustomerName.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
      Me.lblCustomerName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
      Me.lblCustomerName.Location = New System.Drawing.Point(274, 270)
      Me.lblCustomerName.Name = "lblCustomerName"
      Me.lblCustomerName.Size = New System.Drawing.Size(182, 16)
      Me.lblCustomerName.TabIndex = 50
      '
      'lblPNbr
      '
      Me.lblPNbr.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
      Me.lblPNbr.AutoSize = True
      Me.lblPNbr.Location = New System.Drawing.Point(218, 310)
      Me.lblPNbr.Name = "lblPNbr"
      Me.lblPNbr.Size = New System.Drawing.Size(37, 13)
      Me.lblPNbr.TabIndex = 49
      Me.lblPNbr.Text = "Phone"
      '
      'lblCont
      '
      Me.lblCont.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
      Me.lblCont.AutoSize = True
      Me.lblCont.Location = New System.Drawing.Point(218, 291)
      Me.lblCont.Name = "lblCont"
      Me.lblCont.Size = New System.Drawing.Size(43, 13)
      Me.lblCont.TabIndex = 48
      Me.lblCont.Text = "Contact"
      '
      'lblCustName
      '
      Me.lblCustName.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
      Me.lblCustName.AutoSize = True
      Me.lblCustName.Location = New System.Drawing.Point(218, 270)
      Me.lblCustName.Name = "lblCustName"
      Me.lblCustName.Size = New System.Drawing.Size(53, 13)
      Me.lblCustName.TabIndex = 47
      Me.lblCustName.Text = "Customer"
      '
      'lblEndDate
      '
      Me.lblEndDate.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
      Me.lblEndDate.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
      Me.lblEndDate.Location = New System.Drawing.Point(88, 350)
      Me.lblEndDate.Name = "lblEndDate"
      Me.lblEndDate.Size = New System.Drawing.Size(117, 16)
      Me.lblEndDate.TabIndex = 46
      '
      'lblReservationDate
      '
      Me.lblReservationDate.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
      Me.lblReservationDate.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
      Me.lblReservationDate.Location = New System.Drawing.Point(88, 330)
      Me.lblReservationDate.Name = "lblReservationDate"
      Me.lblReservationDate.Size = New System.Drawing.Size(117, 16)
      Me.lblReservationDate.TabIndex = 45
      '
      'lblReservationPeriod
      '
      Me.lblReservationPeriod.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
      Me.lblReservationPeriod.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
      Me.lblReservationPeriod.Location = New System.Drawing.Point(88, 310)
      Me.lblReservationPeriod.Name = "lblReservationPeriod"
      Me.lblReservationPeriod.Size = New System.Drawing.Size(117, 16)
      Me.lblReservationPeriod.TabIndex = 44
      '
      'lblNumberPeriods
      '
      Me.lblNumberPeriods.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
      Me.lblNumberPeriods.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
      Me.lblNumberPeriods.Location = New System.Drawing.Point(88, 290)
      Me.lblNumberPeriods.Name = "lblNumberPeriods"
      Me.lblNumberPeriods.Size = New System.Drawing.Size(117, 16)
      Me.lblNumberPeriods.TabIndex = 43
      '
      'lblEDate
      '
      Me.lblEDate.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
      Me.lblEDate.AutoSize = True
      Me.lblEDate.Location = New System.Drawing.Point(29, 350)
      Me.lblEDate.Name = "lblEDate"
      Me.lblEDate.Size = New System.Drawing.Size(51, 13)
      Me.lblEDate.TabIndex = 42
      Me.lblEDate.Text = "End Date"
      '
      'lblResDate
      '
      Me.lblResDate.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
      Me.lblResDate.AutoSize = True
      Me.lblResDate.Location = New System.Drawing.Point(29, 330)
      Me.lblResDate.Name = "lblResDate"
      Me.lblResDate.Size = New System.Drawing.Size(51, 13)
      Me.lblResDate.TabIndex = 41
      Me.lblResDate.Text = "Res Date"
      '
      'lblResPer
      '
      Me.lblResPer.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
      Me.lblResPer.AutoSize = True
      Me.lblResPer.Location = New System.Drawing.Point(20, 310)
      Me.lblResPer.Name = "lblResPer"
      Me.lblResPer.Size = New System.Drawing.Size(60, 13)
      Me.lblResPer.TabIndex = 40
      Me.lblResPer.Text = "Res Period"
      '
      'lblNPeriods
      '
      Me.lblNPeriods.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
      Me.lblNPeriods.AutoSize = True
      Me.lblNPeriods.Location = New System.Drawing.Point(16, 290)
      Me.lblNPeriods.Name = "lblNPeriods"
      Me.lblNPeriods.Size = New System.Drawing.Size(64, 13)
      Me.lblNPeriods.TabIndex = 39
      Me.lblNPeriods.Text = "Nbr Periods"
      '
      'lblEquipName
      '
      Me.lblEquipName.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
      Me.lblEquipName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
      Me.lblEquipName.Location = New System.Drawing.Point(88, 270)
      Me.lblEquipName.Name = "lblEquipName"
      Me.lblEquipName.Size = New System.Drawing.Size(117, 16)
      Me.lblEquipName.TabIndex = 38
      '
      'lblEName
      '
      Me.lblEName.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
      Me.lblEName.AutoSize = True
      Me.lblEName.Location = New System.Drawing.Point(14, 270)
      Me.lblEName.Name = "lblEName"
      Me.lblEName.Size = New System.Drawing.Size(66, 13)
      Me.lblEName.TabIndex = 37
      Me.lblEName.Text = "Equip Name"
      '
      'btnPrint
      '
      Me.btnPrint.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
      Me.btnPrint.Location = New System.Drawing.Point(483, 283)
      Me.btnPrint.Name = "btnPrint"
      Me.btnPrint.Size = New System.Drawing.Size(57, 24)
      Me.btnPrint.TabIndex = 36
      Me.btnPrint.Text = "&Print"
      Me.btnPrint.Visible = False
      '
      'btnClose2
      '
      Me.btnClose2.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
      Me.btnClose2.BackColor = System.Drawing.SystemColors.Control
      Me.btnClose2.Cursor = System.Windows.Forms.Cursors.Default
      Me.btnClose2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.btnClose2.ForeColor = System.Drawing.SystemColors.ControlText
      Me.btnClose2.Location = New System.Drawing.Point(483, 347)
      Me.btnClose2.Name = "btnClose2"
      Me.btnClose2.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.btnClose2.Size = New System.Drawing.Size(57, 24)
      Me.btnClose2.TabIndex = 28
      Me.btnClose2.Text = "&Close"
      '
      'tabMakeRes
      '
      Me.tabMakeRes.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnAddnClose, Me.GroupBox1, Me.dgEquip, Me.grpAddRes, Me.cmdAdd, Me.cmdClose})
      Me.tabMakeRes.Location = New System.Drawing.Point(4, 22)
      Me.tabMakeRes.Name = "tabMakeRes"
      Me.tabMakeRes.Size = New System.Drawing.Size(560, 404)
      Me.tabMakeRes.TabIndex = 1
      Me.tabMakeRes.Text = "Make New Reservation"
      '
      'btnAddnClose
      '
      Me.btnAddnClose.Location = New System.Drawing.Point(296, 372)
      Me.btnAddnClose.Name = "btnAddnClose"
      Me.btnAddnClose.Size = New System.Drawing.Size(96, 24)
      Me.btnAddnClose.TabIndex = 0
      Me.btnAddnClose.Text = "Add && C&lose"
      '
      'GroupBox1
      '
      Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnAddCustomer, Me.Label4, Me.Label5, Me._lblFieldLable_0, Me.txtContact, Me.txtCustomerID, Me.lblCustId, Me.cbCust, Me.txtPhone, Me.Label6, Me.txtNotes, Me.chkCashCustomer})
      Me.GroupBox1.Location = New System.Drawing.Point(267, 123)
      Me.GroupBox1.Name = "GroupBox1"
      Me.GroupBox1.Size = New System.Drawing.Size(285, 242)
      Me.GroupBox1.TabIndex = 46
      Me.GroupBox1.TabStop = False
      Me.GroupBox1.Text = "Customer Data"
      '
      'Label4
      '
      Me.Label4.AutoSize = True
      Me.Label4.Location = New System.Drawing.Point(14, 118)
      Me.Label4.Name = "Label4"
      Me.Label4.Size = New System.Drawing.Size(37, 13)
      Me.Label4.TabIndex = 40
      Me.Label4.Text = "Phone"
      '
      'Label5
      '
      Me.Label5.AutoSize = True
      Me.Label5.Location = New System.Drawing.Point(14, 102)
      Me.Label5.Name = "Label5"
      Me.Label5.Size = New System.Drawing.Size(43, 13)
      Me.Label5.TabIndex = 42
      Me.Label5.Text = "Contact"
      '
      '_lblFieldLable_0
      '
      Me._lblFieldLable_0.BackColor = System.Drawing.SystemColors.Control
      Me._lblFieldLable_0.Cursor = System.Windows.Forms.Cursors.Default
      Me._lblFieldLable_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me._lblFieldLable_0.ForeColor = System.Drawing.SystemColors.ControlText
      Me._lblFieldLable_0.Location = New System.Drawing.Point(5, 40)
      Me._lblFieldLable_0.Name = "_lblFieldLable_0"
      Me._lblFieldLable_0.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me._lblFieldLable_0.Size = New System.Drawing.Size(56, 27)
      Me._lblFieldLable_0.TabIndex = 39
      Me._lblFieldLable_0.Text = "Select Customer"
      '
      'txtContact
      '
      Me.txtContact.Location = New System.Drawing.Point(61, 99)
      Me.txtContact.Name = "txtContact"
      Me.txtContact.Size = New System.Drawing.Size(136, 20)
      Me.txtContact.TabIndex = 3
      Me.txtContact.Text = ""
      '
      'txtCustomerID
      '
      Me.txtCustomerID.AcceptsReturn = True
      Me.txtCustomerID.AutoSize = False
      Me.txtCustomerID.BackColor = System.Drawing.SystemColors.Window
      Me.txtCustomerID.Cursor = System.Windows.Forms.Cursors.IBeam
      Me.txtCustomerID.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.txtCustomerID.ForeColor = System.Drawing.SystemColors.WindowText
      Me.txtCustomerID.Location = New System.Drawing.Point(61, 75)
      Me.txtCustomerID.MaxLength = 0
      Me.txtCustomerID.Name = "txtCustomerID"
      Me.txtCustomerID.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.txtCustomerID.Size = New System.Drawing.Size(76, 21)
      Me.txtCustomerID.TabIndex = 2
      Me.txtCustomerID.Text = ""
      Me.txtCustomerID.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      '
      'lblCustId
      '
      Me.lblCustId.BackColor = System.Drawing.SystemColors.Control
      Me.lblCustId.Cursor = System.Windows.Forms.Cursors.Default
      Me.lblCustId.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.lblCustId.ForeColor = System.Drawing.SystemColors.ControlText
      Me.lblCustId.Location = New System.Drawing.Point(6, 72)
      Me.lblCustId.Name = "lblCustId"
      Me.lblCustId.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.lblCustId.Size = New System.Drawing.Size(57, 24)
      Me.lblCustId.TabIndex = 38
      Me.lblCustId.Text = "Customer          ID"
      '
      'cbCust
      '
      Me.cbCust.Location = New System.Drawing.Point(61, 43)
      Me.cbCust.Name = "cbCust"
      Me.cbCust.Size = New System.Drawing.Size(215, 21)
      Me.cbCust.TabIndex = 1
      '
      'txtPhone
      '
      Me.txtPhone.Location = New System.Drawing.Point(61, 123)
      Me.txtPhone.Name = "txtPhone"
      Me.txtPhone.Size = New System.Drawing.Size(96, 20)
      Me.txtPhone.TabIndex = 4
      Me.txtPhone.Text = ""
      '
      'Label6
      '
      Me.Label6.Location = New System.Drawing.Point(8, 160)
      Me.Label6.Name = "Label6"
      Me.Label6.Size = New System.Drawing.Size(47, 27)
      Me.Label6.TabIndex = 44
      Me.Label6.Text = "Optional Note:"
      '
      'txtNotes
      '
      Me.txtNotes.Location = New System.Drawing.Point(60, 156)
      Me.txtNotes.Multiline = True
      Me.txtNotes.Name = "txtNotes"
      Me.txtNotes.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
      Me.txtNotes.Size = New System.Drawing.Size(212, 72)
      Me.txtNotes.TabIndex = 5
      Me.txtNotes.Text = ""
      '
      'chkCashCustomer
      '
      Me.chkCashCustomer.ForeColor = System.Drawing.Color.Red
      Me.chkCashCustomer.Location = New System.Drawing.Point(61, 16)
      Me.chkCashCustomer.Name = "chkCashCustomer"
      Me.chkCashCustomer.Size = New System.Drawing.Size(104, 15)
      Me.chkCashCustomer.TabIndex = 0
      Me.chkCashCustomer.Text = "Cash Customer"
      '
      'dgEquip
      '
      Me.dgEquip.DataMember = ""
      Me.dgEquip.HeaderForeColor = System.Drawing.SystemColors.ControlText
      Me.dgEquip.Location = New System.Drawing.Point(8, 8)
      Me.dgEquip.Name = "dgEquip"
      Me.dgEquip.Size = New System.Drawing.Size(252, 389)
      Me.dgEquip.TabIndex = 0
      '
      'btnAddCustomer
      '
      Me.btnAddCustomer.Location = New System.Drawing.Point(176, 12)
      Me.btnAddCustomer.Name = "btnAddCustomer"
      Me.btnAddCustomer.Size = New System.Drawing.Size(96, 24)
      Me.btnAddCustomer.TabIndex = 45
      Me.btnAddCustomer.Text = "Add C&ustomer"
      '
      'frmReservations
      '
      Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
      Me.ClientSize = New System.Drawing.Size(568, 430)
      Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabControl1})
      Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
      Me.Name = "frmReservations"
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "Reservations"
      CType(Me.dgReservations, System.ComponentModel.ISupportInitialize).EndInit()
      Me.grpAddRes.ResumeLayout(False)
      Me.TabControl1.ResumeLayout(False)
      Me.tabCurrRes.ResumeLayout(False)
      Me.tabMakeRes.ResumeLayout(False)
      Me.GroupBox1.ResumeLayout(False)
      CType(Me.dgEquip, System.ComponentModel.ISupportInitialize).EndInit()
      Me.ResumeLayout(False)

   End Sub

#End Region

#Region "Private Variables"
   Private oDA As New CDataAccess()
   Dim SQL As String
   Private dtRes As DataTable
   Private dtEquip As DataTable
   Private iHitRow As Integer
   Private oCG As New CGrid()
   Private equipHitRow As Integer
   Private resHitRow As Integer
   Private Noise As Boolean
#End Region

#Region "Helper Methods"
   Private Sub AddEquipment(ByVal closeForm As Boolean)

      Dim SQL As String
      Dim sErr As String
      Dim resEndDate As DateTime
      Dim oCE As New CCheckReservations()
      Dim equipName As String

      Try

         With Me
            If .dgEquip.DataSource.rows(equipHitRow).item("Reserve") = "false" Then
               MsgBox("You must select equipment to reserve.", MsgBoxStyle.Exclamation)
               Exit Sub
            End If

            If Me.cbCust.Text.Trim.Length > 0 And _
               Me.txtCustomerID.Text.Trim.Length > 0 Then
               If Val(.txtCustomerID.Text) = 0 Then
                  If .txtContact.Text.Trim.Length = 0 Or _
                     .txtPhone.Text.Trim.Length = 0 Then
                     MsgBox("You must enter a contact name and phone number for the cash customer.", MsgBoxStyle.Information)
                     Exit Sub
                  End If
               End If
            Else
               MsgBox("You must select a customer for the reservation.", MsgBoxStyle.Exclamation)
               Exit Sub
            End If

            ' the default time should be 8am or 1pm or 5pm???
            Dim s As String = Format(.dtpStartRes.Value, "MM/dd/yyyy HH:mm tt") ' 08/19/2003 08:00 AM
            If s.IndexOf("8:00 AM") = -1 AndAlso _
               s.IndexOf("1:00 PM") = -1 AndAlso _
               s.IndexOf("5:00 PM") = -1 Then
               Dim sMsg As String
               Dim iRV As Integer
               sMsg = "Normal start times for a reservation should be" & Chr(10)
               sMsg &= "08:00 AM, 01:00 PM, or 05:00 PM.  " & Chr(10)
               sMsg &= "" & Chr(10)
               sMsg &= "The time you have selected is not standard and " & Chr(10)
               sMsg &= "the return time will be computed from that time making" & Chr(10)
               sMsg &= "a non-standard return time." & Chr(10)
               sMsg &= "" & Chr(10)
               sMsg &= "Are you sure you want to reserve starting at the " & Chr(10)
               sMsg &= "selected time?" & Chr(10)
               sMsg &= "" & Chr(10)
               sMsg &= "Click Yes to continue or No to reset the time." & Chr(10)
               sMsg &= "" & Chr(10)
               iRV = MsgBox(sMsg, CType(308, Microsoft.VisualBasic.MsgBoxStyle), "Confirm Reservation Start Time")

               If iRV = 6 Then
                  ' Yes Code goes here
               Else
                  ' No code goes here
                  Me.dtpStartRes.Value = CType(Today, String) & " 08:00 AM"
                  Exit Sub
               End If
            End If

            If .cbCust.Text.Length = 0 Or _
               .cbNbrPeriods.Text.Length = 0 Or _
               .cbPeriod.Text.Length = 0 Then
               MsgBox("Please select all parameters and try again.", MsgBoxStyle.Exclamation)
               Exit Sub
            End If

            Dim period As String = .cbPeriod.Text
            Select Case period
               Case HOURLY
                  resEndDate = DateAdd(DateInterval.Hour, CType(.cbNbrPeriods.Text, Double), .dtpStartRes.Value)
               Case DAILY
                  resEndDate = DateAdd(DateInterval.Day, CType(.cbNbrPeriods.Text, Double), .dtpStartRes.Value)
               Case HALF_DAY
                  resEndDate = DateAdd(DateInterval.Hour, 5, .dtpStartRes.Value)
               Case WEEKLY
                  resEndDate = DateAdd(DateInterval.Day, 7 * CType(.cbNbrPeriods.Text, Double), .dtpStartRes.Value)
               Case MONTHLY
                  resEndDate = DateAdd(DateInterval.Month, CType(.cbNbrPeriods.Text, Double), .dtpStartRes.Value)
               Case WEEK_END
                  resEndDate = DateAdd(DateInterval.Day, CType(.cbNbrPeriods.Text, Double), .dtpStartRes.Value)
            End Select
            Dim equipClass As Integer = .dtEquip.Rows(.equipHitRow).Item("price_id")
            Dim resDate As DateTime = .dtpStartRes.Value

            equipName = .dtEquip.Rows(.equipHitRow).Item("equip_name")

            ' before we can insert the reservation, we
            ' must ensure that it can be made
            If oCE.IsReservable(equipClass, resDate, resEndDate) Then
               GoTo MakeTheReservation
            Else
               Dim sMsg As String
               Dim iRV As Integer
               sMsg = "Making this reservation conflicts with equipment" & Chr(10)
               sMsg &= "already rented or reserved.  The type of equipment" & Chr(10)
               sMsg &= "that you want to reserve is not available for the requested" & Chr(10)
               sMsg &= "time period." & Chr(10)
               sMsg &= "" & Chr(10)
               sMsg &= "Click Ok to make the reservation anyway, or click No" & Chr(10)
               sMsg &= "to cancel the reservation." & Chr(10)
               sMsg &= "" & Chr(10)
               iRV = MsgBox(sMsg, CType(33, Microsoft.VisualBasic.MsgBoxStyle), "Confirm Making Reservation")

               If iRV = 1 Then
                  ' Ok Code goes here
                  GoTo MakeTheReservation
               Else
                  ' Cancel code goes here
                  Exit Sub
               End If
            End If

MakeTheReservation:
            SQL = "insert into  reservations"
            SQL &= "(customer_name,equip_name,customerid,num_periods, "
            SQL &= "res_date,notes,res_period,equip_class,"
            SQL &= "phone,contact,res_end_date) "
            SQL &= "values("
            SQL &= "'" & .cbCust.Text & "', "
            SQL &= "'" & equipName & "', "
            SQL &= .txtCustomerID.Text & ", "
            SQL &= .cbNbrPeriods.Text & ", "
            SQL &= "#" & resDate & "#, "
            SQL &= "'" & Replace(.txtNotes.Text, "'", "''") & "', "
            SQL &= "'" & .cbPeriod.Text & "', "
            SQL &= equipClass & ", "
            SQL &= "'" & .txtPhone.Text & "', "
            SQL &= "'" & .txtContact.Text & "',"
            SQL &= "'" & resEndDate.ToString & "'"
            SQL &= ")"
            If oDA.SendActionSql(SQL, ConnectString, sErr) < 1 Then
               MsgBox("Save failed: " & Chr(10) & sErr, MsgBoxStyle.Critical)
            End If

            LoadResGrid()
            Me.dtpStartRes.Value = CType(Today, String) & " 08:00 AM"
            MsgBox("Reservation was successfully made.", MsgBoxStyle.Information)
            If closeForm Then
               Me.Close()
               System.Windows.Forms.Application.DoEvents()
            End If
         End With
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try

   End Sub
   ''' <summary>
   ''' Load the reservation data into the respective labels
   ''' so that user can see details without scrolling grid.
   ''' <summary>
   Private Sub LoadResBoxes()
      Try
         With Me.dtRes.Rows(resHitRow)
            Me.lblContactName.Text = MNS(.Item("contact"))
            Me.lblCustomerName.Text = MNS(.Item("customer_name"))
            Me.lblEquipName.Text = MNS(.Item("equip_name"))
            Me.lblNumberPeriods.Text = MNI(.Item("num_periods"))
            Me.lblReservationPeriod.Text = MNS(.Item("res_period"))
            Me.lblReservationDate.Text = Format(.Item("res_date"), "MM/dd/yyyy HH:mm tt")
            Me.lblEndDate.Text = Format(.Item("res_end_date"), "MM/dd/yyyy HH:mm tt")
            Me.lblPhoneNumber.Text = MNS(.Item("phone"))
            Me.textResNotes.Text = MNS(Me.dtRes.Rows(resHitRow).Item("notes"))
         End With
      Catch
      End Try
   End Sub

   ''' <summary>
   '''  Load the reservations grid from table reservations
   ''' </summary>
   Private Sub LoadResGrid()

      Try
         SQL = "select  equip_name,  Num_Periods, "
         SQL &= "Res_Period,Res_Date,Res_End_Date,Contact,"
         SQL &= "customer_name,Phone,CustomerID,Equip_Class, "
         SQL &= "reservationid,notes,rerent_company "
         SQL &= "from reservations "
         SQL &= "order by equip_name"
         oCG.ClearDataTableForRebinding(dtRes)
         oDA.SendQuery(SQL, dtRes, ConnectString, "dt")
         Dim formats() As String = _
                  {"", "150", "T", "L", _
                  "", "60", "T", "R", _
                  "", "60", "T", "L", _
                  "MM/dd/yyyy hh:mm tt", "120", "T", "L", _
                  "MM/dd/yyyy hh:mm tt", "120", "T", "L", _
                  "", "100", "T", "L", _
                  "", "150", "T", "L", _
                  "", "60", "T", "L", _
                  "", "60", "T", "R", _
                  "", "60", "T", "R", _
                  "", "60", "T", "R", _
                  "", "100", "T", "R", _
                  "", "100", "T", "L", _
                  "", "100", "T", "L"}
         oCG.SetTablesStyle("Select", dtRes, Me.dgReservations, formats)
         oCG.BindDataTableToGrid(dtRes, dgReservations)
         oCG.DisableAddNew(dgReservations, Me)
         Me.resHitRow = 0
         Try
            oCG.SelectCkBoxRow(dgReservations, resHitRow)
            LoadResBoxes()
         Catch
         End Try

      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub


   ''' <summary>
   ''' Load the equipment class grid from the rental rates table.
   ''' </summary>
   Private Sub LoadEquipGrid()
      Dim i As Integer


      Try
         SQL = "Select Equip_Name,Price_Id from rental_rates "
         SQL &= "order by Equip_name"
         oCG.ClearDataTableForRebinding(dtEquip)

         If oDA.SendQuery(SQL, dtEquip, ConnectString, "dt") > 0 Then
            Dim formats() As String = _
               {"", "200", "T", "L", _
               "", "60", "T", "R"}
            oCG.SetTablesStyle("Reserve", dtEquip, Me.dgEquip, formats)
            oCG.BindDataTableToGrid(dtEquip, dgEquip)
            oCG.DisableAddNew(dgEquip, Me)
         End If
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub
   Private Sub LoadCustomerCombo()
      Dim dt As New DataTable()
      Dim i As Integer
      SQL = "select companyname from customers "
      SQL &= "order by companyname"
      oDA.SendQuery(SQL, dt, ConnectString)
      Me.cbCust.Items.Clear()

      For i = 0 To dt.Rows.Count - 1
         With dt.Rows(i)
            Me.cbCust.Items.Add(.Item("companyname"))
         End With
      Next
   End Sub
   ''' <summary>
   ''' Load the customer id
   ''' </summary>
   Private Sub LoadCustomerBoxes()
      Dim dt As New DataTable()

      SQL = ""
      SQL &= "select customerid,companyname,phonenumber,contactname "
      SQL &= "from customers "
      SQL &= "where companyname = '" & Me.cbCust.Text & "'"
      If oDA.SendQuery(SQL, dt, ConnectString) > 0 Then
         Me.txtCustomerID.Text = MNI(dt.Rows(0).Item("customerid"))
         Me.txtPhone.Text = MNS(dt.Rows(0).Item("phonenumber"))
         Me.txtContact.Text = MNS(dt.Rows(0).Item("contactname"))
      End If
   End Sub
   Private Sub LoadCustomerBoxes(ByVal CustId As Integer)
      Dim SQL As String
      Dim dt As New DataTable()

      SQL = ""
      SQL = SQL & "select customerid,companyname, phonenumber, contactname "
      SQL = SQL & "from customers "
      SQL = SQL & "where customerid = " & CustId & " "
      oDA.SendQuery(SQL, dt, ConnectString)

      If dt.Rows.Count > 0 Then
         With dt.Rows(0)
            Me.cbCust.Text = MNS(.Item("companyname"))
            Me.txtCustomerID.Text = MNS(.Item("customerid"))
            Me.txtContact.Text = String.Empty 'MNS(.Item("contactname"))
            Me.txtPhone.Text = String.Empty 'MNS(.Item("phonenumber"))
         End With
      End If
   End Sub
#End Region

#Region "Form & Control Events"
   Private Sub btnAddCustomer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddCustomer.Click
      Dim oFrm As New frmCustomerMaintenance()
      oFrm.ShowDialog()
      LoadCustomerCombo()
   End Sub
   Private Sub btnAddnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddnClose.Click
      AddEquipment(True)
   End Sub

   ''' <summary>
   '''
   ''' </summary>
   ''' <param name = "sender"></param>
   ''' <param name = "e"></param>
   Private Sub cmdClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdClose.Click, btnClose2.Click

      Me.Close()
      System.Windows.Forms.Application.DoEvents()
   End Sub

   ''' <summary>
   ''' Delete the selected row in the reservations grid.
   ''' </summary>
   ''' <param name = "sender"></param>
   ''' <param name = "e"></param>
   Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
      Dim dt As New DataTable()
      Dim oDA As New CDataAccess()
      Dim iRows As Integer

      Try
         Dim sErr As String = ""
         If dtRes.Rows(resHitRow).Item("Select") = "false" Then
            MsgBox("You have no row checked to delete.", MsgBoxStyle.Exclamation)
            Exit Sub
         End If
         If MsgBox("Are you sure you want to remove the selected reservation?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
            Exit Sub
         End If
         SQL = "delete from reservations "
         SQL &= "where reservationid= " & Me.dtRes.Rows(Me.resHitRow).Item("reservationid") & " "
         iRows = oDA.SendActionSql(SQL, ConnectString, sErr)
         If iRows = 0 Then
            MsgBox("Delete of reservation item failed.  " & Chr(10) & sErr, MsgBoxStyle.Critical)
            Exit Sub
         End If
         Me.LoadResGrid()
         Me.equipHitRow = 0
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

   ''' <summary>
   ''' Call the method for printing the contents of the reservations grid
   ''' </summary>
   ''' <param name = "sender"></param>
   ''' <param name = "e"></param>
   Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
      MsgBox("Unimplemented")
   End Sub


   ''' <summary>
   ''' Load the customer text boxes
   ''' </summary>
   ''' <param name = "sender"></param>
   ''' <param name = "e"></param>
   Private Sub cbCust_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbCust.SelectedIndexChanged
      If Noise Then Exit Sub
      LoadCustomerBoxes()
      Me.chkCashCustomer.Checked = False
   End Sub

   ''' <summary>
   ''' Form load will load the reservations grid and the equip class grid
   ''' </summary>
   ''' <param name = "sender"></param>
   ''' <param name = "e"></param>
   Private Sub frmReservations_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
      LoadResGrid()
      LoadEquipGrid()
      LoadCustomerCombo()
      With Me
         Me.dtpStartRes.CustomFormat = "MM/dd/yyyy hh:mm tt"
         Me.dtpStartRes.Value = CType(Today, String) & " 08:00 AM"

      End With
   End Sub

   Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
      AddEquipment(False)
   End Sub


   Private Sub dgEquip_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgEquip.MouseUp
      Dim b As Boolean
      Try
         equipHitRow = oCG.GetClickedRow(e, dgEquip)
         If dgEquip.DataSource.rows(equipHitRow).item("Reserve") = "true" Then
            dgEquip.DataSource.rows(equipHitRow).item("Reserve") = "false"
            Exit Sub
         End If
         oCG.UncheckAllBoxes(dtEquip, "Reserve")
         equipHitRow = oCG.SelectCkBoxRow(dtEquip, dgEquip, e, "Reserve", b)
      Catch ex As System.Exception
      End Try

   End Sub

   Private Sub dgReservations_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgReservations.MouseUp
      Dim b As Boolean
      Try
         Dim i As Integer = oCG.GetClickedRow(e, dgReservations)
         If dgReservations.DataSource.rows(i).item("Select") = "true" Then
            Me.dgReservations.DataSource.rows(i).item("Select") = "false"
            Exit Sub
         End If
         oCG.UncheckAllBoxes(dtRes, "Select")
         resHitRow = oCG.SelectCkBoxRow(dtRes, dgReservations, e, "Select", b)
         LoadResBoxes()
      Catch ex As System.Exception
      End Try
   End Sub


   Private Sub chkCashCustomer_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCashCustomer.CheckedChanged
      'Static Busy As Boolean
      'If Busy Then Exit Sub
      'Busy = True
      Noise = True
      If chkCashCustomer.Checked Then
         LoadCustomerBoxes(0)
      Else
         Me.cbCust.Text = String.Empty
         Me.txtCustomerID.Text = String.Empty
      End If
      'Busy = False
      Noise = False
      Me.btnAddCustomer.Enabled = Not Me.chkCashCustomer.Checked
   End Sub
#End Region


End Class
