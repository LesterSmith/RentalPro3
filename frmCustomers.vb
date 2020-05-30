'****************************************
'* Purpose:
'*
'* Author:  Les Smith
'* Date Created: 05/15/2003 at 09:09:04
'* CopyRight:  HHI Software
'****************************************
'*
Option Strict Off
Option Explicit On
Imports System.Windows.Forms.Application
Imports System.Runtime.InteropServices

Public Class frmCustomers
    Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
    Public Sub New()
        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()
        oDA = New CDataAccess()
    End Sub
    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
        If Disposing Then
            If Not components Is Nothing Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(Disposing)
    End Sub
    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents _Label1_0 As System.Windows.Forms.Label
    Public WithEvents _Label1_1 As System.Windows.Forms.Label
    Public WithEvents _Label1_2 As System.Windows.Forms.Label
    Public WithEvents _lblFieldLable_0 As System.Windows.Forms.Label
    Public WithEvents _lblLabels_0 As System.Windows.Forms.Label
    Public WithEvents _lblLabels_1 As System.Windows.Forms.Label
    Public WithEvents _lblLabels_4 As System.Windows.Forms.Label
    Public WithEvents _lblLabels_5 As System.Windows.Forms.Label
    Public WithEvents _lblLabels_6 As System.Windows.Forms.Label
    Public WithEvents btnCheckout As System.Windows.Forms.Button
    Friend WithEvents btnManualRecalc As System.Windows.Forms.Button
    Friend WithEvents btnOtherCharges As System.Windows.Forms.Button
    Friend WithEvents btnPrintInvoice As System.Windows.Forms.Button
    Friend WithEvents cbEmployees As System.Windows.Forms.ComboBox
    Friend WithEvents cbExpMon As System.Windows.Forms.ComboBox
    Friend WithEvents cbExpYr As System.Windows.Forms.ComboBox
    Public WithEvents cboDeliveryDistance As System.Windows.Forms.ComboBox
    Friend WithEvents chkCashCustomer As System.Windows.Forms.CheckBox
    Public WithEvents chkChargeTax As System.Windows.Forms.CheckBox
    Public WithEvents chkDeliveryRequired As System.Windows.Forms.CheckBox
    Public WithEvents chkDepositRequired As System.Windows.Forms.RadioButton
    Public WithEvents chkManualDepositOverride As System.Windows.Forms.RadioButton
    Public WithEvents cmdAdd As System.Windows.Forms.Button
    Public WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents dbgShoppingList As System.Windows.Forms.DataGrid
    Friend WithEvents dtpCkOutDateReset As System.Windows.Forms.DateTimePicker
    Public WithEvents fraDelivery As System.Windows.Forms.GroupBox
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents HelpProvider1 As System.Windows.Forms.HelpProvider
    Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Public WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents lblAmtPaid As System.Windows.Forms.Label
    Friend WithEvents lblBalDue As System.Windows.Forms.Label
    Public WithEvents lblCustId As System.Windows.Forms.Label
    Friend WithEvents lblDelivery As System.Windows.Forms.Label
    Friend WithEvents lblDeposit As System.Windows.Forms.Label
    Public WithEvents lblFieldLable As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Friend WithEvents lblItemTotal As System.Windows.Forms.Label
    Public WithEvents lblLabels As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Friend WithEvents lblLine As System.Windows.Forms.Label
    Friend WithEvents lblLine2 As System.Windows.Forms.Label
    Friend WithEvents lblNotes As System.Windows.Forms.Label
    Friend WithEvents lblTax As System.Windows.Forms.Label
    Friend WithEvents lblTotal As System.Windows.Forms.Label
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddCustomer As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDeleteSelectedRow As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExit As System.Windows.Forms.MenuItem
    Friend WithEvents mnuLoadCustomerMaintForm As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMain As System.Windows.Forms.MainMenu
    Friend WithEvents mnuPrint As System.Windows.Forms.MenuItem
    Friend WithEvents optBillTo As System.Windows.Forms.RadioButton
    Friend WithEvents optCash As System.Windows.Forms.RadioButton
    Friend WithEvents optLeftBlankCheck As System.Windows.Forms.RadioButton
    Friend WithEvents optLeftCardNumber As System.Windows.Forms.RadioButton
    Friend WithEvents optPaidByCheck As System.Windows.Forms.RadioButton
    Friend WithEvents optPaidByCreditCard As System.Windows.Forms.RadioButton
    Friend WithEvents SSOleDBCombo1 As System.Windows.Forms.ComboBox
    Public WithEvents txtManualEntryDelPrice As System.Windows.Forms.TextBox
    Friend WithEvents txtAmtPaid As System.Windows.Forms.TextBox
    Friend WithEvents txtBalDue As System.Windows.Forms.TextBox
    Public WithEvents txtBillingAddress1 As System.Windows.Forms.TextBox
    Public WithEvents txtBillingAddress2 As System.Windows.Forms.TextBox
    Friend WithEvents txtCardId As System.Windows.Forms.TextBox
    Friend WithEvents txtCheckNumber As System.Windows.Forms.TextBox
    Public WithEvents txtCity As System.Windows.Forms.TextBox
    Public WithEvents txtCompanyName As System.Windows.Forms.TextBox
    Public WithEvents txtCustomerID As System.Windows.Forms.TextBox
    Friend WithEvents txtDelivery As System.Windows.Forms.TextBox
    Friend WithEvents txtDeposit As System.Windows.Forms.TextBox
    Friend WithEvents txtItemTotal As System.Windows.Forms.TextBox
    Public WithEvents txtManualDeposit As System.Windows.Forms.TextBox
    Friend WithEvents txtNotes As System.Windows.Forms.TextBox
    Public WithEvents txtPostalCode As System.Windows.Forms.TextBox
    Friend WithEvents txtSalesTax As System.Windows.Forms.TextBox
    Public WithEvents txtState As System.Windows.Forms.TextBox
    Friend WithEvents chkPrintToFile As System.Windows.Forms.CheckBox
    Friend WithEvents btnShowPrinters As System.Windows.Forms.Button
    Friend WithEvents btnHelp As System.Windows.Forms.Button
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Public WithEvents txtPhone As System.Windows.Forms.TextBox
    Friend WithEvents chkCustomerPickup As System.Windows.Forms.CheckBox
    Friend WithEvents textDriversLicence As System.Windows.Forms.TextBox
    Friend WithEvents chkCkOutAndIN As System.Windows.Forms.CheckBox
    Friend WithEvents txtTaxID As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents chkShipBillSame As System.Windows.Forms.CheckBox
    Friend WithEvents txtPONbr As System.Windows.Forms.TextBox
    Friend WithEvents lblPONbr As System.Windows.Forms.Label
    Public WithEvents txtShipZip As System.Windows.Forms.TextBox
    Public WithEvents txtShipState As System.Windows.Forms.TextBox
    Public WithEvents txtShipCity As System.Windows.Forms.TextBox
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents txtContactName As System.Windows.Forms.TextBox
    Public WithEvents txtShipAddress1 As System.Windows.Forms.TextBox
    Public WithEvents txtShipToCustomer As System.Windows.Forms.TextBox
    Public WithEvents lblContact As System.Windows.Forms.Label
    Public WithEvents _lblLabels_3 As System.Windows.Forms.Label
    Public WithEvents _lblLabels_2 As System.Windows.Forms.Label
    Friend WithEvents chkSaveCreditCard As System.Windows.Forms.CheckBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents txtTotal As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCustomers))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me._Label1_2 = New System.Windows.Forms.Label()
        Me.cmdAdd = New System.Windows.Forms.Button()
        Me.chkCashCustomer = New System.Windows.Forms.CheckBox()
        Me.chkCkOutAndIN = New System.Windows.Forms.CheckBox()
        Me.dbgShoppingList = New System.Windows.Forms.DataGrid()
        Me.fraDelivery = New System.Windows.Forms.GroupBox()
        Me.chkChargeTax = New System.Windows.Forms.CheckBox()
        Me.chkManualDepositOverride = New System.Windows.Forms.RadioButton()
        Me.chkDepositRequired = New System.Windows.Forms.RadioButton()
        Me.txtManualDeposit = New System.Windows.Forms.TextBox()
        Me.txtManualEntryDelPrice = New System.Windows.Forms.TextBox()
        Me.cboDeliveryDistance = New System.Windows.Forms.ComboBox()
        Me.chkDeliveryRequired = New System.Windows.Forms.CheckBox()
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.txtPhone = New System.Windows.Forms.TextBox()
        Me.chkCustomerPickup = New System.Windows.Forms.CheckBox()
        Me.txtBillingAddress2 = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.textDriversLicence = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.txtTaxID = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.chkShipBillSame = New System.Windows.Forms.CheckBox()
        Me.lblLine = New System.Windows.Forms.Label()
        Me.txtPONbr = New System.Windows.Forms.TextBox()
        Me.lblPONbr = New System.Windows.Forms.Label()
        Me.txtShipZip = New System.Windows.Forms.TextBox()
        Me.txtShipState = New System.Windows.Forms.TextBox()
        Me.txtShipCity = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtContactName = New System.Windows.Forms.TextBox()
        Me.txtPostalCode = New System.Windows.Forms.TextBox()
        Me.txtState = New System.Windows.Forms.TextBox()
        Me.txtShipAddress1 = New System.Windows.Forms.TextBox()
        Me.txtCity = New System.Windows.Forms.TextBox()
        Me.txtShipToCustomer = New System.Windows.Forms.TextBox()
        Me.txtBillingAddress1 = New System.Windows.Forms.TextBox()
        Me.txtCompanyName = New System.Windows.Forms.TextBox()
        Me.lblContact = New System.Windows.Forms.Label()
        Me._lblLabels_6 = New System.Windows.Forms.Label()
        Me._lblLabels_5 = New System.Windows.Forms.Label()
        Me._lblLabels_4 = New System.Windows.Forms.Label()
        Me._lblLabels_3 = New System.Windows.Forms.Label()
        Me._lblLabels_2 = New System.Windows.Forms.Label()
        Me._lblLabels_1 = New System.Windows.Forms.Label()
        Me._lblLabels_0 = New System.Windows.Forms.Label()
        Me.lblCustId = New System.Windows.Forms.Label()
        Me.txtCustomerID = New System.Windows.Forms.TextBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.btnHelp = New System.Windows.Forms.Button()
        Me.SSOleDBCombo1 = New System.Windows.Forms.ComboBox()
        Me._lblFieldLable_0 = New System.Windows.Forms.Label()
        Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblFieldLable = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblLabels = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.txtDelivery = New System.Windows.Forms.TextBox()
        Me.lblDelivery = New System.Windows.Forms.Label()
        Me.lblTax = New System.Windows.Forms.Label()
        Me.txtSalesTax = New System.Windows.Forms.TextBox()
        Me.txtDeposit = New System.Windows.Forms.TextBox()
        Me.lblDeposit = New System.Windows.Forms.Label()
        Me.lblTotal = New System.Windows.Forms.Label()
        Me.txtTotal = New System.Windows.Forms.TextBox()
        Me.lblItemTotal = New System.Windows.Forms.Label()
        Me.txtItemTotal = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.optCash = New System.Windows.Forms.RadioButton()
        Me.optLeftCardNumber = New System.Windows.Forms.RadioButton()
        Me.optBillTo = New System.Windows.Forms.RadioButton()
        Me.optLeftBlankCheck = New System.Windows.Forms.RadioButton()
        Me.optPaidByCreditCard = New System.Windows.Forms.RadioButton()
        Me.optPaidByCheck = New System.Windows.Forms.RadioButton()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtCheckNumber = New System.Windows.Forms.TextBox()
        Me.lblAmtPaid = New System.Windows.Forms.Label()
        Me.txtAmtPaid = New System.Windows.Forms.TextBox()
        Me.txtBalDue = New System.Windows.Forms.TextBox()
        Me.lblBalDue = New System.Windows.Forms.Label()
        Me.lblLine2 = New System.Windows.Forms.Label()
        Me.HelpProvider1 = New System.Windows.Forms.HelpProvider()
        Me.btnCheckout = New System.Windows.Forms.Button()
        Me.btnManualRecalc = New System.Windows.Forms.Button()
        Me.btnPrintInvoice = New System.Windows.Forms.Button()
        Me.mnuMain = New System.Windows.Forms.MainMenu(Me.components)
        Me.MenuItem1 = New System.Windows.Forms.MenuItem()
        Me.mnuPrint = New System.Windows.Forms.MenuItem()
        Me.MenuItem2 = New System.Windows.Forms.MenuItem()
        Me.mnuLoadCustomerMaintForm = New System.Windows.Forms.MenuItem()
        Me.mnuDeleteSelectedRow = New System.Windows.Forms.MenuItem()
        Me.mnuExit = New System.Windows.Forms.MenuItem()
        Me.MenuItem3 = New System.Windows.Forms.MenuItem()
        Me.mnuAddCustomer = New System.Windows.Forms.MenuItem()
        Me.btnOtherCharges = New System.Windows.Forms.Button()
        Me.lblNotes = New System.Windows.Forms.Label()
        Me.txtNotes = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.cbExpMon = New System.Windows.Forms.ComboBox()
        Me.cbExpYr = New System.Windows.Forms.ComboBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txtCardId = New System.Windows.Forms.TextBox()
        Me.cbEmployees = New System.Windows.Forms.ComboBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.dtpCkOutDateReset = New System.Windows.Forms.DateTimePicker()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.chkPrintToFile = New System.Windows.Forms.CheckBox()
        Me.btnShowPrinters = New System.Windows.Forms.Button()
        Me.chkSaveCreditCard = New System.Windows.Forms.CheckBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        CType(Me.dbgShoppingList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraDelivery.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me.Frame1.SuspendLayout()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblFieldLable, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblLabels, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        '_Label1_2
        '
        Me._Label1_2.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_2, CType(2, Short))
        Me._Label1_2.Location = New System.Drawing.Point(400, 42)
        Me._Label1_2.Name = "_Label1_2"
        Me._Label1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_2.Size = New System.Drawing.Size(82, 28)
        Me._Label1_2.TabIndex = 30
        Me._Label1_2.Text = "Manual Deposit Override"
        Me.ToolTip1.SetToolTip(Me._Label1_2, "Enter amount to override deposit, can be zero")
        '
        'cmdAdd
        '
        Me.cmdAdd.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAdd.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAdd.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAdd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.HelpProvider1.SetHelpString(Me.cmdAdd, "Click to display customer  screen and click Add on it to add customer.")
        Me.cmdAdd.Location = New System.Drawing.Point(382, 12)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.HelpProvider1.SetShowHelp(Me.cmdAdd, True)
        Me.cmdAdd.Size = New System.Drawing.Size(45, 25)
        Me.cmdAdd.TabIndex = 1
        Me.cmdAdd.Text = "&Add"
        Me.ToolTip1.SetToolTip(Me.cmdAdd, "Add new customer")
        Me.cmdAdd.UseVisualStyleBackColor = False
        '
        'chkCashCustomer
        '
        Me.chkCashCustomer.ForeColor = System.Drawing.Color.Red
        Me.chkCashCustomer.Location = New System.Drawing.Point(446, 16)
        Me.chkCashCustomer.Name = "chkCashCustomer"
        Me.chkCashCustomer.Size = New System.Drawing.Size(104, 15)
        Me.chkCashCustomer.TabIndex = 16
        Me.chkCashCustomer.Text = "Cash Customer"
        Me.ToolTip1.SetToolTip(Me.chkCashCustomer, "Use Cash Customer Record , customer will not enter database")
        Me.chkCashCustomer.Visible = False
        '
        'chkCkOutAndIN
        '
        Me.HelpProvider1.SetHelpString(Me.chkCkOutAndIN, "Quick Check out and check back in to facilitate late check in of equipment at a h" & _
                "alf day rate")
        Me.chkCkOutAndIN.Location = New System.Drawing.Point(528, 122)
        Me.chkCkOutAndIN.Name = "chkCkOutAndIN"
        Me.HelpProvider1.SetShowHelp(Me.chkCkOutAndIN, True)
        Me.chkCkOutAndIN.Size = New System.Drawing.Size(44, 16)
        Me.chkCkOutAndIN.TabIndex = 57
        Me.chkCkOutAndIN.Text = "CkOut && In At Same Time"
        Me.ToolTip1.SetToolTip(Me.chkCkOutAndIN, "Quick Check out and check back in to facilitate late check in of equipment at a h" & _
                "alf day rate")
        Me.chkCkOutAndIN.Visible = False
        '
        'dbgShoppingList
        '
        Me.dbgShoppingList.AllowSorting = False
        Me.dbgShoppingList.CaptionVisible = False
        Me.dbgShoppingList.DataMember = ""
        Me.dbgShoppingList.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dbgShoppingList.Location = New System.Drawing.Point(5, 307)
        Me.dbgShoppingList.Name = "dbgShoppingList"
        Me.dbgShoppingList.Size = New System.Drawing.Size(586, 109)
        Me.dbgShoppingList.TabIndex = 44
        '
        'fraDelivery
        '
        Me.fraDelivery.BackColor = System.Drawing.SystemColors.Control
        Me.fraDelivery.Controls.Add(Me.chkChargeTax)
        Me.fraDelivery.Controls.Add(Me.chkManualDepositOverride)
        Me.fraDelivery.Controls.Add(Me.chkDepositRequired)
        Me.fraDelivery.Controls.Add(Me.txtManualDeposit)
        Me.fraDelivery.Controls.Add(Me.txtManualEntryDelPrice)
        Me.fraDelivery.Controls.Add(Me.cboDeliveryDistance)
        Me.fraDelivery.Controls.Add(Me.chkDeliveryRequired)
        Me.fraDelivery.Controls.Add(Me._Label1_2)
        Me.fraDelivery.Controls.Add(Me._Label1_1)
        Me.fraDelivery.Controls.Add(Me._Label1_0)
        Me.fraDelivery.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraDelivery.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraDelivery.Location = New System.Drawing.Point(6, 225)
        Me.fraDelivery.Name = "fraDelivery"
        Me.fraDelivery.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraDelivery.Size = New System.Drawing.Size(587, 76)
        Me.fraDelivery.TabIndex = 26
        Me.fraDelivery.TabStop = False
        Me.fraDelivery.Text = "Delivery/Deposit Information"
        '
        'chkChargeTax
        '
        Me.chkChargeTax.BackColor = System.Drawing.SystemColors.Control
        Me.chkChargeTax.Checked = True
        Me.chkChargeTax.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkChargeTax.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkChargeTax.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkChargeTax.ForeColor = System.Drawing.SystemColors.ControlText
        Me.HelpProvider1.SetHelpString(Me.chkChargeTax, "Uncheck to remove Salse Tax.")
        Me.chkChargeTax.Location = New System.Drawing.Point(314, 52)
        Me.chkChargeTax.Name = "chkChargeTax"
        Me.chkChargeTax.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.HelpProvider1.SetShowHelp(Me.chkChargeTax, True)
        Me.chkChargeTax.Size = New System.Drawing.Size(75, 17)
        Me.chkChargeTax.TabIndex = 34
        Me.chkChargeTax.Text = "&Sales Tax"
        Me.chkChargeTax.UseVisualStyleBackColor = False
        '
        'chkManualDepositOverride
        '
        Me.chkManualDepositOverride.BackColor = System.Drawing.SystemColors.Control
        Me.chkManualDepositOverride.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkManualDepositOverride.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkManualDepositOverride.ForeColor = System.Drawing.SystemColors.ControlText
        Me.HelpProvider1.SetHelpString(Me.chkManualDepositOverride, "Use manual override deposit charge.")
        Me.chkManualDepositOverride.Location = New System.Drawing.Point(9, 52)
        Me.chkManualDepositOverride.Name = "chkManualDepositOverride"
        Me.chkManualDepositOverride.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.HelpProvider1.SetShowHelp(Me.chkManualDepositOverride, True)
        Me.chkManualDepositOverride.Size = New System.Drawing.Size(122, 18)
        Me.chkManualDepositOverride.TabIndex = 33
        Me.chkManualDepositOverride.TabStop = True
        Me.chkManualDepositOverride.Text = "Use &Manual Deposit Override"
        Me.chkManualDepositOverride.UseVisualStyleBackColor = False
        '
        'chkDepositRequired
        '
        Me.chkDepositRequired.BackColor = System.Drawing.SystemColors.Control
        Me.chkDepositRequired.Checked = True
        Me.chkDepositRequired.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkDepositRequired.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkDepositRequired.ForeColor = System.Drawing.SystemColors.ControlText
        Me.HelpProvider1.SetHelpString(Me.chkDepositRequired, "Use total of default deposit for all items.")
        Me.chkDepositRequired.Location = New System.Drawing.Point(161, 52)
        Me.chkDepositRequired.Name = "chkDepositRequired"
        Me.chkDepositRequired.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.HelpProvider1.SetShowHelp(Me.chkDepositRequired, True)
        Me.chkDepositRequired.Size = New System.Drawing.Size(134, 17)
        Me.chkDepositRequired.TabIndex = 32
        Me.chkDepositRequired.TabStop = True
        Me.chkDepositRequired.Text = "Auto De&posit Required"
        Me.chkDepositRequired.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.chkDepositRequired.UseVisualStyleBackColor = False
        '
        'txtManualDeposit
        '
        Me.txtManualDeposit.AcceptsReturn = True
        Me.txtManualDeposit.BackColor = System.Drawing.SystemColors.Window
        Me.txtManualDeposit.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtManualDeposit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtManualDeposit.ForeColor = System.Drawing.SystemColors.WindowText
        Me.HelpProvider1.SetHelpString(Me.txtManualDeposit, "Enter manual deposit override")
        Me.txtManualDeposit.Location = New System.Drawing.Point(487, 44)
        Me.txtManualDeposit.MaxLength = 0
        Me.txtManualDeposit.Name = "txtManualDeposit"
        Me.txtManualDeposit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.HelpProvider1.SetShowHelp(Me.txtManualDeposit, True)
        Me.txtManualDeposit.Size = New System.Drawing.Size(88, 20)
        Me.txtManualDeposit.TabIndex = 29
        Me.txtManualDeposit.Tag = "$#,##0.00;($#,##0.00)"
        Me.txtManualDeposit.Text = "$0.00"
        Me.txtManualDeposit.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtManualEntryDelPrice
        '
        Me.txtManualEntryDelPrice.AcceptsReturn = True
        Me.txtManualEntryDelPrice.BackColor = System.Drawing.SystemColors.Window
        Me.txtManualEntryDelPrice.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtManualEntryDelPrice.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtManualEntryDelPrice.ForeColor = System.Drawing.SystemColors.WindowText
        Me.HelpProvider1.SetHelpString(Me.txtManualEntryDelPrice, "Enter override delivery charge.")
        Me.txtManualEntryDelPrice.Location = New System.Drawing.Point(487, 17)
        Me.txtManualEntryDelPrice.MaxLength = 0
        Me.txtManualEntryDelPrice.Name = "txtManualEntryDelPrice"
        Me.txtManualEntryDelPrice.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.HelpProvider1.SetShowHelp(Me.txtManualEntryDelPrice, True)
        Me.txtManualEntryDelPrice.Size = New System.Drawing.Size(88, 20)
        Me.txtManualEntryDelPrice.TabIndex = 11
        Me.txtManualEntryDelPrice.Tag = "$#,##0.00;($#,##0.00)"
        Me.txtManualEntryDelPrice.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'cboDeliveryDistance
        '
        Me.cboDeliveryDistance.BackColor = System.Drawing.SystemColors.Window
        Me.cboDeliveryDistance.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboDeliveryDistance.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboDeliveryDistance.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboDeliveryDistance.Location = New System.Drawing.Point(145, 17)
        Me.cboDeliveryDistance.Name = "cboDeliveryDistance"
        Me.cboDeliveryDistance.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboDeliveryDistance.Size = New System.Drawing.Size(248, 22)
        Me.cboDeliveryDistance.TabIndex = 10
        '
        'chkDeliveryRequired
        '
        Me.chkDeliveryRequired.BackColor = System.Drawing.SystemColors.Control
        Me.chkDeliveryRequired.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkDeliveryRequired.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkDeliveryRequired.ForeColor = System.Drawing.SystemColors.ControlText
        Me.HelpProvider1.SetHelpString(Me.chkDeliveryRequired, "Check to charge delivery price.  Uncheck for no delivery charge.")
        Me.chkDeliveryRequired.Location = New System.Drawing.Point(9, 18)
        Me.chkDeliveryRequired.Name = "chkDeliveryRequired"
        Me.chkDeliveryRequired.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.HelpProvider1.SetShowHelp(Me.chkDeliveryRequired, True)
        Me.chkDeliveryRequired.Size = New System.Drawing.Size(73, 32)
        Me.chkDeliveryRequired.TabIndex = 9
        Me.chkDeliveryRequired.Text = "Delivery &Required"
        Me.chkDeliveryRequired.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.chkDeliveryRequired.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.chkDeliveryRequired.UseVisualStyleBackColor = False
        Me.chkDeliveryRequired.Visible = False
        '
        '_Label1_1
        '
        Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_1, CType(1, Short))
        Me._Label1_1.Location = New System.Drawing.Point(400, 13)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.Size = New System.Drawing.Size(73, 28)
        Me._Label1_1.TabIndex = 28
        Me._Label1_1.Text = "Manual Entry Delivery Price"
        '
        '_Label1_0
        '
        Me._Label1_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_0, CType(0, Short))
        Me._Label1_0.Location = New System.Drawing.Point(91, 20)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_0.Size = New System.Drawing.Size(49, 14)
        Me._Label1_0.TabIndex = 27
        Me._Label1_0.Text = "Distance"
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.HelpProvider1.SetHelpString(Me.cmdCancel, "Cancel the checkout process.")
        Me.cmdCancel.Location = New System.Drawing.Point(297, 617)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.HelpProvider1.SetShowHelp(Me.cmdCancel, True)
        Me.cmdCancel.Size = New System.Drawing.Size(131, 22)
        Me.cmdCancel.TabIndex = 130
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.Label13)
        Me.Frame2.Controls.Add(Me.txtPhone)
        Me.Frame2.Controls.Add(Me.chkCustomerPickup)
        Me.Frame2.Controls.Add(Me.txtBillingAddress2)
        Me.Frame2.Controls.Add(Me.Label11)
        Me.Frame2.Controls.Add(Me.textDriversLicence)
        Me.Frame2.Controls.Add(Me.Label15)
        Me.Frame2.Controls.Add(Me.chkCkOutAndIN)
        Me.Frame2.Controls.Add(Me.txtTaxID)
        Me.Frame2.Controls.Add(Me.Label6)
        Me.Frame2.Controls.Add(Me.chkShipBillSame)
        Me.Frame2.Controls.Add(Me.lblLine)
        Me.Frame2.Controls.Add(Me.txtPONbr)
        Me.Frame2.Controls.Add(Me.lblPONbr)
        Me.Frame2.Controls.Add(Me.txtShipZip)
        Me.Frame2.Controls.Add(Me.txtShipState)
        Me.Frame2.Controls.Add(Me.txtShipCity)
        Me.Frame2.Controls.Add(Me.Label3)
        Me.Frame2.Controls.Add(Me.Label4)
        Me.Frame2.Controls.Add(Me.Label5)
        Me.Frame2.Controls.Add(Me.txtContactName)
        Me.Frame2.Controls.Add(Me.txtPostalCode)
        Me.Frame2.Controls.Add(Me.txtState)
        Me.Frame2.Controls.Add(Me.txtShipAddress1)
        Me.Frame2.Controls.Add(Me.txtCity)
        Me.Frame2.Controls.Add(Me.txtShipToCustomer)
        Me.Frame2.Controls.Add(Me.txtBillingAddress1)
        Me.Frame2.Controls.Add(Me.txtCompanyName)
        Me.Frame2.Controls.Add(Me.lblContact)
        Me.Frame2.Controls.Add(Me._lblLabels_6)
        Me.Frame2.Controls.Add(Me._lblLabels_5)
        Me.Frame2.Controls.Add(Me._lblLabels_4)
        Me.Frame2.Controls.Add(Me._lblLabels_3)
        Me.Frame2.Controls.Add(Me._lblLabels_2)
        Me.Frame2.Controls.Add(Me._lblLabels_1)
        Me.Frame2.Controls.Add(Me._lblLabels_0)
        Me.Frame2.Controls.Add(Me.lblCustId)
        Me.Frame2.Controls.Add(Me.txtCustomerID)
        Me.Frame2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(5, 55)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(589, 164)
        Me.Frame2.TabIndex = 16
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Customer"
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(208, 120)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(37, 16)
        Me.Label13.TabIndex = 69
        Me.Label13.Text = "Phone"
        '
        'txtPhone
        '
        Me.txtPhone.AcceptsReturn = True
        Me.txtPhone.BackColor = System.Drawing.SystemColors.Window
        Me.txtPhone.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPhone.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPhone.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPhone.Location = New System.Drawing.Point(248, 120)
        Me.txtPhone.MaxLength = 0
        Me.txtPhone.Name = "txtPhone"
        Me.txtPhone.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPhone.Size = New System.Drawing.Size(104, 20)
        Me.txtPhone.TabIndex = 13
        Me.txtPhone.Tag = "(No Auto Formatting)"
        '
        'chkCustomerPickup
        '
        Me.chkCustomerPickup.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkCustomerPickup.Location = New System.Drawing.Point(160, 144)
        Me.chkCustomerPickup.Name = "chkCustomerPickup"
        Me.chkCustomerPickup.Size = New System.Drawing.Size(112, 16)
        Me.chkCustomerPickup.TabIndex = 67
        Me.chkCustomerPickup.Text = "Customer Pickup"
        '
        'txtBillingAddress2
        '
        Me.txtBillingAddress2.AcceptsReturn = True
        Me.txtBillingAddress2.BackColor = System.Drawing.SystemColors.Window
        Me.txtBillingAddress2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtBillingAddress2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBillingAddress2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtBillingAddress2.Location = New System.Drawing.Point(90, 50)
        Me.txtBillingAddress2.MaxLength = 0
        Me.txtBillingAddress2.Name = "txtBillingAddress2"
        Me.txtBillingAddress2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtBillingAddress2.Size = New System.Drawing.Size(261, 20)
        Me.txtBillingAddress2.TabIndex = 2
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.BackColor = System.Drawing.SystemColors.Control
        Me.Label11.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label11.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label11.Location = New System.Drawing.Point(6, 55)
        Me.Label11.Name = "Label11"
        Me.Label11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label11.Size = New System.Drawing.Size(82, 14)
        Me.Label11.TabIndex = 66
        Me.Label11.Text = "BillingAddress2"
        '
        'textDriversLicence
        '
        Me.textDriversLicence.Location = New System.Drawing.Point(458, 141)
        Me.textDriversLicence.MaxLength = 20
        Me.textDriversLicence.Name = "textDriversLicence"
        Me.textDriversLicence.Size = New System.Drawing.Size(126, 20)
        Me.textDriversLicence.TabIndex = 16
        Me.textDriversLicence.Tag = "(No Auto Formatting)"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(430, 144)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(26, 14)
        Me.Label15.TabIndex = 63
        Me.Label15.Text = "DL#"
        '
        'txtTaxID
        '
        Me.txtTaxID.Location = New System.Drawing.Point(398, 121)
        Me.txtTaxID.Name = "txtTaxID"
        Me.txtTaxID.Size = New System.Drawing.Size(120, 20)
        Me.txtTaxID.TabIndex = 14
        Me.txtTaxID.Tag = "(No Auto Formatting)"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(359, 121)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(36, 14)
        Me.Label6.TabIndex = 55
        Me.Label6.Text = "Tax ID"
        '
        'chkShipBillSame
        '
        Me.chkShipBillSame.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.HelpProvider1.SetHelpString(Me.chkShipBillSame, "Check to make Ship to same as custome or uncheck to clear ship to data.")
        Me.chkShipBillSame.Location = New System.Drawing.Point(5, 143)
        Me.chkShipBillSame.Name = "chkShipBillSame"
        Me.HelpProvider1.SetShowHelp(Me.chkShipBillSame, True)
        Me.chkShipBillSame.Size = New System.Drawing.Size(142, 16)
        Me.chkShipBillSame.TabIndex = 54
        Me.chkShipBillSame.Text = "Ship To Same as Bill To"
        '
        'lblLine
        '
        Me.lblLine.BackColor = System.Drawing.Color.Gray
        Me.lblLine.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLine.Location = New System.Drawing.Point(1, 72)
        Me.lblLine.Name = "lblLine"
        Me.lblLine.Size = New System.Drawing.Size(586, 1)
        Me.lblLine.TabIndex = 53
        '
        'txtPONbr
        '
        Me.txtPONbr.Location = New System.Drawing.Point(322, 141)
        Me.txtPONbr.Name = "txtPONbr"
        Me.txtPONbr.Size = New System.Drawing.Size(104, 20)
        Me.txtPONbr.TabIndex = 15
        Me.txtPONbr.Tag = "(No Auto Formatting)"
        Me.txtPONbr.Text = "None"
        '
        'lblPONbr
        '
        Me.lblPONbr.AutoSize = True
        Me.lblPONbr.Location = New System.Drawing.Point(282, 144)
        Me.lblPONbr.Name = "lblPONbr"
        Me.lblPONbr.Size = New System.Drawing.Size(30, 14)
        Me.lblPONbr.TabIndex = 18
        Me.lblPONbr.Text = "PO #"
        '
        'txtShipZip
        '
        Me.txtShipZip.AcceptsReturn = True
        Me.txtShipZip.BackColor = System.Drawing.SystemColors.Window
        Me.txtShipZip.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtShipZip.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtShipZip.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtShipZip.Location = New System.Drawing.Point(464, 97)
        Me.txtShipZip.MaxLength = 0
        Me.txtShipZip.Name = "txtShipZip"
        Me.txtShipZip.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtShipZip.Size = New System.Drawing.Size(91, 20)
        Me.txtShipZip.TabIndex = 12
        Me.txtShipZip.Tag = "(No Auto Formatting)"
        '
        'txtShipState
        '
        Me.txtShipState.AcceptsReturn = True
        Me.txtShipState.BackColor = System.Drawing.SystemColors.Window
        Me.txtShipState.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtShipState.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtShipState.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtShipState.Location = New System.Drawing.Point(399, 97)
        Me.txtShipState.MaxLength = 2
        Me.txtShipState.Name = "txtShipState"
        Me.txtShipState.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtShipState.Size = New System.Drawing.Size(27, 20)
        Me.txtShipState.TabIndex = 11
        Me.txtShipState.Tag = "(No Auto Formatting)"
        '
        'txtShipCity
        '
        Me.txtShipCity.AcceptsReturn = True
        Me.txtShipCity.BackColor = System.Drawing.SystemColors.Window
        Me.txtShipCity.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtShipCity.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtShipCity.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtShipCity.Location = New System.Drawing.Point(399, 76)
        Me.txtShipCity.MaxLength = 0
        Me.txtShipCity.Name = "txtShipCity"
        Me.txtShipCity.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtShipCity.Size = New System.Drawing.Size(173, 20)
        Me.txtShipCity.TabIndex = 10
        Me.txtShipCity.Tag = "(No Auto Formatting)"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(435, 99)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(25, 14)
        Me.Label3.TabIndex = 50
        Me.Label3.Text = "Zip:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(361, 99)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(35, 14)
        Me.Label4.TabIndex = 49
        Me.Label4.Text = "State:"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(361, 79)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(28, 14)
        Me.Label5.TabIndex = 48
        Me.Label5.Text = "City:"
        '
        'txtContactName
        '
        Me.txtContactName.AcceptsReturn = True
        Me.txtContactName.BackColor = System.Drawing.SystemColors.Window
        Me.txtContactName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtContactName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtContactName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtContactName.Location = New System.Drawing.Point(89, 119)
        Me.txtContactName.MaxLength = 0
        Me.txtContactName.Name = "txtContactName"
        Me.txtContactName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtContactName.Size = New System.Drawing.Size(113, 20)
        Me.txtContactName.TabIndex = 9
        Me.txtContactName.Tag = "(No Auto Formatting)"
        '
        'txtPostalCode
        '
        Me.txtPostalCode.AcceptsReturn = True
        Me.txtPostalCode.BackColor = System.Drawing.SystemColors.Window
        Me.txtPostalCode.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPostalCode.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPostalCode.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPostalCode.Location = New System.Drawing.Point(460, 31)
        Me.txtPostalCode.MaxLength = 0
        Me.txtPostalCode.Name = "txtPostalCode"
        Me.txtPostalCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPostalCode.Size = New System.Drawing.Size(91, 20)
        Me.txtPostalCode.TabIndex = 5
        '
        'txtState
        '
        Me.txtState.AcceptsReturn = True
        Me.txtState.BackColor = System.Drawing.SystemColors.Window
        Me.txtState.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtState.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtState.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtState.Location = New System.Drawing.Point(400, 31)
        Me.txtState.MaxLength = 2
        Me.txtState.Name = "txtState"
        Me.txtState.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtState.Size = New System.Drawing.Size(27, 20)
        Me.txtState.TabIndex = 4
        '
        'txtShipAddress1
        '
        Me.txtShipAddress1.AcceptsReturn = True
        Me.txtShipAddress1.BackColor = System.Drawing.SystemColors.Window
        Me.txtShipAddress1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtShipAddress1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtShipAddress1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtShipAddress1.Location = New System.Drawing.Point(90, 97)
        Me.txtShipAddress1.MaxLength = 0
        Me.txtShipAddress1.Name = "txtShipAddress1"
        Me.txtShipAddress1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtShipAddress1.Size = New System.Drawing.Size(261, 20)
        Me.txtShipAddress1.TabIndex = 8
        Me.txtShipAddress1.Tag = "(No Auto Formatting)"
        '
        'txtCity
        '
        Me.txtCity.AcceptsReturn = True
        Me.txtCity.BackColor = System.Drawing.SystemColors.Window
        Me.txtCity.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCity.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCity.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCity.Location = New System.Drawing.Point(400, 11)
        Me.txtCity.MaxLength = 0
        Me.txtCity.Name = "txtCity"
        Me.txtCity.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCity.Size = New System.Drawing.Size(173, 20)
        Me.txtCity.TabIndex = 3
        '
        'txtShipToCustomer
        '
        Me.txtShipToCustomer.AcceptsReturn = True
        Me.txtShipToCustomer.BackColor = System.Drawing.SystemColors.Window
        Me.txtShipToCustomer.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtShipToCustomer.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtShipToCustomer.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtShipToCustomer.Location = New System.Drawing.Point(90, 76)
        Me.txtShipToCustomer.MaxLength = 0
        Me.txtShipToCustomer.Name = "txtShipToCustomer"
        Me.txtShipToCustomer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtShipToCustomer.Size = New System.Drawing.Size(261, 20)
        Me.txtShipToCustomer.TabIndex = 7
        Me.txtShipToCustomer.Tag = "(No Auto Formatting)"
        '
        'txtBillingAddress1
        '
        Me.txtBillingAddress1.AcceptsReturn = True
        Me.txtBillingAddress1.BackColor = System.Drawing.SystemColors.Window
        Me.txtBillingAddress1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtBillingAddress1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBillingAddress1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtBillingAddress1.Location = New System.Drawing.Point(90, 31)
        Me.txtBillingAddress1.MaxLength = 0
        Me.txtBillingAddress1.Name = "txtBillingAddress1"
        Me.txtBillingAddress1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtBillingAddress1.Size = New System.Drawing.Size(261, 20)
        Me.txtBillingAddress1.TabIndex = 1
        '
        'txtCompanyName
        '
        Me.txtCompanyName.AcceptsReturn = True
        Me.txtCompanyName.BackColor = System.Drawing.SystemColors.Window
        Me.txtCompanyName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCompanyName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCompanyName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCompanyName.Location = New System.Drawing.Point(90, 11)
        Me.txtCompanyName.MaxLength = 0
        Me.txtCompanyName.Name = "txtCompanyName"
        Me.txtCompanyName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCompanyName.Size = New System.Drawing.Size(261, 20)
        Me.txtCompanyName.TabIndex = 0
        '
        'lblContact
        '
        Me.lblContact.AutoSize = True
        Me.lblContact.BackColor = System.Drawing.SystemColors.Control
        Me.lblContact.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblContact.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblContact.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblContact.Location = New System.Drawing.Point(3, 121)
        Me.lblContact.Name = "lblContact"
        Me.lblContact.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblContact.Size = New System.Drawing.Size(77, 14)
        Me.lblContact.TabIndex = 37
        Me.lblContact.Text = "Contact Name:"
        '
        '_lblLabels_6
        '
        Me._lblLabels_6.AutoSize = True
        Me._lblLabels_6.BackColor = System.Drawing.SystemColors.Control
        Me._lblLabels_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_6, CType(6, Short))
        Me._lblLabels_6.Location = New System.Drawing.Point(436, 33)
        Me._lblLabels_6.Name = "_lblLabels_6"
        Me._lblLabels_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_6.Size = New System.Drawing.Size(25, 14)
        Me._lblLabels_6.TabIndex = 23
        Me._lblLabels_6.Text = "Zip:"
        '
        '_lblLabels_5
        '
        Me._lblLabels_5.AutoSize = True
        Me._lblLabels_5.BackColor = System.Drawing.SystemColors.Control
        Me._lblLabels_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_5, CType(5, Short))
        Me._lblLabels_5.Location = New System.Drawing.Point(362, 33)
        Me._lblLabels_5.Name = "_lblLabels_5"
        Me._lblLabels_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_5.Size = New System.Drawing.Size(35, 14)
        Me._lblLabels_5.TabIndex = 22
        Me._lblLabels_5.Text = "State:"
        '
        '_lblLabels_4
        '
        Me._lblLabels_4.AutoSize = True
        Me._lblLabels_4.BackColor = System.Drawing.SystemColors.Control
        Me._lblLabels_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_4, CType(4, Short))
        Me._lblLabels_4.Location = New System.Drawing.Point(362, 14)
        Me._lblLabels_4.Name = "_lblLabels_4"
        Me._lblLabels_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_4.Size = New System.Drawing.Size(28, 14)
        Me._lblLabels_4.TabIndex = 21
        Me._lblLabels_4.Text = "City:"
        '
        '_lblLabels_3
        '
        Me._lblLabels_3.AutoSize = True
        Me._lblLabels_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblLabels_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblLabels_3.Location = New System.Drawing.Point(5, 100)
        Me._lblLabels_3.Name = "_lblLabels_3"
        Me._lblLabels_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_3.Size = New System.Drawing.Size(58, 14)
        Me._lblLabels_3.TabIndex = 20
        Me._lblLabels_3.Text = "Addresss:"
        '
        '_lblLabels_2
        '
        Me._lblLabels_2.AutoSize = True
        Me._lblLabels_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblLabels_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblLabels_2.Location = New System.Drawing.Point(5, 79)
        Me._lblLabels_2.Name = "_lblLabels_2"
        Me._lblLabels_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_2.Size = New System.Drawing.Size(45, 14)
        Me._lblLabels_2.TabIndex = 19
        Me._lblLabels_2.Text = "Ship To:"
        '
        '_lblLabels_1
        '
        Me._lblLabels_1.AutoSize = True
        Me._lblLabels_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblLabels_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_1, CType(1, Short))
        Me._lblLabels_1.Location = New System.Drawing.Point(5, 32)
        Me._lblLabels_1.Name = "_lblLabels_1"
        Me._lblLabels_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_1.Size = New System.Drawing.Size(85, 14)
        Me._lblLabels_1.TabIndex = 18
        Me._lblLabels_1.Text = "BillingAddress1:"
        '
        '_lblLabels_0
        '
        Me._lblLabels_0.AutoSize = True
        Me._lblLabels_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblLabels_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_0, CType(0, Short))
        Me._lblLabels_0.Location = New System.Drawing.Point(5, 12)
        Me._lblLabels_0.Name = "_lblLabels_0"
        Me._lblLabels_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_0.Size = New System.Drawing.Size(82, 14)
        Me._lblLabels_0.TabIndex = 17
        Me._lblLabels_0.Text = "CompanyName:"
        '
        'lblCustId
        '
        Me.lblCustId.AutoSize = True
        Me.lblCustId.BackColor = System.Drawing.SystemColors.Control
        Me.lblCustId.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCustId.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCustId.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCustId.Location = New System.Drawing.Point(388, 52)
        Me.lblCustId.Name = "lblCustId"
        Me.lblCustId.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCustId.Size = New System.Drawing.Size(65, 14)
        Me.lblCustId.TabIndex = 36
        Me.lblCustId.Text = "Customer ID"
        '
        'txtCustomerID
        '
        Me.txtCustomerID.AcceptsReturn = True
        Me.txtCustomerID.BackColor = System.Drawing.SystemColors.Window
        Me.txtCustomerID.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCustomerID.Enabled = False
        Me.txtCustomerID.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustomerID.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCustomerID.Location = New System.Drawing.Point(460, 52)
        Me.txtCustomerID.MaxLength = 0
        Me.txtCustomerID.Name = "txtCustomerID"
        Me.txtCustomerID.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCustomerID.Size = New System.Drawing.Size(63, 20)
        Me.txtCustomerID.TabIndex = 6
        Me.txtCustomerID.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.btnHelp)
        Me.Frame1.Controls.Add(Me.chkCashCustomer)
        Me.Frame1.Controls.Add(Me.SSOleDBCombo1)
        Me.Frame1.Controls.Add(Me.cmdAdd)
        Me.Frame1.Controls.Add(Me._lblFieldLable_0)
        Me.Frame1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(5, 3)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(588, 46)
        Me.Frame1.TabIndex = 14
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Customer Name Search"
        '
        'btnHelp
        '
        Me.btnHelp.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.btnHelp.Location = New System.Drawing.Point(558, 15)
        Me.btnHelp.Name = "btnHelp"
        Me.btnHelp.Size = New System.Drawing.Size(25, 23)
        Me.btnHelp.TabIndex = 17
        Me.btnHelp.Text = "?"
        Me.btnHelp.UseVisualStyleBackColor = True
        '
        'SSOleDBCombo1
        '
        Me.SSOleDBCombo1.Location = New System.Drawing.Point(80, 15)
        Me.SSOleDBCombo1.Name = "SSOleDBCombo1"
        Me.SSOleDBCombo1.Size = New System.Drawing.Size(296, 22)
        Me.SSOleDBCombo1.TabIndex = 0
        '
        '_lblFieldLable_0
        '
        Me._lblFieldLable_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblFieldLable_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblFieldLable_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblFieldLable_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFieldLable.SetIndex(Me._lblFieldLable_0, CType(0, Short))
        Me._lblFieldLable_0.Location = New System.Drawing.Point(12, 13)
        Me._lblFieldLable_0.Name = "_lblFieldLable_0"
        Me._lblFieldLable_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblFieldLable_0.Size = New System.Drawing.Size(56, 30)
        Me._lblFieldLable_0.TabIndex = 15
        Me._lblFieldLable_0.Text = "Select Customer"
        '
        'txtDelivery
        '
        Me.txtDelivery.Location = New System.Drawing.Point(497, 438)
        Me.txtDelivery.Name = "txtDelivery"
        Me.txtDelivery.Size = New System.Drawing.Size(91, 20)
        Me.txtDelivery.TabIndex = 16
        Me.txtDelivery.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblDelivery
        '
        Me.lblDelivery.AutoSize = True
        Me.lblDelivery.Location = New System.Drawing.Point(449, 441)
        Me.lblDelivery.Name = "lblDelivery"
        Me.lblDelivery.Size = New System.Drawing.Size(46, 14)
        Me.lblDelivery.TabIndex = 46
        Me.lblDelivery.Text = "Delivery"
        '
        'lblTax
        '
        Me.lblTax.AutoSize = True
        Me.lblTax.Location = New System.Drawing.Point(441, 463)
        Me.lblTax.Name = "lblTax"
        Me.lblTax.Size = New System.Drawing.Size(54, 14)
        Me.lblTax.TabIndex = 47
        Me.lblTax.Text = "Sales Tax"
        '
        'txtSalesTax
        '
        Me.txtSalesTax.Location = New System.Drawing.Point(497, 460)
        Me.txtSalesTax.Name = "txtSalesTax"
        Me.txtSalesTax.Size = New System.Drawing.Size(91, 20)
        Me.txtSalesTax.TabIndex = 17
        Me.txtSalesTax.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtDeposit
        '
        Me.txtDeposit.Location = New System.Drawing.Point(497, 482)
        Me.txtDeposit.Name = "txtDeposit"
        Me.txtDeposit.Size = New System.Drawing.Size(91, 20)
        Me.txtDeposit.TabIndex = 18
        Me.txtDeposit.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblDeposit
        '
        Me.lblDeposit.AutoSize = True
        Me.lblDeposit.Location = New System.Drawing.Point(449, 486)
        Me.lblDeposit.Name = "lblDeposit"
        Me.lblDeposit.Size = New System.Drawing.Size(43, 14)
        Me.lblDeposit.TabIndex = 50
        Me.lblDeposit.Text = "Deposit"
        '
        'lblTotal
        '
        Me.lblTotal.AutoSize = True
        Me.lblTotal.Location = New System.Drawing.Point(465, 509)
        Me.lblTotal.Name = "lblTotal"
        Me.lblTotal.Size = New System.Drawing.Size(29, 14)
        Me.lblTotal.TabIndex = 51
        Me.lblTotal.Text = "Total"
        Me.lblTotal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtTotal
        '
        Me.txtTotal.Location = New System.Drawing.Point(497, 504)
        Me.txtTotal.Name = "txtTotal"
        Me.txtTotal.Size = New System.Drawing.Size(91, 20)
        Me.txtTotal.TabIndex = 19
        Me.txtTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblItemTotal
        '
        Me.lblItemTotal.AutoSize = True
        Me.lblItemTotal.Location = New System.Drawing.Point(441, 422)
        Me.lblItemTotal.Name = "lblItemTotal"
        Me.lblItemTotal.Size = New System.Drawing.Size(51, 14)
        Me.lblItemTotal.TabIndex = 53
        Me.lblItemTotal.Text = "Item Total"
        '
        'txtItemTotal
        '
        Me.txtItemTotal.Location = New System.Drawing.Point(497, 417)
        Me.txtItemTotal.Name = "txtItemTotal"
        Me.txtItemTotal.Size = New System.Drawing.Size(91, 20)
        Me.txtItemTotal.TabIndex = 15
        Me.txtItemTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.optCash)
        Me.GroupBox1.Controls.Add(Me.optLeftCardNumber)
        Me.GroupBox1.Controls.Add(Me.optBillTo)
        Me.GroupBox1.Controls.Add(Me.optLeftBlankCheck)
        Me.GroupBox1.Controls.Add(Me.optPaidByCreditCard)
        Me.GroupBox1.Controls.Add(Me.optPaidByCheck)
        Me.GroupBox1.Location = New System.Drawing.Point(6, 422)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(272, 77)
        Me.GroupBox1.TabIndex = 55
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Payment Arrangement"
        '
        'optCash
        '
        Me.optCash.Location = New System.Drawing.Point(136, 56)
        Me.optCash.Name = "optCash"
        Me.optCash.Size = New System.Drawing.Size(96, 16)
        Me.optCash.TabIndex = 5
        Me.optCash.Text = "Paid by Cash"
        '
        'optLeftCardNumber
        '
        Me.optLeftCardNumber.Location = New System.Drawing.Point(9, 56)
        Me.optLeftCardNumber.Name = "optLeftCardNumber"
        Me.optLeftCardNumber.Size = New System.Drawing.Size(110, 16)
        Me.optLeftCardNumber.TabIndex = 4
        Me.optLeftCardNumber.Text = "Left Card Number"
        '
        'optBillTo
        '
        Me.optBillTo.Location = New System.Drawing.Point(137, 37)
        Me.optBillTo.Name = "optBillTo"
        Me.optBillTo.Size = New System.Drawing.Size(104, 16)
        Me.optBillTo.TabIndex = 3
        Me.optBillTo.Text = "Bill To Customer"
        '
        'optLeftBlankCheck
        '
        Me.optLeftBlankCheck.Location = New System.Drawing.Point(137, 16)
        Me.optLeftBlankCheck.Name = "optLeftBlankCheck"
        Me.optLeftBlankCheck.Size = New System.Drawing.Size(106, 16)
        Me.optLeftBlankCheck.TabIndex = 2
        Me.optLeftBlankCheck.Text = "Left Blank Check"
        '
        'optPaidByCreditCard
        '
        Me.optPaidByCreditCard.Location = New System.Drawing.Point(9, 36)
        Me.optPaidByCreditCard.Name = "optPaidByCreditCard"
        Me.optPaidByCreditCard.Size = New System.Drawing.Size(118, 16)
        Me.optPaidByCreditCard.TabIndex = 1
        Me.optPaidByCreditCard.Text = "Paid by Credit Card"
        '
        'optPaidByCheck
        '
        Me.optPaidByCheck.Location = New System.Drawing.Point(8, 16)
        Me.optPaidByCheck.Name = "optPaidByCheck"
        Me.optPaidByCheck.Size = New System.Drawing.Size(94, 16)
        Me.optPaidByCheck.TabIndex = 0
        Me.optPaidByCheck.Text = "Paid by Check"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(7, 527)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(106, 14)
        Me.Label2.TabIndex = 56
        Me.Label2.Text = "Check/Card  Number"
        '
        'txtCheckNumber
        '
        Me.HelpProvider1.SetHelpString(Me.txtCheckNumber, "Enter check or credit card number if left.")
        Me.txtCheckNumber.Location = New System.Drawing.Point(9, 545)
        Me.txtCheckNumber.MaxLength = 16
        Me.txtCheckNumber.Name = "txtCheckNumber"
        Me.HelpProvider1.SetShowHelp(Me.txtCheckNumber, True)
        Me.txtCheckNumber.Size = New System.Drawing.Size(103, 20)
        Me.txtCheckNumber.TabIndex = 57
        Me.txtCheckNumber.Tag = "(No Auto Formatting)"
        '
        'lblAmtPaid
        '
        Me.lblAmtPaid.AutoSize = True
        Me.lblAmtPaid.Location = New System.Drawing.Point(449, 534)
        Me.lblAmtPaid.Name = "lblAmtPaid"
        Me.lblAmtPaid.Size = New System.Drawing.Size(49, 14)
        Me.lblAmtPaid.TabIndex = 58
        Me.lblAmtPaid.Text = "Amt Paid"
        '
        'txtAmtPaid
        '
        Me.HelpProvider1.SetHelpString(Me.txtAmtPaid, "Enter 0 if no payment made.")
        Me.txtAmtPaid.Location = New System.Drawing.Point(498, 531)
        Me.txtAmtPaid.Name = "txtAmtPaid"
        Me.HelpProvider1.SetShowHelp(Me.txtAmtPaid, True)
        Me.txtAmtPaid.Size = New System.Drawing.Size(91, 20)
        Me.txtAmtPaid.TabIndex = 20
        Me.txtAmtPaid.Tag = "(No Auto Formatting)"
        Me.txtAmtPaid.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtBalDue
        '
        Me.txtBalDue.Location = New System.Drawing.Point(498, 558)
        Me.txtBalDue.Name = "txtBalDue"
        Me.txtBalDue.Size = New System.Drawing.Size(91, 20)
        Me.txtBalDue.TabIndex = 20
        Me.txtBalDue.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblBalDue
        '
        Me.lblBalDue.AutoSize = True
        Me.lblBalDue.Location = New System.Drawing.Point(451, 561)
        Me.lblBalDue.Name = "lblBalDue"
        Me.lblBalDue.Size = New System.Drawing.Size(44, 14)
        Me.lblBalDue.TabIndex = 60
        Me.lblBalDue.Text = "Bal Due"
        '
        'lblLine2
        '
        Me.lblLine2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLine2.Location = New System.Drawing.Point(493, 527)
        Me.lblLine2.Name = "lblLine2"
        Me.lblLine2.Size = New System.Drawing.Size(95, 2)
        Me.lblLine2.TabIndex = 62
        '
        'btnCheckout
        '
        Me.btnCheckout.BackColor = System.Drawing.SystemColors.Control
        Me.btnCheckout.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnCheckout.Enabled = False
        Me.btnCheckout.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCheckout.ForeColor = System.Drawing.SystemColors.ControlText
        Me.HelpProvider1.SetHelpString(Me.btnCheckout, "Print contract and put on rent.")
        Me.btnCheckout.Location = New System.Drawing.Point(297, 551)
        Me.btnCheckout.Name = "btnCheckout"
        Me.btnCheckout.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.HelpProvider1.SetShowHelp(Me.btnCheckout, True)
        Me.btnCheckout.Size = New System.Drawing.Size(131, 22)
        Me.btnCheckout.TabIndex = 132
        Me.btnCheckout.Text = "Chec&kout"
        Me.btnCheckout.UseVisualStyleBackColor = False
        '
        'btnManualRecalc
        '
        Me.btnManualRecalc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnManualRecalc.Location = New System.Drawing.Point(297, 595)
        Me.btnManualRecalc.Name = "btnManualRecalc"
        Me.btnManualRecalc.Size = New System.Drawing.Size(131, 22)
        Me.btnManualRecalc.TabIndex = 131
        Me.btnManualRecalc.Text = "Manual Recalc"
        '
        'btnPrintInvoice
        '
        Me.btnPrintInvoice.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPrintInvoice.Location = New System.Drawing.Point(297, 529)
        Me.btnPrintInvoice.Name = "btnPrintInvoice"
        Me.btnPrintInvoice.Size = New System.Drawing.Size(131, 22)
        Me.btnPrintInvoice.TabIndex = 134
        Me.btnPrintInvoice.Text = "&Print Invoice"
        '
        'mnuMain
        '
        Me.mnuMain.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem3})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuPrint, Me.MenuItem2, Me.mnuLoadCustomerMaintForm, Me.mnuDeleteSelectedRow, Me.mnuExit})
        Me.MenuItem1.Text = "&File"
        '
        'mnuPrint
        '
        Me.mnuPrint.Index = 0
        Me.mnuPrint.Shortcut = System.Windows.Forms.Shortcut.CtrlP
        Me.mnuPrint.Text = "Print"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 1
        Me.MenuItem2.Shortcut = System.Windows.Forms.Shortcut.CtrlO
        Me.MenuItem2.Text = "Check Out"
        '
        'mnuLoadCustomerMaintForm
        '
        Me.mnuLoadCustomerMaintForm.Index = 2
        Me.mnuLoadCustomerMaintForm.Shortcut = System.Windows.Forms.Shortcut.CtrlL
        Me.mnuLoadCustomerMaintForm.Text = "Load Customer Maintenance Form"
        '
        'mnuDeleteSelectedRow
        '
        Me.mnuDeleteSelectedRow.Index = 3
        Me.mnuDeleteSelectedRow.Text = "&Delete Selected Item"
        '
        'mnuExit
        '
        Me.mnuExit.Index = 4
        Me.mnuExit.Shortcut = System.Windows.Forms.Shortcut.CtrlC
        Me.mnuExit.Text = "Exit"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 1
        Me.MenuItem3.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuAddCustomer})
        Me.MenuItem3.Text = "Help"
        '
        'mnuAddCustomer
        '
        Me.mnuAddCustomer.Index = 0
        Me.mnuAddCustomer.Text = "Adding a Customer"
        '
        'btnOtherCharges
        '
        Me.btnOtherCharges.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOtherCharges.Location = New System.Drawing.Point(297, 573)
        Me.btnOtherCharges.Name = "btnOtherCharges"
        Me.btnOtherCharges.Size = New System.Drawing.Size(131, 22)
        Me.btnOtherCharges.TabIndex = 137
        Me.btnOtherCharges.Text = "&Labor, Fuel, Supp"
        '
        'lblNotes
        '
        Me.lblNotes.AutoSize = True
        Me.lblNotes.Location = New System.Drawing.Point(8, 570)
        Me.lblNotes.Name = "lblNotes"
        Me.lblNotes.Size = New System.Drawing.Size(72, 14)
        Me.lblNotes.TabIndex = 140
        Me.lblNotes.Text = "Invoice Notes"
        '
        'txtNotes
        '
        Me.txtNotes.Location = New System.Drawing.Point(8, 588)
        Me.txtNotes.MaxLength = 255
        Me.txtNotes.Multiline = True
        Me.txtNotes.Name = "txtNotes"
        Me.txtNotes.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtNotes.Size = New System.Drawing.Size(269, 49)
        Me.txtNotes.TabIndex = 139
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(120, 527)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(48, 14)
        Me.Label7.TabIndex = 141
        Me.Label7.Text = "Exp Mon"
        '
        'cbExpMon
        '
        Me.cbExpMon.Items.AddRange(New Object() {"01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"})
        Me.cbExpMon.Location = New System.Drawing.Point(120, 544)
        Me.cbExpMon.Name = "cbExpMon"
        Me.cbExpMon.Size = New System.Drawing.Size(48, 22)
        Me.cbExpMon.TabIndex = 142
        '
        'cbExpYr
        '
        Me.cbExpYr.Items.AddRange(New Object() {"03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "40", "42", "42", "43", "44", "45", "46", "47", "48,", "49,", "50"})
        Me.cbExpYr.Location = New System.Drawing.Point(178, 544)
        Me.cbExpYr.Name = "cbExpYr"
        Me.cbExpYr.Size = New System.Drawing.Size(48, 22)
        Me.cbExpYr.TabIndex = 143
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(180, 527)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(40, 14)
        Me.Label8.TabIndex = 144
        Me.Label8.Text = "Exp Yr"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(231, 527)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(42, 14)
        Me.Label9.TabIndex = 145
        Me.Label9.Text = "Card ID"
        '
        'txtCardId
        '
        Me.txtCardId.Location = New System.Drawing.Point(239, 545)
        Me.txtCardId.MaxLength = 4
        Me.txtCardId.Name = "txtCardId"
        Me.txtCardId.Size = New System.Drawing.Size(40, 20)
        Me.txtCardId.TabIndex = 146
        Me.txtCardId.Tag = "(No Auto Formatting)"
        '
        'cbEmployees
        '
        Me.cbEmployees.Location = New System.Drawing.Point(290, 437)
        Me.cbEmployees.Name = "cbEmployees"
        Me.cbEmployees.Size = New System.Drawing.Size(135, 22)
        Me.cbEmployees.TabIndex = 147
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(289, 420)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(53, 14)
        Me.Label10.TabIndex = 148
        Me.Label10.Text = "Employee"
        '
        'dtpCkOutDateReset
        '
        Me.dtpCkOutDateReset.CustomFormat = "MM/dd/yyyy hh:mm tt"
        Me.dtpCkOutDateReset.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpCkOutDateReset.Location = New System.Drawing.Point(288, 480)
        Me.dtpCkOutDateReset.Name = "dtpCkOutDateReset"
        Me.dtpCkOutDateReset.Size = New System.Drawing.Size(136, 20)
        Me.dtpCkOutDateReset.TabIndex = 149
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(288, 464)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(125, 13)
        Me.Label12.TabIndex = 150
        Me.Label12.Text = "Ck Out Date-Time Set"
        '
        'chkPrintToFile
        '
        Me.chkPrintToFile.AutoSize = True
        Me.chkPrintToFile.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkPrintToFile.ForeColor = System.Drawing.Color.Red
        Me.chkPrintToFile.Location = New System.Drawing.Point(297, 508)
        Me.chkPrintToFile.Name = "chkPrintToFile"
        Me.chkPrintToFile.Size = New System.Drawing.Size(95, 18)
        Me.chkPrintToFile.TabIndex = 151
        Me.chkPrintToFile.Text = "Print to File?"
        Me.chkPrintToFile.UseVisualStyleBackColor = True
        '
        'btnShowPrinters
        '
        Me.btnShowPrinters.Location = New System.Drawing.Point(498, 584)
        Me.btnShowPrinters.Name = "btnShowPrinters"
        Me.btnShowPrinters.Size = New System.Drawing.Size(75, 23)
        Me.btnShowPrinters.TabIndex = 152
        Me.btnShowPrinters.Text = "Button1"
        Me.btnShowPrinters.UseVisualStyleBackColor = True
        Me.btnShowPrinters.Visible = False
        '
        'chkSaveCreditCard
        '
        Me.chkSaveCreditCard.AutoSize = True
        Me.chkSaveCreditCard.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSaveCreditCard.ForeColor = System.Drawing.Color.Red
        Me.chkSaveCreditCard.Location = New System.Drawing.Point(16, 507)
        Me.chkSaveCreditCard.Name = "chkSaveCreditCard"
        Me.chkSaveCreditCard.Size = New System.Drawing.Size(189, 18)
        Me.chkSaveCreditCard.TabIndex = 153
        Me.chkSaveCreditCard.Text = "Save Credit Card and DL Info?"
        Me.chkSaveCreditCard.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.AutoScroll = True
        Me.Panel1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.Panel1.Controls.Add(Me.Frame1)
        Me.Panel1.Controls.Add(Me.lblNotes)
        Me.Panel1.Controls.Add(Me.txtNotes)
        Me.Panel1.Controls.Add(Me.btnShowPrinters)
        Me.Panel1.Controls.Add(Me.chkSaveCreditCard)
        Me.Panel1.Controls.Add(Me.Frame2)
        Me.Panel1.Controls.Add(Me.fraDelivery)
        Me.Panel1.Controls.Add(Me.btnOtherCharges)
        Me.Panel1.Controls.Add(Me.chkPrintToFile)
        Me.Panel1.Controls.Add(Me.btnPrintInvoice)
        Me.Panel1.Controls.Add(Me.btnCheckout)
        Me.Panel1.Controls.Add(Me.dbgShoppingList)
        Me.Panel1.Controls.Add(Me.btnManualRecalc)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.cmdCancel)
        Me.Panel1.Controls.Add(Me.Label12)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.dtpCkOutDateReset)
        Me.Panel1.Controls.Add(Me.Label10)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.txtCardId)
        Me.Panel1.Controls.Add(Me.lblBalDue)
        Me.Panel1.Controls.Add(Me.cbExpYr)
        Me.Panel1.Controls.Add(Me.cbEmployees)
        Me.Panel1.Controls.Add(Me.cbExpMon)
        Me.Panel1.Controls.Add(Me.lblAmtPaid)
        Me.Panel1.Controls.Add(Me.lblItemTotal)
        Me.Panel1.Controls.Add(Me.txtDelivery)
        Me.Panel1.Controls.Add(Me.txtSalesTax)
        Me.Panel1.Controls.Add(Me.lblTotal)
        Me.Panel1.Controls.Add(Me.txtDeposit)
        Me.Panel1.Controls.Add(Me.txtCheckNumber)
        Me.Panel1.Controls.Add(Me.lblDeposit)
        Me.Panel1.Controls.Add(Me.txtTotal)
        Me.Panel1.Controls.Add(Me.lblTax)
        Me.Panel1.Controls.Add(Me.txtItemTotal)
        Me.Panel1.Controls.Add(Me.lblDelivery)
        Me.Panel1.Controls.Add(Me.txtAmtPaid)
        Me.Panel1.Controls.Add(Me.txtBalDue)
        Me.Panel1.Controls.Add(Me.lblLine2)
        Me.Panel1.Location = New System.Drawing.Point(10, 1)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(623, 650)
        Me.Panel1.TabIndex = 154
        '
        'frmCustomers
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(634, 653)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.HelpButton = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(141, 141)
        Me.MaximizeBox = False
        Me.Menu = Me.mnuMain
        Me.MinimizeBox = False
        Me.Name = "frmCustomers"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Customer Checkout"
        CType(Me.dbgShoppingList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraDelivery.ResumeLayout(False)
        Me.fraDelivery.PerformLayout()
        Me.Frame2.ResumeLayout(False)
        Me.Frame2.PerformLayout()
        Me.Frame1.ResumeLayout(False)
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblFieldLable, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblLabels, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
#End Region

#Region "Module Level Variables"
    Private ZeroBalanceOk As Boolean = False
    Public dtFC As DataTable
    Dim mbFormLoading As Boolean = True
    Private oDA As CDataAccess
    Dim iInvID As Integer
    Dim iHitRow As Integer
    Private mbPrinted As Boolean
    Private mbClosing As Boolean
    Dim miRentals As Short
    Dim miSales As Short
    Dim mbAllSales As Boolean
    Private oCG As New CGrid()
    Private ignoreKeyPreview As Boolean
    Public CheckOutEmployee As String
    Public CustomerPhone As String
#End Region

#Region "Processing Methods"
    Private Sub AddCustomerData()
        Dim SQL As String
        Dim sErr As String
        Dim dt As New DataTable()


        Try
            With Me
                ' First, ensure that we don't already have a customer id 
                ' selected
                If .txtCustomerID.Text.Length > 0 Then
                    SQL = "select customerid,companyname from customers "
                    SQL &= "where customerid = " & Replace(.txtCustomerID.Text, "'", "''")
                    If oDA.SendQuery(SQL, dt, ConnectString) > 0 Then
                        If CompanyName.ToUpper = dt.Rows(0).Item(1) Then
                            Dim sMsg As String
                            Dim iRV As Integer
                            sMsg = "It appears that you have selected a customer and then" & Chr(10)
                            sMsg &= "clicked the Add Button by mistake.  You cannot add a " & Chr(10)
                            sMsg &= "customer that already exists." & Chr(10)
                            sMsg &= "" & Chr(10)
                            sMsg &= "If you are trying to add a new customer, click the OK" & Chr(10)
                            sMsg &= "button and then enter the new customer data.  Then" & Chr(10)
                            sMsg &= "click the Add button." & Chr(10)
                            sMsg &= "" & Chr(10)
                            sMsg &= "If you clicked the Add Button by mistake click the " & Chr(10)
                            sMsg &= "Cancel button." & Chr(10)
                            sMsg &= "" & Chr(10)
                            iRV = MsgBox(sMsg, CType(305, Microsoft.VisualBasic.MsgBoxStyle), "Customer Already Exists")

                            If iRV = 1 Then
                                ' Ok Code goes here
                                Me.txtCustomerID.Text = String.Empty
                                Me.txtCompanyName.Text = String.Empty
                                Exit Sub
                            Else
                                ' Cancel code goes here
                                Exit Sub
                            End If
                        Else
                            Dim sMsg As String
                            Dim iRV As Integer
                            sMsg = "There is a Customer ID present and it exists in the " & Chr(10)
                            sMsg &= "database, but the Customer Name does not match the" & Chr(10)
                            sMsg &= "Customer ID.  If you selected a customer and then" & Chr(10)
                            sMsg &= "decided to enter new customer data into the database," & Chr(10)
                            sMsg &= "a new Customer ID will have to be assigned." & Chr(10)
                            sMsg &= "" & Chr(10)
                            sMsg &= "If you have entered new customer data in the " & Chr(10)
                            sMsg &= "appropriate text boxes and want to add the new " & Chr(10)
                            sMsg &= "customer to the database, click OK." & Chr(10)
                            sMsg &= "" & Chr(10)
                            sMsg &= "If you do not want to add the changed data to the" & Chr(10)
                            sMsg &= "database, click Cancel and reselect the customer." & Chr(10)
                            sMsg &= "" & Chr(10)
                            iRV = MsgBox(sMsg, CType(305, Microsoft.VisualBasic.MsgBoxStyle), "Customer ID Present")

                            If iRV = 1 Then
                                ' Ok Code goes here
                            Else
                                ' Cancel code goes here
                                Exit Sub
                            End If
                        End If
                    Else
                        Dim sMsg As String
                        Dim iRV As Integer
                        sMsg = "There is a Customer ID in the text box that does not" & Chr(10)
                        sMsg &= "exist in the database.  The system must assign the " & Chr(10)
                        sMsg &= "ID for a new customer.  " & Chr(10)
                        sMsg &= "" & Chr(10)
                        sMsg &= "If you are trying to add a new customer, and you have" & Chr(10)
                        sMsg &= "entered the required Customer Data into the text boxes," & Chr(10)
                        sMsg &= "click OK." & Chr(10)
                        sMsg &= "" & Chr(10)
                        sMsg &= "If you have mistakedly clicked the Add Button, click the" & Chr(10)
                        sMsg &= "Cancel Button and reselect the customer." & Chr(10)
                        sMsg &= "" & Chr(10)
                        iRV = MsgBox(sMsg, CType(305, Microsoft.VisualBasic.MsgBoxStyle), "Confirm Customer Add")

                        If iRV = 1 Then
                            ' Ok Code goes here

                        Else
                            ' Cancel code goes here
                            Exit Sub
                        End If
                    End If

                End If

                If .txtCompanyName.Text.Length = 0 Or _
                   .txtPhone.Text.Length = 0 Or _
                   .txtContactName.Text.Length = 0 Then
                    MsgBox("You must enter a minimum of Company Name, Phone, and Contact Name in order to add a new customer to the Database.", MsgBoxStyle.Exclamation)
                    Exit Sub
                End If

                SQL = "select max(customerid) from customers"
                dt.Reset()
                If oDA.SendQuery(SQL, dt, ConnectString) = 0 OrElse _
                   IsDBNull(dt.Rows(0).Item(0)) Then
                    Throw New System.Exception("Failed to retrieve max(customerid)")
                End If
                Me.txtCustomerID.Text = dt.Rows(0).Item(0) + 1

                SQL = "insert into customers "
                SQL &= "(companyname, customerid, Contactname,Billingaddress1, billingaddress2, "
                SQL &= "city,state,postalcode,Phonenumber,Tax_id) "
                SQL &= "values('"
                SQL &= .txtCompanyName.Text.Replace("'", "''") & "', "
                SQL &= .txtCustomerID.Text.Replace("'", "''") & ", "
                SQL &= "'" & .txtContactName.Text.Replace("'", "''") & "', "
                SQL &= "'" & .txtBillingAddress1.Text.Replace("'", "''") & "', "
                SQL &= "'" & .txtBillingAddress2.Text.Replace("'", "''") & "', "
                SQL &= "'" & .txtCity.Text.Replace("'", "''") & "', "
                SQL &= "'" & .txtState.Text.Replace("'", "") & "', "
                SQL &= "'" & .txtPostalCode.Text.Replace("'", "") & "', "
                SQL &= "'" & .txtPhone.Text.Replace("'", "") & "', "
                SQL &= "'" & .txtTaxID.Text.Replace("'", "") & "') "
            End With

            If oDA.SendActionSql(SQL, ConnectString, sErr) < 1 Then
                MsgBox("Add of customer data failed.  " & Chr(10) & sErr, MsgBoxStyle.Critical)
                Exit Sub
            End If
            Dim s As String = Me.txtCompanyName.Text
            Me.LoadCustomerCombo()
            Me.SSOleDBCombo1.Text = s
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub GetInvoiceID()
        Dim sql As String
        ' first we want to create invoice header and details
        sql = "select count(*) from invoices"
        Dim dt As New DataTable()
        If oDA.SendQuery(sql, dt, ConnectString) < 1 Then
            Throw New System.Exception("Can't create invoice header")
        End If
        iInvID = dt.Rows(0).Item(0)
        If iInvID = 0 Then
            iInvID = 1001
        Else
            sql = "select max(invoiceid) from invoices"
            dt.Reset()
            If oDA.SendQuery(sql, dt, ConnectString) < 1 Then
                Throw New System.Exception("Can't create invoice header")
            End If
            iInvID = dt.Rows(0).Item(0) + 1
        End If
    End Sub

    Private Function ValidateRequiredFields() As Boolean
        Dim i As Integer

        Try
            With Me
                If Me.dbgShoppingList.DataSource Is Nothing OrElse _
                   Me.dbgShoppingList.DataSource.rows.count = 0 Then
                    MsgBox("You don't have any items selected to rent or sell.", MsgBoxStyle.Exclamation)
                    Return False
                End If

                If Me.chkCashCustomer.Checked Then
                    If .txtContactName.Text.Trim.Length = 0 Or _
                       .txtPhone.Text.Trim.Length = 0 Then
                        Dim sMsg As String
                        Dim iRV As Integer
                        sMsg = "For cash customers, you should enter at least the" & Chr(10)
                        sMsg &= "Contact Name, Phone Number, and Drivers License #." & Chr(10)
                        sMsg &= "" & Chr(10)
                        sMsg &= "Click Ok to proceed with the checkout without this" & Chr(10)
                        sMsg &= "information or click Cancel to enter the Contact Name" & Chr(10)
                        sMsg &= "and Phone Number." & Chr(10)
                        sMsg &= "" & Chr(10)
                        iRV = MsgBox(sMsg, CType(305, Microsoft.VisualBasic.MsgBoxStyle), "Confirm Leaving Contact Name Blank")

                        If iRV = 1 Then
                            ' Ok Code goes here
                        Else
                            ' Cancel code goes here
                            Return False
                        End If
                    End If
                End If

                If Me.txtCompanyName.Text.Length > 0 And Me.txtCustomerID.Text.Length = 0 Then
                    Dim sMsg As String
                    sMsg = "You have either entered a new customer and failed" & Chr(10)
                    sMsg &= "to press the ADD button or you have selected a " & Chr(10)
                    sMsg &= "customer and then entered data in the Customer" & Chr(10)
                    sMsg &= "Name box.  " & Chr(10)
                    sMsg &= "" & Chr(10)
                    sMsg &= "In either case, you currently have no Customer  ID " & Chr(10)
                    sMsg &= "selected." & Chr(10)
                    sMsg &= "" & Chr(10)
                    sMsg &= "Please select an existing customer or enter a new " & Chr(10)
                    sMsg &= "customer and press the Add Button." & Chr(10)
                    sMsg &= "" & Chr(10)
                    MsgBox(sMsg, CType(48, Microsoft.VisualBasic.MsgBoxStyle), "No Customer Selected")
                    Exit Function
                End If

                If Me.cbEmployees.Text.Length = 0 Then
                    MsgBox("Please choose your employee name for checkout.", MsgBoxStyle.Information)
                    Return False
                End If
                If Me.txtCompanyName.Text.Trim.Length = 0 Then
                    MsgBox("Please select a customer from the customer drop down list.", MsgBoxStyle.Exclamation)
                    Return False
                End If

                If (Me.chkChargeTax.Checked = False AndAlso _
                    Me.txtTaxID.Text.Trim.Length = 0) Or _
                    (Me.chkChargeTax.Checked = True AndAlso _
                    Me.txtTaxID.Text.Trim.Length > 0) Then
                    Dim sMsg As String
                    sMsg = "If you turn off sales tax, you must enter a Tax ID." & Chr(10)
                    sMsg &= "Also, if you have a tax id, you can't charge sales tax." & Chr(10)
                    sMsg &= "" & Chr(10)
                    sMsg &= "To enter or delete a tax id, click the Add button in the" & Chr(10)
                    sMsg &= "customer area and enter or delete the tax id for the selected" & Chr(10)
                    sMsg &= "customer.  Next, close the customer maintenance form" & Chr(10)
                    sMsg &= "and reselect the customer." & Chr(10)
                    sMsg &= "" & Chr(10)
                    MsgBox(sMsg, CType(48, Microsoft.VisualBasic.MsgBoxStyle), "Tax ID Required")
                    Return False
                End If

                If Me.txtCheckNumber.Text.Trim.Length = 0 And _
                   (Me.optLeftBlankCheck.Checked Or _
                   Me.optLeftCardNumber.Checked Or _
                   Me.optPaidByCheck.Checked Or _
                   Me.optPaidByCreditCard.Checked Or _
                   Me.txtCheckNumber.Text.Trim.ToUpper = "N/A") Then
                    MsgBox("Please enter a valid check number.", MsgBoxStyle.Exclamation)
                    Return False
                End If
                If Me.txtShipToCustomer.Text.Length = 0 Then
                    MsgBox("You must enter the Ship To Customer Name, Address, etc., or click the 'Ship to same as Bill To' check box.", MsgBoxStyle.Exclamation)
                    Return False
                End If

                ' determine if we have only sale items
                For i = 0 To Me.dtFC.Rows.Count - 1
                    With dtFC.Rows(i)
                        If .Item("rentorsale") = RENT Or .Item("itemid") = RERENT Then
                            miRentals += 1
                        Else
                            miSales += 1
                        End If
                    End With
                Next
                mbAllSales = (miSales > 0) And (miRentals = 0)
                If Not Me.optPaidByCheck.Checked AndAlso Not Me.optPaidByCreditCard.Checked AndAlso Not Me.optBillTo.Checked AndAlso Not Me.optCash.Checked AndAlso Not Me.optLeftBlankCheck.Checked AndAlso Not optLeftCardNumber.Checked Then
                    MsgBox("You must select a Payment Arrangement Option", MsgBoxStyle.Exclamation)
                    Return False
                End If

                If mbAllSales Then
                    If (Me.optLeftBlankCheck.Checked Or Me.optLeftCardNumber.Checked) Then
                        Dim sMsg As String
                        sMsg = "For invoices with sale items only, the items must either" & Chr(10)
                        sMsg &= "be paid for or billed.  You cannot leave a blank check" & Chr(10)
                        sMsg &= "or credit card for sale items only." & Chr(10)
                        sMsg &= "" & Chr(10)
                        sMsg &= "Please select the proper payment option and click the" & Chr(10)
                        sMsg &= "Print or Check Out button again." & Chr(10)
                        sMsg &= "" & Chr(10)
                        MsgBox(sMsg, CType(48, Microsoft.VisualBasic.MsgBoxStyle), "Select Payment Type")
                        Return False
                    End If
                End If

                ' make sure if Billed to, we have a balance
                If Me.optBillTo.Checked Or _
                   Me.optLeftBlankCheck.Checked Or _
                   Me.optLeftCardNumber.Checked Then
                    If UnFormat(Me.txtBalDue.Text) = 0 And Not ZeroBalanceOk Then
                        Dim sMsg As String
                        Dim iRV As Integer
                        sMsg = "You have no balance due, but you have not " & Chr(10)
                        sMsg &= "checked a payment option that tells how the " & Chr(10)
                        sMsg &= "invoice was paid." & Chr(10)
                        sMsg &= "" & Chr(10)
                        sMsg &= "Click Ok to check out anyway, or Cancel to " & Chr(10)
                        sMsg &= "cancel the CheckOut." & Chr(10)
                        sMsg &= "" & Chr(10)
                        iRV = MsgBox(sMsg, CType(33, Microsoft.VisualBasic.MsgBoxStyle), "Confirm Zero Balance")

                        If iRV = 1 Then
                            ' Ok Code goes here
                            ZeroBalanceOk = True
                        Else
                            ' Cancel code goes here
                            Return False
                        End If
                    End If
                End If
                If iInvID = 0 Then
                    Me.GetInvoiceID()
                End If
                Return True
            End With
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Function


    Private Sub CheckOut()
        ' Here we check them out, we must:
        ' 1) Mark each Item rented
        ' 2) Update customer information
        ' 2) Print the Chekout bill
        Dim SQL As String
        Dim i As Short
        Dim oRES As New CTransaction()
        Dim b As Boolean
        Dim oCust As New CCustomer()
        Dim sReport As String
        Dim sTitle As String
        Dim sSubTitle As String
        Dim sOrginization As String
        Dim sFooter As String
        Dim rs As New DataTable()
        Dim o As CItems
        Dim sErr As String
        Dim iRows As Integer
        Dim shtPmtType As Short = 0
        Dim iCustID As Integer = Val(Me.txtCustomerID.Text)
        Dim dtDueBack As DateTime
        Dim sngPeriod As Single
        Dim myTransaction As OleDb.OleDbTransaction
        Dim conn As OleDb.OleDbConnection
        ' Record type for invoice details
        '1 5=equip,25=deposit,35=tax,45=delivery,55=AmtPaid(only if pmt made),
        ' 65=refund,75=bal due, 79=discount

        ' First, we mark the items rented

        Try
            If Me.chkSaveCreditCard.Checked Then
                If String.IsNullOrEmpty(txtCheckNumber.Text) OrElse String.IsNullOrEmpty(cbExpMon.Text) OrElse String.IsNullOrEmpty(cbExpYr.Text) OrElse String.IsNullOrEmpty(txtCardId.Text) Then
                    MessageBox.Show("You have checked that you want to save the Credit Card Info but not all boxes are filled.  Either fill the boxes or uncheck the Save CheckBox.", "Missing Credit Card Info", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Return
                End If
            End If

            ' set up connection for transaction
            conn = New OleDb.OleDbConnection(ConnectString)
            Try
                conn.Open()
            Catch
                MsgBox("Can't open connection to database.", MsgBoxStyle.Critical)
                Exit Sub
            End Try
            myTransaction = conn.BeginTransaction
            Dim cmd As New OleDb.OleDbCommand()
            cmd.Transaction = myTransaction
            cmd.Connection = conn


            ' turn off the flag that causes cancel to prompt
            mbPrinted = False

            b = oRES.PlaceShoppingCartListOnRent()

            Dim sPaidBy As String = ""

            ' ******* Create Invoice Header Record ***************
            SQL = "insert into invoices "
            SQL &= "(invoiceid,customerid, Status, InvoiceDate,PONumber,CkCardNumber, ContactName, PaidOption, "
            SQL &= "ShipToCustomer,ShipToAddress,ShipToCity,ShipToState, "
            SQL &= "BalanceDue,ShipToZip,Notes,Exp_Month,Exp_Yr,Card_id,"
            SQL &= "check_out_employee,Drivers_License) "
            SQL &= "values(" & iInvID & ",'" & Me.txtCustomerID.Text.Trim & "', "

            ' if all sales and paid
            If mbAllSales Then
                If UnFormat(Me.txtBalDue.Text) = 0 Then
                    SQL &= "'CLOSED', "
                Else
                    SQL &= "'OPEN', "
                End If
            Else
                If Me.chkCkOutAndIN.Checked Then
                    If UnFormat(Me.txtBalDue.Text) = 0 Then
                        SQL &= "'CLOSED', "
                    Else
                        SQL &= "'OPEN', "
                    End If
                Else
                    SQL &= "'CheckedOut', "
                End If
            End If

            SQL &= "#" & Me.dtpCkOutDateReset.Value.ToString & "#, "

            ' if cash customer put phone number in po number
            If Me.chkCashCustomer.Checked Then
                SQL &= "'" & Me.txtPhone.Text.Replace("'", "''") & "', "
            Else
                SQL &= "'" & Me.txtPONbr.Text.Replace("'", "''").Trim & "', "
            End If

            SQL &= "'" & Me.txtCheckNumber.Text.Replace("'", "''").Trim & "', "
            SQL &= "'" & Me.txtContactName.Text.Replace("'", "''").Trim & "', "
            Select Case True
                Case Me.optBillTo.Checked : sPaidBy = "BT"
                Case Me.optCash.Checked : sPaidBy = "CA"
                Case Me.optLeftBlankCheck.Checked : sPaidBy = "BC"
                Case Me.optLeftCardNumber.Checked : sPaidBy = "LC"
                Case Me.optPaidByCheck.Checked : sPaidBy = "CK"
                Case Me.optPaidByCreditCard.Checked : sPaidBy = "CC"
                Case Else : sPaidBy = ""
            End Select
            SQL &= "'" & sPaidBy & "', "
            SQL &= "'" & Me.txtShipToCustomer.Text.Replace("'", "''") & "', "
            SQL &= "'" & Me.txtShipAddress1.Text.Replace("'", "''") & "', "
            SQL &= "'" & Me.txtShipCity.Text.Replace("'", "''") & "', "
            SQL &= "'" & Me.txtShipState.Text.Replace("'", "''") & "', "
            SQL &= UnFormat(Me.txtBalDue.Text).ToString & ", "
            SQL &= "'" & Me.txtShipZip.Text.Replace("'", "''") & "', "
            SQL &= "'" & Me.txtNotes.Text.Replace("'", "''") & "', "
            SQL &= "'" & MNS(Me.cbExpMon.Text.Replace("'", "''")) & "', "
            SQL &= "'" & MNS(Me.cbExpYr.Text.Replace("'", "''")) & "', "
            SQL &= "'" & MNS(Me.txtCardId.Text.Replace("'", "''")) & "', "
            SQL &= "'" & Me.cbEmployees.Text & "',"
            SQL &= "'" & MNS(Me.textDriversLicence.Text.Replace("'", "''")) & "'"
            SQL &= ")"
            cmd.CommandText = SQL
            If oDA.SendActionSql(cmd, sErr) < 1 Then
                Throw New System.Exception("Unable to create Invoice header: " & iInvID.ToString & "," & sErr & vbCrLf & SQL & vbCrLf)
            End If

            ' clear the temp items table
            SQL = "delete from tempitems where user_id = '" & UserName & "'"
            cmd.CommandText = SQL
            If oDA.SendActionSql(cmd, sErr) < 0 Then
                Throw New Exception("Failed to delete tempitems." & Chr(10) & SQL)
            End If
            mbClosing = True ' so we dont' confirm clearing

            ' include pmt type on every record for billing purposes
            Dim dr As DataRow
            For i = 0 To Me.dtFC.Rows.Count - 1
                ' itemid,itemname,price,deposit,itemtotal,rentalperiod,rentaltime
                ' priceperunit,priceperunit
                dr = dtFC.Rows(i)
                dtDueBack = CShare.GetDueBackTime(dr("itemperiod"), dr("itemcount"), Me.dtpCkOutDateReset.Value)

                SQL = "insert into invoice_details "
                SQL &= "(invoiceid, equip_id, equip_name, deposit, rental_period,"
                SQL &= "priceperunit, quantity, rented_date, rentalduetoreturn, record_type, record_description,"
                SQL &= "Customer_Id,meter_out,"
                ' 8 new fields
                SQL &= "hourrate,halfday,daily,weekly,monthly,weekend,newprices,rerent_id"
                SQL &= ") "
                SQL &= "values("
                SQL &= iInvID & ", "
                SQL &= "'" & dr("ItemID") & "', "
                SQL &= "'" & Replace(dr("ItemName"), "'", "''") & "', "
                If Me.chkManualDepositOverride.Checked Then
                    SQL &= "0, "
                Else
                    SQL &= dr("ItemDeposit") & ", "
                End If
                SQL &= "'" & dr("ItemPeriod") & "', "  ' day, halfday, etc
                SQL &= dr("ItemPrice") & ", "
                SQL &= dr("ItemCount") & ", "  ' count
                SQL &= "#" & Me.dtpCkOutDateReset.Value.ToString & "#, "
                SQL &= "#" & dtDueBack & "#, "
                ' for rerents, store the po in record desc
                If dr("itemid") <> RERENT Then
                    SQL &= "15, 'Rent/Sale Item', "
                Else
                    SQL &= "15, '" & dr("rentorsale") & "', "
                End If
                SQL &= iCustID.ToString & ", "
                If IsDBNull(dr("hour_meter")) Then
                    SQL &= "0, "
                Else
                    SQL &= dr("hour_meter") & ", "
                End If
                ' new price fields
                If IsDBNull(dr("newprices")) Then
                    SQL &= "0,0,0,0,0,0,False,"
                Else
                    SQL &= dr("hourrate") & ", " & _
                           dr("halfday") & ", " & _
                           dr("daily") & ", " & _
                           dr("weekly") & ", " & _
                           dr("monthly") & ", " & _
                           dr("weekend") & ", " & _
                           True & ", "
                End If
                If IsDBNull(dr("rerent_id")) Then
                    SQL &= "0"
                Else
                    SQL &= dr("rerent_id")
                End If
                SQL &= ")"
                cmd.CommandText = SQL
                iRows = oDA.SendActionSql(cmd, sErr)

                If iRows < 1 Then
                    Throw New Exception("Invoice Detail Update Failure: " & Chr(10) & _
                       sErr & Chr(10) & SQL)
                End If

                ' update meter table if applicalble
                If MNSng(dr("hour_meter")) > 0 Then
                    SQL = "insert into  meter_reading"
                    SQL &= "(equip_id, meter_reading, date_entered, invoice_id,entry_type) "
                    SQL &= "values("
                    SQL &= "'" & dr("itemid") & "', "
                    SQL &= dr("hour_meter") & ", "
                    SQL &= "#" & Me.dtpCkOutDateReset.Value.ToString & "#, "
                    SQL &= iInvID & ", "
                    SQL &= "'Out' "
                    SQL &= ")"
                    cmd.CommandText = SQL
                    iRows = oDA.SendActionSql(cmd, sErr)
                End If

                If dr("rentorsale") = RENT Or dr("itemid") = RERENT Then
                    If Not Me.chkCkOutAndIN.Checked Then
                        If dtFC.Rows(i).Item("itemid") <> RERENT Then
                            ' update the equipment table
                            SQL = "update equipment "
                            SQL &= "set rented_date = #" & Me.dtpCkOutDateReset.Value.ToString & "#, "
                            SQL &= "Available = 'ON RENT', "
                            SQL &= "Available_date= #" & dtDueBack.ToString & "#, "
                            SQL &= "renting_company_id = " & Me.txtCustomerID.Text & " "
                            SQL &= "where equip_id = '" & dr("itemid") & "'"
                            cmd.CommandText = SQL
                            iRows = oDA.SendActionSql(cmd, sErr)
                            If iRows < 1 Then
                                Throw New Exception("Equipment Table Update Failure: " & Chr(10) & _
                                   sErr & Chr(10) & SQL)
                            End If
                        End If
                    End If
                ElseIf dr("rentorsale") = "SOLD" Then
                    If Not Me.chkCkOutAndIN.Checked Then
                        ' update the equipment table
                        SQL = "update equipment "
                        SQL &= "set rented_date = #" & Me.dtpCkOutDateReset.Value.ToString & "#, "
                        SQL &= "Available = 'SOLD', "
                        SQL &= "Available_date= #" & dtDueBack.ToString & "#, "
                        SQL &= "renting_company_id = " & Me.txtCustomerID.Text & " "
                        SQL &= "where equip_id = '" & dr("itemid") & "'"
                        cmd.CommandText = SQL
                        iRows = oDA.SendActionSql(cmd, sErr)
                        If iRows < 1 Then
                            Throw New Exception("Equipment Table Update Failure: " & Chr(10) & _
                               sErr & Chr(10) & SQL)
                        End If
                    End If
                Else
                    If dr("itemid") <> "Labor" AndAlso _
                       dr("itemid") <> "Fuel" AndAlso _
                       dr("itemid") <> "Misc" Then
                        ' update the products table to lower the inventory count
                        SQL = "update products "
                        SQL &= "set unitsinstock = unitsinstock - " & dr("ItemCount") & " "
                        SQL &= "where productid= '" & dr("ItemId") & "'"
                        cmd.CommandText = SQL
                        If oDA.SendActionSql(cmd, sErr) <> 1 Then
                            Throw New Exception("Unable to update products table to reduce inventory count." & Chr(10) & SQL)
                        End If
                    End If
                End If
            Next i

            ' now determine deposit type if any
            If Val(UnFormat(Me.txtDeposit.Text)) > 0 Then
                SQL = "Insert into invoice_details "
                SQL &= "(invoiceid,record_type,record_description,Deposit, Customer_id) "
                SQL &= "values("
                SQL &= iInvID & ", "
                SQL &= "25, 'Deposit', " 'manual deposit record
                SQL &= UnFormat(Me.txtDeposit.Text) & ", "
                SQL &= iCustID.ToString
                SQL &= ")"
                cmd.CommandText = SQL
                iRows = oDA.SendActionSql(cmd, sErr)
                If iRows < 1 Then
                    Throw New Exception("Invoice Detail Update Failure: " & Chr(10) & _
                       sErr & Chr(10) & SQL)
                End If
            End If

            ' determine delivery if applicable
            If Val(UnFormat(Me.txtDelivery.Text)) > 0 Then
                SQL = "Insert into invoice_details "
                SQL &= "(invoiceid,record_type,record_description,delivery, customer_id) "
                SQL &= "values("
                SQL &= iInvID & ", "
                SQL &= "45, 'Delivery', " 'delivery
                SQL &= UnFormat(Me.txtDelivery.Text) & ", "
                'SQL &= shtPmtType.ToString & ", "
                SQL &= iCustID.ToString
                SQL &= ")"
                cmd.CommandText = SQL
                iRows = oDA.SendActionSql(cmd, sErr)
                If iRows < 1 Then
                    Throw New Exception("Invoice Detail Update Failure: " & Chr(10) & _
                       sErr & Chr(10) & SQL)
                End If
            End If

            ' create sales tax record if applicable
            If UnFormat(Me.txtSalesTax.Text) > 0 Then
                SQL = "Insert into invoice_details "
                SQL &= "(invoiceid,record_type,record_description,salestax, customer_id) "
                SQL &= "values("
                SQL &= iInvID & ", "
                SQL &= "35, 'Sales Tax', " ' sales tax record
                SQL &= UnFormat(Me.txtSalesTax.Text) & ", "
                'SQL &= shtPmtType.ToString & ", "
                SQL &= iCustID.ToString & " "
                'SQL &= "'" & Me.txtTaxID.Text & "' "
                SQL &= ")"
                cmd.CommandText = SQL
                iRows = oDA.SendActionSql(cmd, sErr)
                If iRows < 1 Then
                    Throw New Exception("Invoice Detail Update Failure: " & Chr(10) & _
                       sErr & Chr(10) & SQL)
                End If
            End If

            ' create amt paid record is applicable
            If Val(UnFormat(Me.txtAmtPaid.Text)) > 0 Then
                SQL = "Insert into invoice_details "
                SQL &= "(invoiceid,record_type,record_description,amtpaid, customer_id) "
                SQL &= "values("
                SQL &= iInvID & ", "
                SQL &= "55, 'Amt Paid at CkOut', " ' sales tax record
                SQL &= UnFormat(Me.txtAmtPaid.Text) & ", "
                SQL &= iCustID.ToString
                SQL &= ")"
                cmd.CommandText = SQL
                iRows = oDA.SendActionSql(cmd, sErr)
                If iRows < 1 Then
                    Throw New Exception("Invoice Detail Update Failure: " & Chr(10) & _
                       sErr & Chr(10) & SQL)
                End If
            End If

            SQL = String.Empty
            If chkSaveCreditCard.Checked Then
                SQL = "update customers set CreditCard = '" & StringEncryption.EncryptString(txtCheckNumber.Text) & "' "
                SQL &= ", CardExpires = '" & cbExpMon.Text & "/" & cbExpYr.Text & "' "
                SQL &= ", SecCode = '" & txtCardId.Text & "' "
            End If

            If Not String.IsNullOrEmpty(textDriversLicence.Text) Then
                If String.IsNullOrEmpty(SQL) Then
                    SQL = "update customers set DLNumber = '" & StringEncryption.EncryptString(textDriversLicence.Text) & "' "
                Else
                    SQL &= " , DLNumber = '" & StringEncryption.EncryptString(textDriversLicence.Text) & "' "
                End If
            End If

            If Not String.IsNullOrEmpty(txtTaxID.Text) Then
                If String.IsNullOrEmpty(SQL) Then
                    SQL = "update customers set tax_id = '" & StringEncryption.EncryptString(txtTaxID.Text) & "' "
                Else
                    SQL &= " , tax_id = '" & StringEncryption.EncryptString(txtTaxID.Text) & "' "
                End If
            End If

            If Not String.IsNullOrEmpty(SQL) Then
                SQL &= " where customerid = " & iCustID.ToString
            End If

            If Not String.IsNullOrEmpty(SQL) Then
                cmd.CommandText = SQL
                iRows = oDA.SendActionSql(cmd, sErr)
                If iRows < 1 Then
                    Throw New Exception("Customer Record Update Failure: " & Chr(10) & _
                       sErr & Chr(10) & SQL)
                End If

            End If

            myTransaction.Commit()
            Try
                conn.Close()
            Catch
            End Try

            Me.Close()
            DoEvents()
            modMain.fMainForm.LoadEquipGridFromType(0)

        Catch ex As System.Exception
            myTransaction.Rollback()
            StructuredErrorHandler(ex)
        End Try
    End Sub

    'Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    '   (ByVal hwnd As Integer, ByVal Msg As Integer, ByVal wParam As Integer, _
    '   <MarshalAs(UnmanagedType.AsAny)> ByVal lParam As Object) As Integer
    'Private Const WM_SYSCOMMAND As Integer = &H112
    'Private Const SC_CONTEXTHELP As Integer = &HF180


    'Private Sub btnHelp_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnHelp.Click
    '    btnHelp.Capture = False
    '    SendMessage(Me.Handle, WM_SYSCOMMAND, DirectCast(SC_CONTEXTHELP, IntPtr), IntPtr.Zero)
    'End Sub

    Private Sub PrintInvoice(ByVal InvoiceId As Integer)
        ' format the print line
        Dim ps As New System.Text.StringBuilder()
        Dim decEP As Decimal
        Dim SQL As String
        Dim i As Integer
        Dim dt As New DataTable()
        Dim oPD As CPioneerPrint
        Dim oUtil As New CUtilities()
        Dim decTotal As Decimal
        Dim sName As String
        ' get customer data and print

        Try
            '#If CustomerApp = RELIABLE Then
            '         Dim oCR As New CReliablePrint.CReliablePrint(Me)
            '         oCR.PrintCheckOutInvoice(InvoiceId)
            '#ElseIf CustomerApp = PIONEER Then
            Dim oPP As New CPioneerInvoice.CPioneerPrepInvoice(Me)
            oPP.PrintCheckOutInvoice(InvoiceId)
            '#End If
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

    Private Function CkKeyPressNumeric(ByRef riKeyAscii As Short, ByRef roTB As System.Windows.Forms.TextBox) As Short
        Dim liKeyReturn As Short
        ' allow 0-9,., Back, Del,-,Ins, and / if in tag format
        On Error Resume Next
        CkKeyPressNumeric = riKeyAscii
        If riKeyAscii = System.Windows.Forms.Keys.Back Or riKeyAscii = System.Windows.Forms.Keys.Insert Or riKeyAscii = System.Windows.Forms.Keys.Delete Or riKeyAscii = 46 Or (riKeyAscii >= System.Windows.Forms.Keys.D0 And riKeyAscii <= System.Windows.Forms.Keys.D9) Or riKeyAscii = 45 Or riKeyAscii = 46 Or (InStr(roTB.Tag, " / ") > 0 And riKeyAscii = System.Windows.Forms.Keys.Divide) Then
            If roTB.SelectionLength = 0 Then
                If InStr(roTB.Text, ".") > 0 Then
                    If Len(Mid(roTB.Text, InStr(roTB.Text, ".") + 1)) > 1 Then
                        System.Windows.Forms.SendKeys.SendWait("{TAB}")
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
    Public Function UnFmt_T_B(ByRef roTB As System.Windows.Forms.TextBox) As Object
        On Error Resume Next
        'If IsDate(roTB.Text) Then
        '   UnFmt_T_B = DateValue(roTB.Text)
        'Else
        UnFmt_T_B = Val(Replace(Replace(Replace(Replace(Replace(roTB.Text, "$", ""), ",", ""), ")", ""), "(", ""), "%", ""))
        If InStr(roTB.Text, "%") Then
            UnFmt_T_B = UnFmt_T_B / 100
        End If
        If InStr(roTB.Text, "(") > 0 And InStr(roTB.Text, ")") > 0 Then
            UnFmt_T_B = UnFmt_T_B * -1
        End If
        'End If
    End Function
    Public Function Fmt_T_B(ByRef roTB As System.Windows.Forms.TextBox) As String
        On Error Resume Next
        If InStr(1, roTB.Tag, ";", 1) > 0 Then
            If InStr(roTB.Text, "-") > 0 Or (InStr(roTB.Text, "(") > 0 And InStr(roTB.Text, ")") > 0) Then
                Fmt_T_B = VB6.Format(System.Math.Abs(CDbl(roTB.Text)), Mid(roTB.Tag, InStr(roTB.Tag, ";") + 1))
            Else
                Fmt_T_B = VB6.Format(roTB.Text, Mid(roTB.Tag, 1, InStr(roTB.Tag, ";") - 1))
            End If
        ElseIf InStr(1, roTB.Tag, "%", 1) > 0 Then
            Fmt_T_B = Format(roTB.Text, roTB.Tag)
        Else
            Fmt_T_B = Format(roTB.Text, roTB.Tag)
        End If
    End Function
    Public Function Fmt_D_F(ByRef rsTxt As Object, ByRef roTB As System.Windows.Forms.TextBox) As String
        On Error Resume Next

        If InStr(1, roTB.Tag, ";", 1) > 0 Then
            If InStr(rsTxt, "-") Then

                Fmt_D_F = VB6.Format(Replace(rsTxt, "-", ""), Mid(roTB.Tag, InStr(roTB.Tag, ";") + 1))
            Else
                Fmt_D_F = VB6.Format(rsTxt, Mid(roTB.Tag, 1, InStr(roTB.Tag, ";") - 1))
            End If
        ElseIf InStr(1, roTB.Tag, "%", 1) > 0 Then
            Fmt_D_F = VB6.Format(rsTxt, roTB.Tag)
        Else
            Fmt_D_F = VB6.Format(rsTxt, roTB.Tag)
        End If
    End Function


    Private Sub CancelHandler()
        Dim sMsg As String
        Dim iRV As Short
        Dim oRES As CTransaction
        Dim serr As String


        Try
            If mbPrinted Then
                sMsg = "You have printed or previewed the invoice.  Are " & Chr(10)
                sMsg &= "you sure you want to close without checking out?" & Chr(10)
                sMsg &= "" & Chr(10)
                sMsg &= "Click Yes to close the form without checkout, No" & Chr(10)
                sMsg &= "to automatically chechout the equipment, or Cancel" & Chr(10)
                sMsg &= "to cancel the Close operation." & Chr(10)
                sMsg &= "" & Chr(10)
                iRV = MsgBox(sMsg, CType(35, Microsoft.VisualBasic.MsgBoxStyle), "Confirm Checkout")

                If iRV = DialogResult.Yes Then
                    ' Yes Code goes here
                    sMsg = "Do you want unload the held items from the" & Chr(10)
                    sMsg = sMsg & "shopping cart?"
                    iRV = MsgBox(sMsg, 35, "Confirm Clearing Hold On Equipment")

                    If iRV = DialogResult.Yes Then
                        ' Yes Code goes here
                        oRES = New CTransaction()
                        Call oRES.RemoveTempReservation(True)
                        Dim SQL As String = ""
                        SQL &= "delete from TempItems "
                        Dim iRows As Integer = oDA.SendActionSql(SQL, ConnectString, serr)
                        'If iRows < 1 Then
                        '   'MsgBox("Failure to delete temporary items." & Chr(10) & SQL, MsgBoxStyle.Critical)
                        'End If
                        mbClosing = True
                        modMain.ReloadRentalGrid = True
                        Me.Close()
                        System.Windows.Forms.Application.DoEvents()
                    ElseIf iRV = DialogResult.No Then
                        ' No code goes here
                        mbClosing = True
                        Me.Close()
                        System.Windows.Forms.Application.DoEvents()
                    Else
                        ' Cancel code goes here
                    End If
                ElseIf iRV = DialogResult.No Then
                    ' No code goes here
                    CheckOut()
                Else
                    ' Cancel code goes here
                    Exit Sub
                End If
            Else
                sMsg = "Do you want unload the held items from the" & Chr(10)
                sMsg = sMsg & "shopping cart?"
                iRV = MsgBox(sMsg, 35, "Confirm Clearing Hold On Equipment")

                If iRV = DialogResult.Yes Then
                    ' Yes Code goes here
                    oRES = New CTransaction()
                    Call oRES.RemoveTempReservation(True)
                    Dim SQL As String = ""
                    'SQL &= "delete from TempItems "
                    'Dim iRows As Integer = oDA.SendActionSql(SQL, ConnectString, serr)
                    'If iRows < 1 Then
                    '   MsgBox("Failure to delete temporary items." & Chr(10) & SQL, MsgBoxStyle.Critical)
                    'End If
                    modMain.ReloadRentalGrid = True
                    mbClosing = True
                    Me.Close()
                    System.Windows.Forms.Application.DoEvents()
                ElseIf iRV = DialogResult.No Then
                    ' No code goes here
                    mbClosing = True
                    Me.Close()
                    System.Windows.Forms.Application.DoEvents()
                Else
                    ' Cancel code goes here
                    Exit Sub
                End If
            End If
            mbClosing = True
            Me.Close()
            System.Windows.Forms.Application.DoEvents()

        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub



    Private Sub frmCustomers_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        CenterForm(Me)
        Me.cbEmployees.Text = GetSetting(RENTALPRO, SETTINGS, "CKOUTEMP", "")
        mbFormLoading = True
        LoadCustomerCombo()
        With Me
            .txtPONbr.Text = "None"
            .txtCheckNumber.Text = ""
            .txtShipAddress1.Text = ""
            .txtShipCity.Text = ""
            .txtShipState.Text = ""
            .txtShipToCustomer.Text = ""
            .txtShipZip.Text = ""
            .txtCheckNumber.Text = ""
            .txtContactName.Text = "None"
            .txtTaxID.Text = ""
        End With
        LoadDeliveryCombo()
        LoadEmployeeCombo()
    End Sub
    ''' <summary>
    ''' Load Checkout Emploeee Combo box.
    ''' </summary>
    Private Sub LoadEmployeeCombo()

        Try
            Dim SQL As String = ""
            Dim dt As New DataTable()
            Dim i As Integer

            SQL &= "select employee_initials,employee_name  "
            SQL &= "from employee "
            SQL &= "order by employee_initials "
            If oDA.SendQuery(SQL, dt, ConnectString) > 0 Then
                For i = 0 To dt.Rows.Count - 1
                    With dt.Rows(i)
                        Me.cbEmployees.Items.Add(.Item("Employee_name"))
                    End With
                Next
            End If
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    ''' <summary>
    ''' Load the delivery prices and descriptions into combo.
    ''' </summary>
    Private Sub LoadDeliveryCombo()
        Dim SQL As String = ""
        Dim dt As New DataTable()
        Dim i As Integer

        SQL &= "select delivery_price,delivery_desc "
        SQL &= "from delivery_pickup "
        SQL &= "where delivery_pickup ='B' or "
        SQL &= "delivery_pickup='D' "
        SQL &= "order by delivery_price "
        If oDA.SendQuery(SQL, dt, ConnectString) > 0 Then
            For i = 0 To dt.Rows.Count - 1
                With dt.Rows(i)
                    Me.cboDeliveryDistance.Items.Add(FormatCurrency(.Item("delivery_price")) & _
                      " - " & .Item("delivery_desc"))
                End With
            Next
        End If
    End Sub


    ''' <summary>
    ''' Load the customer combo from the customers table.
    ''' </summary>
    Private Sub LoadCustomerCombo()
        Dim dt As New DataTable()
        Dim i As Integer
        Dim sql As String = "select companyname from customers "
        sql &= "order by companyname"
        oDA.SendQuery(sql, dt, ConnectString)
        Me.SSOleDBCombo1.Items.Clear()

        For i = 0 To dt.Rows.Count - 1
            With dt.Rows(i)
                Me.SSOleDBCombo1.Items.Add(.Item("companyname"))
            End With
        Next
    End Sub


    ''' <summary>
    ''' Load the detail grid from the TempItems table.
    ''' </summary>
    Private Sub LoadTheGrid()
        Dim i As Short
        Dim o As CItems
        Dim lsLine As String
        Dim lcurDeposit As Decimal = 0
        Dim lcurDelivery As Decimal = 0
        Dim lcurItemTotal As Decimal = 0
        Dim lcurPrice As Decimal = 0
        Dim lcurTax As Decimal = 0
        Dim SQL As String
        Dim dt As New DataTable()
        Dim dr As DataRow()
        Dim lTotal As Decimal
        Static bGridloaded As Boolean
        Dim laborCost As Decimal = 0

        Try

            lcurItemTotal = 0

            If Me.chkDeliveryRequired.Checked Then
                lcurDelivery = Val(Replace(GetToken((Me.cboDeliveryDistance.Text), " ", "D"), "$", ""))
            End If


            If Not bGridloaded Then
                dtFC = New DataTable("dt")

                SQL = "select * from tempitems where user_id = '" & UserName & "' order by ItemId"
                oDA.SendQuery(SQL, dtFC, ConnectString, "dt")
            End If

            ' loop thru dt to accumulate and calc totals
            ' ItemTotal, Delivery, Sales Tax, Deposit, Total
            If dtFC.Rows.Count > 0 Then
                For i = 0 To dtFC.Rows.Count - 1
                    With dtFC.Rows(i)
                        ' now compute the running total
                        lcurDeposit += .Item("ItemDeposit")
                        lcurPrice += .Item("ItemExtendedPrice")
                        If .Item("itemid") = "Labor" Then
                            laborCost += .Item("ItemExtendedPrice")
                        End If
                    End With
                Next i
                Dim formats() As String = _
                  {"", "60", "T", "L", _
                   "", "150", "T", "L", _
                   "0", "60", "T", "R", _
                   "", "60", "T", "L", _
                   "$#,##0.00", "80", "T", "R", _
                   "$#,##0.00", "80", "T", "R", _
                   "$#,##0.00", "80", "T", "R", _
                   "", "60", "T", "L", _
                   "", "60", "T", "L", _
                   "0.00", "60", "T", "R", _
                   "", "60", "T", "L"}
                oCG.SetTablesStyle(dtFC, Me.dbgShoppingList, formats)

                Me.dbgShoppingList.SetDataBinding(dtFC, "")
                oCG.DisableAddNew(dbgShoppingList, Me)

                ' loop thru the dt to see if any of the equip
                ' is damaged, if so put the damage desc into notes
                Dim dt2 As New DataTable()
                For i = 0 To dtFC.Rows.Count - 1
                    Dim dr2 As DataRow = dtFC.Rows(i)
                    With dr2
                        If dr2("rentorsale") = RENT Then
                            SQL = "select damage_desc from equipment "
                            SQL &= "where equip_id = '" & dr2("itemid") & "' "
                            SQL &= "and not isnull(damage_desc) and trim(damage_desc)<>'' "
                            dt2.Reset()
                            If oDA.SendQuery(SQL, dt2, ConnectString) > 0 Then
                                With dt2.Rows(0)
                                    Me.txtNotes.AppendText(dr2("itemid") & " - " & .Item("damage_desc") & vbCrLf)
                                End With
                            End If
                        End If
                    End With
                Next
                ' item total
                Me.txtItemTotal.Text = FormatCurrency(lcurPrice)

                ' delivery if applicable
                If UnFormat(Me.txtManualEntryDelPrice.Text) > 0 Then
                    lcurDelivery = UnFormat(Me.txtManualEntryDelPrice.Text)
                    Me.txtDelivery.Text = FormatCurrency(lcurDelivery)
                    lcurItemTotal += lcurDelivery
                Else
                    Me.txtDelivery.Text = ""
                End If

                ' compute sales tax if applicable
                If Me.chkChargeTax.Checked AndAlso
                    (lcurPrice > 0 Or lcurDelivery > 0) Then
                    lcurTax = (lcurPrice + lcurDelivery - laborCost) * TaxRate ' (laborCost + lcurDelivery)) * TaxRate
                    Me.txtSalesTax.Text = FormatCurrency(lcurTax)
                Else
                    lcurTax = 0
                    Me.txtSalesTax.Text = FormatCurrency(0)
                End If
                DoEvents()
                ' manual deposit if applicable
                If Me.chkManualDepositOverride.Checked Then
                    Me.txtDeposit.Text = FormatCurrency(UnFormat(Me.txtManualDeposit.Text))
                Else
                    Me.txtDeposit.Text = FormatCurrency(lcurDeposit)
                End If

                ' total line:
                ' Total      Price    Delivery   Deposit   Total
                Me.txtTotal.Text = FormatCurrency(lcurPrice + _
                                   lcurTax + _
                                   UnFormat(Me.txtDeposit.Text) + _
                                   lcurDelivery)
                Me.txtAmtPaid.Text = Me.txtTotal.Text
                Me.txtBalDue.Text = FormatCurrency(UnFormat(Me.txtTotal.Text) - _
                   UnFormat(Me.txtAmtPaid.Text))
                DoEvents()
            End If
            bGridloaded = True
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub



    Private Sub LoadCustomerBoxes(ByRef rsCustomerName As String)
        Dim SQL As String
        Dim dt As New DataTable()

        SQL = ""
        SQL = SQL & "select customerid,companyname,billingaddress1,billingaddress2, "
        SQL = SQL & "billingaddress3,city,state,postalcode,phonenumber, "
        SQL = SQL & "contactname,contacttitle,customerid,Tax_Id, EmailAddress,DLNumber, CardExpires, SecCode, CreditCard "
        SQL = SQL & "from customers "
        SQL = SQL & "where companyname = '" & Replace(rsCustomerName, "'", "''") & "'"
        oDA.SendQuery(SQL, dt, ConnectString)

        If dt.Rows.Count > 0 Then
            With dt.Rows(0)
                Me.txtCustomerID.Text = MNS(.Item("customerid"))
                Me.txtCompanyName.Text = MNS(.Item("companyname"))
                Me.txtBillingAddress1.Text = MNS(.Item("billingaddress1"))
                Me.txtBillingAddress2.Text = MNS(.Item("billingaddress2"))
                Me.txtCity.Text = MNS(.Item("city"))
                Me.txtState.Text = MNS(.Item("state"))
                Me.txtPostalCode.Text = MNS(.Item("postalcode"))
                Me.txtContactName.Text = MNS(.Item("contactname"))
                Dim tid As String = MNS(.Item("tax_id"))
                If Not String.IsNullOrEmpty(tid) Then
                    Try
                        Me.txtTaxID.Text = StringEncryption.DecryptString(tid)
                    Catch ex As System.Exception
                        Me.txtTaxID.Text = tid
                    End Try
                Else
                    Me.txtTaxID.Text = tid
                End If
                If Me.txtTaxID.Text.Trim.Length = 0 Then
                    Me.chkChargeTax.Checked = True
                Else
                    Me.chkChargeTax.Checked = False
                End If
                Me.CustomerPhone = MNS(.Item("phonenumber"))
                Dim dl As String = MNS(.Item("DLNumber"))
                Try
                    If Not String.IsNullOrEmpty(dl) Then
                        textDriversLicence.Text = StringEncryption.DecryptString(dl)
                    End If
                Catch ex As System.Exception
                    textDriversLicence.Text = dl
                End Try
                Dim cc As String = MNS(.Item("CreditCard"))

                If Not String.IsNullOrEmpty(cc) Then
                    Me.txtCheckNumber.Text = StringEncryption.DecryptString(cc)
                    Me.txtCardId.Text = MNS(.Item("SecCode"))
                    Dim expires As String = MNS(.Item("CardExpires"))
                    If expires.Length = 5 Then
                        Me.cbExpMon.Text = expires.Substring(0, 2)
                        Me.cbExpYr.Text = MNS(.Item("CardExpires")).Substring(3, 2)
                        Me.txtCardId.Text = MNS(.Item("SecCode"))
                    End If
                End If
            End With
        End If
    End Sub

    Private Sub LoadCustomerBoxes(ByVal CustId As Integer)
        Dim SQL As String
        Dim dt As New DataTable()

        SQL = ""
        SQL = SQL & "select customerid,companyname,billingaddress1,billingaddress2, "
        SQL = SQL & "billingaddress3,city,state,postalcode,phonenumber, "
        SQL = SQL & "contactname,contacttitle,customerid,Tax_Id,DLNumber,CreditCard,CardExpires,SecCode  "
        SQL = SQL & "from customers "
        SQL = SQL & "where customerid = " & CustId & " "
        oDA.SendQuery(SQL, dt, ConnectString)

        If dt.Rows.Count > 0 Then
            With dt.Rows(0)
                Me.SSOleDBCombo1.Text = MNS(.Item("companyname"))
                Me.txtCustomerID.Text = MNS(.Item("customerid"))
                Me.txtCompanyName.Text = MNS(.Item("companyname"))
                Me.txtBillingAddress1.Text = MNS(.Item("billingaddress1"))
                Me.txtBillingAddress2.Text = MNS(.Item("billingaddress2"))
                Me.txtCity.Text = MNS(.Item("city"))
                Me.txtState.Text = MNS(.Item("state"))
                Me.txtPostalCode.Text = MNS(.Item("postalcode"))
                Me.txtContactName.Text = MNS(.Item("contactname"))
                Dim tid As String = MNS(.Item("tax_id"))
                If Not String.IsNullOrEmpty(tid) Then
                    Try
                        Me.txtTaxID.Text = StringEncryption.DecryptString(tid)
                    Catch ex As Exception
                        Me.txtTaxID.Text = tid
                    End Try
                End If
                If Me.txtTaxID.Text.Trim.Length = 0 Then
                    Me.chkChargeTax.Checked = True
                Else
                    Me.chkChargeTax.Checked = False
                End If
                Me.CustomerPhone = MNS(.Item("phonenumber"))
                If MNS(.Item("CreditCard")).Trim = "" Then
                    Me.txtCardId.Text = String.Empty
                Else
                    Me.txtCardId.Text = StringEncryption.DecryptString(MNS(.Item("CreditCard")))
                End If
                If MNS(.Item("DLNumber")).Trim = String.Empty Then
                    Me.textDriversLicence.Text = String.Empty
                Else
                    Try
                        Me.textDriversLicence.Text = StringEncryption.DecryptString(MNS(.Item("DLNumber")))
                    Catch ex As System.Exception
                        Me.textDriversLicence.Text = MNS(.Item("DLNumber"))
                    End Try
                End If

            End With
        End If
    End Sub

    Private Sub MovePaidToDue()
        Me.txtBalDue.Text = Me.txtAmtPaid.Text
        Me.txtAmtPaid.Text = FormatCurrency(0)
    End Sub
    Private Sub MoveDueToPaid()
        Me.txtAmtPaid.Text = Me.txtBalDue.Text
        Me.txtBalDue.Text = FormatCurrency(0)
    End Sub

    Private Sub ManualRecalc()
        ' if discount box filled, just subtract the amount from the total
        ' and recompute the sales tax
        Dim amt As Decimal
        Dim tax As Decimal
        Dim laborCost As Decimal
        Dim lcurDeposit As Decimal
        Dim lcurPrice As Decimal
        Dim i As Integer
        If mbFormLoading Then Exit Sub
        With Me
            For i = 0 To dtFC.Rows.Count - 1
                With dtFC.Rows(i)
                    ' now compute the running total
                    lcurDeposit += .Item("ItemDeposit")
                    lcurPrice += .Item("ItemExtendedPrice")
                    If .Item("itemid") = "Labor" Then
                        laborCost += .Item("ItemExtendedPrice")
                    End If
                End With
            Next i

            amt = UnFormat(Me.txtItemTotal.Text) + _
                  UnFormat(.txtDelivery.Text)

            If Me.chkChargeTax.Checked AndAlso amt > 0 Then
                .txtSalesTax.Text = FormatCurrency((amt - laborCost) * TaxRate) ' (laborCost + UnFormat(.txtDelivery.Text))) * TaxRate)
            Else
                .txtSalesTax.Text = FormatCurrency(0)
            End If

            amt += UnFormat(.txtDeposit.Text) + _
                   UnFormat(.txtSalesTax.Text)

            Me.txtTotal.Text = FormatCurrency(amt)
            If .optBillTo.Checked Or _
               .optLeftBlankCheck.Checked Or _
               .optLeftCardNumber.Checked Then
                .txtAmtPaid.Text = FormatCurrency(0)
                .txtBalDue.Text = FormatCurrency(amt)
            Else
                .txtBalDue.Text = FormatCurrency(0)
                .txtAmtPaid.Text = FormatCurrency(amt)
            End If
        End With
    End Sub

    Private Sub CheckOutHandler()
        If Not Me.ValidateRequiredFields() Then Exit Sub
        CheckOut()
    End Sub

#End Region

#Region "Form & Control Events"
    Private Sub mnuLoadCustomerMaintForm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuLoadCustomerMaintForm.Click
        Dim oFrm As New frmCustomerMaintenance()
        oFrm.ShowDialog()
        LoadCustomerCombo()
    End Sub
    Private Sub cmdAdd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAdd.Click
        AddCustomerData()
    End Sub

    Private Sub frmCustomers_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress

        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 27 Then
            Me.Close()
            System.Windows.Forms.Application.DoEvents()
        End If
        If Not Me.ignoreKeyPreview Then
            Select Case UCase(Chr(KeyAscii))
                Case "D"
                    Me.mnuDeleteSelectedRow_Click(New Object(), New Object())
                Case "P"
                    Me.cmdPrintContract_Click(New Object(), New Object())
                Case "C"
                    Me.cmdCancel_Click(New Object(), New Object())
                Case "R"
                    Me.chkDeliveryRequired.Checked = IIf(Me.chkDeliveryRequired.Checked = False, True, False)
                Case "M"
                    Me.chkManualDepositOverride.Checked = IIf(Me.chkManualDepositOverride.Checked = False, True, False)
                Case "A"
                    Me.chkDepositRequired.Checked = IIf(Me.chkDepositRequired.Checked = True, False, True)
                Case "T"
                    Me.chkChargeTax.Checked = IIf(Me.chkChargeTax.Checked = False, True, False)
                    'Case "S"
                    '   Me.cmdSaveCustomerData_Click(New Object(), New Object())
                Case "K"
                    Me.btnCheckout_Click(New Object(), New Object())
                Case Else
                    Exit Sub
            End Select
            '   KeyAscii= 0
        End If
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Public Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
        CancelHandler()

    End Sub
    Private Sub chkChargeTax_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkChargeTax.CheckStateChanged

        If Not mbFormLoading Then
            LoadTheGrid()
            ManualRecalc()
        End If
    End Sub

    Private Sub chkDeliveryRequired_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkDeliveryRequired.CheckStateChanged
        If Not mbFormLoading Then
            On Error Resume Next
            Me.cboDeliveryDistance.SelectedIndex = 0
            System.Windows.Forms.Application.DoEvents()
            If Me.chkDeliveryRequired.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                Me.txtManualEntryDelPrice.Text = ""
            Else
                Me.txtManualEntryDelPrice.Text = FormatCurrency(UnFormat(GetToken((Me.cboDeliveryDistance.Text), " ", "D")))
            End If
            LoadTheGrid()
            ManualRecalc()
        End If
    End Sub

    Private Sub chkDepositRequired_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkDepositRequired.CheckedChanged
        If Not mbFormLoading Then
            If eventSender.Checked Then
                'If Me.chkDepositRequired.Value = vbChecked Then
                '   Me.cboDeliveryDistance.ListIndex = 0
                'End If
                LoadTheGrid()
            End If
        End If
    End Sub

    Private Sub frmCustomers_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        If mbFormLoading Then
            mbFormLoading = False
            Me.Top = 0
            LoadTheGrid()
        End If
    End Sub

    Private Sub txt_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtManualEntryDelPrice.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then KeyAscii = 0
        'UPGRADE_ISSUE: Assignment not supported: KeyAscii to a non-zero value Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1058"'
        KeyAscii = CkKeyPressNumeric(KeyAscii, txtManualEntryDelPrice)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txt_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtManualEntryDelPrice.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error Resume Next
        If KeyCode = 13 Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
        If KeyCode = System.Windows.Forms.Keys.Up Then System.Windows.Forms.SendKeys.SendWait("+{TAB}")
        If KeyCode = System.Windows.Forms.Keys.Down Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
    End Sub
    Private Sub txt_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtManualEntryDelPrice.Enter
        txtManualEntryDelPrice.Text = UnFmt_T_B(txtManualEntryDelPrice)
        txtManualEntryDelPrice.SelectionStart = 0
        txtManualEntryDelPrice.SelectionLength = Len(Trim(txtManualEntryDelPrice.Text))
    End Sub
    Private Sub txt_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtManualEntryDelPrice.Leave
        txtManualEntryDelPrice.Text = Fmt_T_B(txtManualEntryDelPrice)
        LoadTheGrid()
        ManualRecalc()
    End Sub
    Private Sub cboDeliveryDistance_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboDeliveryDistance.TextChanged
        With Me.cboDeliveryDistance
            Me.txtManualEntryDelPrice.Text = FormatCurrency(UnFormat(GetToken(.Text, " ", "D")))
        End With
    End Sub

    Private Sub cboDeliveryDistance_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboDeliveryDistance.SelectedIndexChanged
        With Me.cboDeliveryDistance
            Me.txtManualEntryDelPrice.Text = FormatCurrency(UnFormat(GetToken(.Text, " ", "D")))
        End With
        LoadTheGrid()
        ManualRecalc()

    End Sub

    Private Sub chkManualDepositOverride_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkManualDepositOverride.CheckedChanged
        LoadTheGrid()
        ManualRecalc()
    End Sub


    Public Sub cmdPrintContract_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        ' Print the checkout bill
        If iInvID = 0 Then
            Me.GetInvoiceID()
        End If
        PrintInvoice(iInvID)
        Me.btnCheckout.Enabled = True
    End Sub

    Private Sub txtManualDeposit_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtManualDeposit.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then KeyAscii = 0
        'UPGRADE_ISSUE: Assignment not supported: KeyAscii to a non-zero value Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1058"'
        KeyAscii = CkKeyPressNumeric(KeyAscii, txtManualDeposit)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtManualDeposit_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtManualDeposit.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error Resume Next
        If KeyCode = 13 Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
        If KeyCode = System.Windows.Forms.Keys.Up Then System.Windows.Forms.SendKeys.SendWait("+{TAB}")
        If KeyCode = System.Windows.Forms.Keys.Down Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
    End Sub
    Private Sub txtManualDeposit_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtManualDeposit.Enter
        txtManualDeposit.Text = UnFmt_T_B(txtManualDeposit)
        txtManualDeposit.SelectionStart = 0
        txtManualDeposit.SelectionLength = Len(Trim(txtManualDeposit.Text))
    End Sub
    Private Sub txtManualDeposit_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtManualDeposit.Leave
        txtManualDeposit.Text = Fmt_T_B(txtManualDeposit)
        If UnFormat((txtManualDeposit.Text)) <> 0 Then
            LoadTheGrid()
            ManualRecalc()

        End If
    End Sub

    Private Sub txtPostalCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPostalCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then KeyAscii = 0
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtPostalCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtPostalCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error Resume Next
        If KeyCode = 13 Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
        If KeyCode = System.Windows.Forms.Keys.Up Then System.Windows.Forms.SendKeys.SendWait("+{TAB}")
        If KeyCode = System.Windows.Forms.Keys.Down Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
    End Sub
    Private Sub txtPostalCode_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPostalCode.Enter
        txtPostalCode.SelectionStart = 0
        txtPostalCode.SelectionLength = Len(Trim(txtPostalCode.Text))
    End Sub

    Private Sub txtState_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtState.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then KeyAscii = 0
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtState_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtState.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error Resume Next
        If KeyCode = 13 Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
        If KeyCode = System.Windows.Forms.Keys.Up Then System.Windows.Forms.SendKeys.SendWait("+{TAB}")
        If KeyCode = System.Windows.Forms.Keys.Down Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
    End Sub
    Private Sub txtState_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtState.Enter
        txtState.SelectionStart = 0
        txtState.SelectionLength = Len(Trim(txtState.Text))
    End Sub


    Private Sub txtCity_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCity.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then KeyAscii = 0
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtCity_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCity.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error Resume Next
        If KeyCode = 13 Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
        If KeyCode = System.Windows.Forms.Keys.Up Then System.Windows.Forms.SendKeys.SendWait("+{TAB}")
        If KeyCode = System.Windows.Forms.Keys.Down Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
    End Sub
    Private Sub txtCity_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCity.Enter
        txtCity.SelectionStart = 0
        txtCity.SelectionLength = Len(Trim(txtCity.Text))
    End Sub


    Private Sub txtBillingAddress1_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtBillingAddress1.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then KeyAscii = 0
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtBillingAddress1_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtBillingAddress1.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error Resume Next
        If KeyCode = 13 Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
        If KeyCode = System.Windows.Forms.Keys.Up Then System.Windows.Forms.SendKeys.SendWait("+{TAB}")
        If KeyCode = System.Windows.Forms.Keys.Down Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
    End Sub
    Private Sub txtBillingAddress1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBillingAddress1.Enter
        txtBillingAddress1.SelectionStart = 0
        txtBillingAddress1.SelectionLength = Len(Trim(txtBillingAddress1.Text))
    End Sub

    Private Sub txtCompanyName_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCompanyName.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then KeyAscii = 0
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtCompanyName_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCompanyName.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error Resume Next
        If KeyCode = 13 Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
        If KeyCode = System.Windows.Forms.Keys.Up Then System.Windows.Forms.SendKeys.SendWait("+{TAB}")
        If KeyCode = System.Windows.Forms.Keys.Down Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
        ClearCustomerBoxes()
    End Sub
    Private Sub ClearCustomerBoxes()
        With Me
            .txtCustomerID.Text = String.Empty
            .txtContactName.Text = String.Empty
            .txtBillingAddress1.Text = String.Empty
            .txtBillingAddress2.Text = String.Empty
            .txtCity.Text = String.Empty
            .txtState.Text = String.Empty
            .txtPhone.Text = String.Empty
            .txtPostalCode.Text = String.Empty
            .txtTaxID.Text = String.Empty
        End With
    End Sub
    Private Sub txtCompanyName_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCompanyName.Enter
        txtCompanyName.SelectionStart = 0
        txtCompanyName.SelectionLength = Len(Trim(txtCompanyName.Text))
    End Sub
    Private Sub btnManualRecalc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnManualRecalc.Click
        ManualRecalc()
    End Sub

    Private Sub SSOleDBCombo1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles SSOleDBCombo1.SelectedIndexChanged
        LoadCustomerBoxes(SSOleDBCombo1.Text)
    End Sub

    Private Sub txtAmtPaid_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAmtPaid.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtAmtPaid_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAmtPaid.KeyDown
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
        If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
        If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
    End Sub
    Private Sub txtAmtPaid_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAmtPaid.Enter
        txtAmtPaid.SelectionStart = 0
        txtAmtPaid.SelectionLength = txtAmtPaid.Text.Trim.Length
    End Sub

    Private Sub txtAmtPaid_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAmtPaid.Leave
        With Me
            .txtBalDue.Text = FormatCurrency(UnFormat(.txtTotal.Text) - UnFormat(.txtAmtPaid.Text))
            .txtAmtPaid.Text = FormatCurrency(.txtAmtPaid.Text)
        End With
    End Sub
    Private Sub txtPONbr_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPONbr.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtPONbr_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPONbr.KeyDown
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
        If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
        If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
    End Sub
    Private Sub txtPONbr_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPONbr.Enter
        txtPONbr.SelectionStart = 0
        txtPONbr.SelectionLength = txtPONbr.Text.Trim.Length
    End Sub
    Private Sub txtContactName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtContactName.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtContactName_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtContactName.KeyDown
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
        If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
        If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
    End Sub
    Private Sub txtContactName_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtContactName.Enter
        txtContactName.SelectionStart = 0
        txtContactName.SelectionLength = txtContactName.Text.Trim.Length
    End Sub
    Private Sub txtShipAddress1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtShipAddress1.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtShipAddress1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtShipAddress1.KeyDown
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
        If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
        If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
    End Sub
    Private Sub txtShipAddress1_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtShipAddress1.Enter
        txtShipAddress1.SelectionStart = 0
        txtShipAddress1.SelectionLength = txtShipAddress1.Text.Trim.Length
    End Sub
    Private Sub txtShipState_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtShipState.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtShipState_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtShipState.KeyDown
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
        If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
        If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
    End Sub
    Private Sub txtShipState_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtShipState.Enter
        txtShipState.SelectionStart = 0
        txtShipState.SelectionLength = txtShipState.Text.Trim.Length
    End Sub
    Private Sub txtShipCity_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtShipCity.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtShipCity_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtShipCity.KeyDown
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
        If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
        If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
    End Sub
    Private Sub txtShipCity_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtShipCity.Enter
        txtShipCity.SelectionStart = 0
        txtShipCity.SelectionLength = txtShipCity.Text.Trim.Length
    End Sub
    Private Sub txtShipZip_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtShipZip.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtShipZip_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtShipZip.KeyDown
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
        If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
        If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
    End Sub
    Private Sub txtShipZip_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtShipZip.Enter
        txtShipZip.SelectionStart = 0
        txtShipZip.SelectionLength = txtShipZip.Text.Trim.Length
    End Sub
    Private Sub txtShipToCustomer_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtShipToCustomer.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtShipToCustomer_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtShipToCustomer.KeyDown
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
        If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
        If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
    End Sub
    Private Sub txtShipToCustomer_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtShipToCustomer.Enter
        txtShipToCustomer.SelectionStart = 0
        txtShipToCustomer.SelectionLength = txtShipToCustomer.Text.Trim.Length
    End Sub

    Private Sub chkShipBillSame_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShipBillSame.CheckedChanged
        With Me
            If .chkShipBillSame.Checked Then
                .txtShipToCustomer.Text = .txtCompanyName.Text
                .txtShipAddress1.Text = .txtBillingAddress1.Text
                .txtShipCity.Text = .txtCity.Text
                .txtShipState.Text = .txtState.Text
                .txtShipZip.Text = .txtPostalCode.Text
            Else
                .txtShipToCustomer.Text = ""
                .txtShipAddress1.Text = ""
                .txtShipCity.Text = ""
                .txtShipState.Text = ""
                .txtShipZip.Text = ""
            End If
        End With
    End Sub

    Private Sub txtBalDue_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBalDue.Enter
        With Me
            .txtBalDue.Text = FormatCurrency(UnFormat(.txtTotal.Text) - UnFormat(.txtAmtPaid.Text))
        End With
    End Sub

    Private Sub optBillTo_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) _
       Handles optBillTo.CheckedChanged
        If optBillTo.Checked Then
            'Me.txtBalDue.Text = Me.txtAmtPaid.Text
            'Me.txtAmtPaid.Text = FormatCurrency(0)
            MovePaidToDue()
            ManualRecalc()
        End If
    End Sub

    Private Sub optLeftBlankCheck_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optLeftBlankCheck.CheckedChanged
        If Me.optLeftBlankCheck.Checked Then
            Me.txtManualDeposit.Text = FormatCurrency(0)
            Me.chkManualDepositOverride.Checked = True
            DoEvents()
            MovePaidToDue()
            ManualRecalc()
        End If
    End Sub
    Private Sub optCash_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles optCash.CheckedChanged
        If Me.optCash.Checked Then
            MoveDueToPaid()
            ManualRecalc()
        End If
    End Sub

    Private Sub optLeftCardNumber_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optLeftCardNumber.CheckedChanged
        Me.txtManualDeposit.Text = FormatCurrency(0)
        Me.chkManualDepositOverride.Checked = True
        DoEvents()
        MovePaidToDue()
        ManualRecalc()
    End Sub

    Private Sub optPaidByCheck_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optPaidByCheck.CheckedChanged
        If Me.optPaidByCheck.Checked Then
            MoveDueToPaid()
            ManualRecalc()
        End If
    End Sub
    Private Sub optPaidByCreditCard_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles optPaidByCreditCard.CheckedChanged
        If Me.optPaidByCreditCard.Checked Then
            MoveDueToPaid()
            ManualRecalc()
        End If
    End Sub

    Private Sub btnCheckout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCheckout.Click
        CheckOutHandler()
    End Sub

    Private Sub dbgShoppingList_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dbgShoppingList.MouseUp
        Dim pt = New Point(e.X, e.Y)
        Dim hti As DataGrid.HitTestInfo = Me.dbgShoppingList.HitTest(pt)
        Try
            Me.dbgShoppingList.Select(hti.Row)
            iHitRow = hti.Row
        Catch
        End Try
    End Sub

    Private Sub frmCustomers_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Dim sMsg As String
        Dim iRV As Short
        Dim oRES As CTransaction
        Dim serr As String

        If Not mbClosing Then
            sMsg = "Do you want unload the held items from the" & Chr(10)
            sMsg = sMsg & "shopping cart?"
            iRV = MsgBox(sMsg, 35, "Confirm Clearing Hold On Equipment")

            If iRV = DialogResult.Yes Then
                ' Yes Code goes here
                oRES = New CTransaction()
                Call oRES.RemoveTempReservation(True)
                Dim SQL As String = ""
                SQL &= "delete from TempItems "
                Dim iRows As Int16 = oDA.SendActionSql(SQL, ConnectString, serr)
                modMain.ReloadRentalGrid = True
                Exit Sub
                System.Windows.Forms.Application.DoEvents()
            ElseIf iRV = DialogResult.No Then
                ' No code goes here
                Exit Sub
                System.Windows.Forms.Application.DoEvents()
            Else
                ' Cancel code goes here
            End If
        End If
    End Sub

    Private Sub txtCheckNumber_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCheckNumber.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtCheckNumber_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCheckNumber.KeyDown
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
        If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
        If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
    End Sub
    Private Sub txtCheckNumber_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCheckNumber.Enter
        txtCheckNumber.SelectionStart = 0
        txtCheckNumber.SelectionLength = txtCheckNumber.Text.Trim.Length
    End Sub

    Private Sub btnPrintInvoice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrintInvoice.Click
        PrintHandler()
    End Sub
    Private Sub PrintHandler()
        If Not Me.ValidateRequiredFields() Then Exit Sub

        PrintInvoice(iInvID)
        Me.mbPrinted = True
        Me.btnCheckout.Enabled = True
    End Sub
    Private Sub mnuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExit.Click
        CancelHandler()
    End Sub

    Private Sub mnuPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPrint.Click
        PrintHandler()
    End Sub

    Private Sub btnOtherCharges_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOtherCharges.Click
        Dim oFrm As New frmMiscCharges(dtFC, True)
        oFrm.ShowDialog()
        oCG.BindDataTableToGrid(dtFC, Me.dbgShoppingList)
        oCG.DisableAddNew(Me.dbgShoppingList, Me)
        LoadTheGrid()
        ManualRecalc()
    End Sub

    Private Sub txtTaxID_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTaxID.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub

    Private Sub txtTaxID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTaxID.KeyDown
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
        If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
        If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
    End Sub

    Private Sub txtTaxID_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTaxID.Enter
        txtTaxID.SelectionStart = 0
        txtTaxID.SelectionLength = txtTaxID.Text.Trim.Length
    End Sub

    Private Sub txtTaxID_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTaxID.Leave
        If txtTaxID.Text.Trim.Length > 0 Then
            Me.chkChargeTax.Checked = False
        Else
            Me.chkChargeTax.Checked = True
        End If
    End Sub

    Private Sub mnuDeleteSelectedRow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDeleteSelectedRow.Click
        If iHitRow > 0 Then
            If MsgBox("Are you sure you want to delete the selected grid row?", MsgBoxStyle.Question) = MsgBoxResult.Yes Then
                Me.dtFC.Rows(iHitRow).Delete()
                Me.LoadTheGrid()
                Me.ManualRecalc()
            End If
        End If
    End Sub

    Private Sub txtNotes_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNotes.Enter
        Me.ignoreKeyPreview = True
    End Sub

    Private Sub txtNotes_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNotes.Leave
        Me.ignoreKeyPreview = False
    End Sub
    Private Sub txtCardId_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCardId.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtCardId_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCardId.KeyDown
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
        If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
        If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
    End Sub
    Private Sub txtCardId_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCardId.Enter
        txtCardId.SelectionStart = 0
        txtCardId.SelectionLength = txtCardId.Text.Trim.Length
    End Sub
    Private Sub textDriversLicence_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles textDriversLicence.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub textDriversLicence_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles textDriversLicence.KeyDown
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
        If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
        If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
    End Sub
    Private Sub textDriversLicense_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles textDriversLicence.Enter
        textDriversLicence.SelectionStart = 0
        textDriversLicence.SelectionLength = txtCardId.Text.Trim.Length
    End Sub
    Private Sub cbEmployees_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbEmployees.SelectedIndexChanged
        If Me.cbEmployees.Text.Trim.Length > 0 Then
            SaveSetting(RENTALPRO, SETTINGS, "CKOUTEMP", Me.cbEmployees.Text.Trim)
        End If
    End Sub

    Private Sub chkCustomerPickup_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCustomerPickup.CheckedChanged
        With Me.chkCustomerPickup
            If .Checked Then
                Me.txtShipAddress1.Text = "N/A"
                Me.txtShipCity.Text = "N/A"
                Me.txtShipState.Text = "NA"
                Me.txtShipZip.Text = "NA"
                Me.txtShipToCustomer.Text = "Customer Pickup"
                Me.txtContactName.Text = String.Empty
            Else
                Me.txtShipAddress1.Text = String.Empty
                Me.txtShipCity.Text = String.Empty
                Me.txtShipState.Text = String.Empty
                Me.txtShipZip.Text = String.Empty
                Me.txtShipToCustomer.Text = String.Empty
                Me.txtContactName.Text = "None"
            End If
        End With
    End Sub

    Private Sub mnuAddCustomer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddCustomer.Click
        Dim sTxt As String = ""
        sTxt &= "To add a customer, do the following:" & Chr(13) & Chr(10)
        sTxt &= "1. Enter all of the customer data that you know into the "
        sTxt &= "Billl To customer information boxes." & Chr(13) & Chr(10)
        sTxt &= "2. You mus enter the minimum of  customer or company name, "
        sTxt &= "phone number and contact name." & Chr(13) & Chr(10)
        sTxt &= "3. Click the Add Button.  " & Chr(13) & Chr(10)
        sTxt &= "4. The customer will be added to the database and you may "
        sTxt &= "proceed with checkout without reselecting the customer." & Chr(13) & Chr(10)

        Dim f As New frmHelp()
        f.CannedMessage = sTxt
        f.ShowDialog()
    End Sub
    Private Sub chkCashCustomer_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCashCustomer.CheckedChanged
        Static busy As Boolean
        If busy Then Exit Sub
        busy = True

        If chkCashCustomer.Checked Then
            Dim sMsg As String
            Dim iRV As Integer
            sMsg = "Selecting Cash Customer means that you do not have" & Chr(10)
            sMsg &= "an existant customer to link this invoice to, or do not " & Chr(10)
            sMsg &= "want this invoice linked to an existing customer." & Chr(10)
            sMsg &= "" & Chr(10)
            sMsg &= "Are you absolutely sure you want to use the cash " & Chr(10)
            sMsg &= "customer option?" & Chr(10)
            sMsg &= "" & Chr(10)
            sMsg &= "Click OK to use Cash Customer or No to cancel the" & Chr(10)
            sMsg &= "use of the Cash Customer." & Chr(10)
            sMsg &= "" & Chr(10)
            iRV = MsgBox(sMsg, CType(308, Microsoft.VisualBasic.MsgBoxStyle), "Confirm Cash Customer")

            If iRV = 6 Then
                ' Yes Code goes here
            Else
                ' No code goes here
                chkCashCustomer.Checked = False
                busy = False
                Exit Sub
            End If
            LoadCustomerBoxes(0)
            Me.chkCustomerPickup.Checked = True
            Me.optBillTo.Enabled = False
            Me.optBillTo.Checked = False
        Else
            Me.optBillTo.Enabled = True
            busy = False
        End If
    End Sub

    Private Sub btnShowPrinters_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShowPrinters.Click
        Dim frm As New frmPrinters
        frm.ShowDialog()
    End Sub
#End Region



End Class