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
Imports VB = Microsoft.VisualBasic
Imports System.Windows.Forms.Application
Public Class frmCheckinNew
    Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
    Public Sub New()
        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()
        oDA = New CDataAccess()
        oCI = New CCheckIn(Me)
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

    Public WithEvents _lblLabels_0 As System.Windows.Forms.Label
    Public WithEvents _lblLabels_1 As System.Windows.Forms.Label
    Public WithEvents _lblLabels_2 As System.Windows.Forms.Label
    Public WithEvents _lblLabels_3 As System.Windows.Forms.Label
    Public WithEvents _lblLabels_4 As System.Windows.Forms.Label
    Public WithEvents _lblLabels_5 As System.Windows.Forms.Label
    Public WithEvents _lblLabels_6 As System.Windows.Forms.Label

    Public WithEvents cmdCancel As System.Windows.Forms.Button
    Public WithEvents cmdPrintContract As System.Windows.Forms.Button

    Friend WithEvents dbgShoppingList As System.Windows.Forms.DataGrid

    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents HelpProvider1 As System.Windows.Forms.HelpProvider
    Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents lblAmtPaid As System.Windows.Forms.Label
    Friend WithEvents lblBalDue As System.Windows.Forms.Label
    Public WithEvents lblContact As System.Windows.Forms.Label

    Friend WithEvents lblDelivery As System.Windows.Forms.Label
    Friend WithEvents lblDeposit As System.Windows.Forms.Label
    Public WithEvents lblFieldLable As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Friend WithEvents lblItemTotal As System.Windows.Forms.Label
    Public WithEvents lblLabels As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Friend WithEvents lblLine As System.Windows.Forms.Label
    Friend WithEvents lblLine2 As System.Windows.Forms.Label
    Friend WithEvents lblPONbr As System.Windows.Forms.Label
    Friend WithEvents lblTax As System.Windows.Forms.Label
    Friend WithEvents lblTotal As System.Windows.Forms.Label
    Friend WithEvents optBillTo As System.Windows.Forms.RadioButton
    Friend WithEvents optPaidByCheck As System.Windows.Forms.RadioButton
    Friend WithEvents optPaidByCreditCard As System.Windows.Forms.RadioButton
    Friend WithEvents txtAmtPaid As System.Windows.Forms.TextBox
    Friend WithEvents txtBalDue As System.Windows.Forms.TextBox
    Public WithEvents txtBillingAddress1 As System.Windows.Forms.TextBox
    Friend WithEvents txtCheckNumber As System.Windows.Forms.TextBox
    Public WithEvents txtCity As System.Windows.Forms.TextBox
    Public WithEvents txtCompanyName As System.Windows.Forms.TextBox
    Public WithEvents txtContactName As System.Windows.Forms.TextBox

    Friend WithEvents txtDelivery As System.Windows.Forms.TextBox
    Friend WithEvents txtDeposit As System.Windows.Forms.TextBox
    Friend WithEvents txtItemTotal As System.Windows.Forms.TextBox
    Friend WithEvents txtPONbr As System.Windows.Forms.TextBox
    Public WithEvents txtPostalCode As System.Windows.Forms.TextBox
    Friend WithEvents txtSalesTax As System.Windows.Forms.TextBox
    Public WithEvents txtShipAddress1 As System.Windows.Forms.TextBox
    Public WithEvents txtShipCity As System.Windows.Forms.TextBox
    Public WithEvents txtShipState As System.Windows.Forms.TextBox
    Public WithEvents txtShipToCustomer As System.Windows.Forms.TextBox
    Public WithEvents txtShipZip As System.Windows.Forms.TextBox
    Public WithEvents txtState As System.Windows.Forms.TextBox
    Friend WithEvents txtTotal As System.Windows.Forms.TextBox
    Friend WithEvents txtCustomerID As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtInvoiceID As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents optCash As System.Windows.Forms.RadioButton
    Friend WithEvents optLeftCardNumber As System.Windows.Forms.RadioButton
    Friend WithEvents optLeftBlankCheck As System.Windows.Forms.RadioButton
    Friend WithEvents mnuContext As System.Windows.Forms.ContextMenu
    Friend WithEvents mnuDaily As System.Windows.Forms.MenuItem
    Friend WithEvents mnuHalfDay As System.Windows.Forms.MenuItem
    Friend WithEvents mnuWeek As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMonth As System.Windows.Forms.MenuItem
    Friend WithEvents mnuWeekEnd As System.Windows.Forms.MenuItem
    Friend WithEvents btnManualRecalc As System.Windows.Forms.Button
    Friend WithEvents txtAmtPaidAtCkIn As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Public WithEvents btnCheckOut As System.Windows.Forms.Button
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents mnuFile As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPrint As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCheckIn As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCancel As System.Windows.Forms.MenuItem
    Friend WithEvents btnOtherCharges As System.Windows.Forms.Button
    Friend WithEvents txtTaxID As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents mnuOptions As System.Windows.Forms.MenuItem
    Friend WithEvents mnuShowPriceTable As System.Windows.Forms.MenuItem
    Friend WithEvents lblNotes As System.Windows.Forms.Label
    Friend WithEvents txtNotes As System.Windows.Forms.TextBox
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuHourly As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddHalfDay As System.Windows.Forms.MenuItem
    Friend WithEvents mnuVoidInvoice As System.Windows.Forms.MenuItem
    Friend WithEvents txtCardId As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Public WithEvents cboDeliveryDistance As System.Windows.Forms.ComboBox
    Public WithEvents _Label1_1 As System.Windows.Forms.Label
    Public WithEvents _Label1_0 As System.Windows.Forms.Label
    Public WithEvents textManualPickup As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents cbEmployees As System.Windows.Forms.ComboBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents lblCkOutDate As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents textDriversLicence As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents lblWrittenBy As System.Windows.Forms.Label
    Public WithEvents txtBillingAddress2 As System.Windows.Forms.TextBox
    Public WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents lblCkInDate As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents dtpCkInDateReset As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents textElapsedTimeHours As System.Windows.Forms.TextBox
    Friend WithEvents chkPrintToFile As System.Windows.Forms.CheckBox
    Friend WithEvents chkSaveCreditCard As System.Windows.Forms.CheckBox
    Friend WithEvents cbExpYr As System.Windows.Forms.ComboBox
    Friend WithEvents cbExpMon As System.Windows.Forms.ComboBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnRerunAutoCalc As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCheckinNew))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.dbgShoppingList = New System.Windows.Forms.DataGrid()
        Me.cmdPrintContract = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.btnCheckOut = New System.Windows.Forms.Button()
        Me.btnOtherCharges = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.mnuContext = New System.Windows.Forms.ContextMenu()
        Me.MenuItem1 = New System.Windows.Forms.MenuItem()
        Me.mnuDaily = New System.Windows.Forms.MenuItem()
        Me.mnuWeek = New System.Windows.Forms.MenuItem()
        Me.mnuWeekEnd = New System.Windows.Forms.MenuItem()
        Me.mnuHalfDay = New System.Windows.Forms.MenuItem()
        Me.mnuMonth = New System.Windows.Forms.MenuItem()
        Me.mnuHourly = New System.Windows.Forms.MenuItem()
        Me.mnuAddHalfDay = New System.Windows.Forms.MenuItem()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.txtBillingAddress2 = New System.Windows.Forms.TextBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.textDriversLicence = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.txtTaxID = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtInvoiceID = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtCustomerID = New System.Windows.Forms.TextBox()
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
        Me.txtAmtPaidAtCkIn = New System.Windows.Forms.TextBox()
        Me.textManualPickup = New System.Windows.Forms.TextBox()
        Me.btnManualRecalc = New System.Windows.Forms.Button()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.MainMenu1 = New System.Windows.Forms.MainMenu(Me.components)
        Me.mnuFile = New System.Windows.Forms.MenuItem()
        Me.mnuPrint = New System.Windows.Forms.MenuItem()
        Me.mnuCheckIn = New System.Windows.Forms.MenuItem()
        Me.mnuCancel = New System.Windows.Forms.MenuItem()
        Me.mnuVoidInvoice = New System.Windows.Forms.MenuItem()
        Me.mnuOptions = New System.Windows.Forms.MenuItem()
        Me.mnuShowPriceTable = New System.Windows.Forms.MenuItem()
        Me.txtNotes = New System.Windows.Forms.TextBox()
        Me.lblNotes = New System.Windows.Forms.Label()
        Me.txtCardId = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.cboDeliveryDistance = New System.Windows.Forms.ComboBox()
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.cbEmployees = New System.Windows.Forms.ComboBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.lblCkOutDate = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.lblWrittenBy = New System.Windows.Forms.Label()
        Me.lblCkInDate = New System.Windows.Forms.Label()
        Me.dtpCkInDateReset = New System.Windows.Forms.DateTimePicker()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.textElapsedTimeHours = New System.Windows.Forms.TextBox()
        Me.btnRerunAutoCalc = New System.Windows.Forms.Button()
        Me.chkPrintToFile = New System.Windows.Forms.CheckBox()
        Me.chkSaveCreditCard = New System.Windows.Forms.CheckBox()
        Me.cbExpYr = New System.Windows.Forms.ComboBox()
        Me.cbExpMon = New System.Windows.Forms.ComboBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        CType(Me.dbgShoppingList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Frame2.SuspendLayout()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblFieldLable, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblLabels, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'dbgShoppingList
        '
        Me.dbgShoppingList.AllowSorting = False
        Me.dbgShoppingList.AlternatingBackColor = System.Drawing.Color.MediumSeaGreen
        Me.dbgShoppingList.CaptionVisible = False
        Me.dbgShoppingList.DataMember = ""
        Me.dbgShoppingList.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.HelpProvider1.SetHelpString(Me.dbgShoppingList, "To change period or add a half day, right click on the row you desire to change.")
        Me.dbgShoppingList.Location = New System.Drawing.Point(5, 320)
        Me.dbgShoppingList.Name = "dbgShoppingList"
        Me.HelpProvider1.SetShowHelp(Me.dbgShoppingList, True)
        Me.dbgShoppingList.Size = New System.Drawing.Size(934, 231)
        Me.dbgShoppingList.TabIndex = 44
        Me.ToolTip1.SetToolTip(Me.dbgShoppingList, "Right click in left margin of desired row to change period or number")
        '
        'cmdPrintContract
        '
        Me.cmdPrintContract.BackColor = System.Drawing.SystemColors.Control
        Me.cmdPrintContract.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdPrintContract.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPrintContract.ForeColor = System.Drawing.SystemColors.ControlText
        Me.HelpProvider1.SetHelpString(Me.cmdPrintContract, "Print contract and put on rent.")
        Me.cmdPrintContract.Location = New System.Drawing.Point(456, 750)
        Me.cmdPrintContract.Name = "cmdPrintContract"
        Me.cmdPrintContract.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.HelpProvider1.SetShowHelp(Me.cmdPrintContract, True)
        Me.cmdPrintContract.Size = New System.Drawing.Size(179, 32)
        Me.cmdPrintContract.TabIndex = 0
        Me.cmdPrintContract.Text = "&Print Contract"
        Me.ToolTip1.SetToolTip(Me.cmdPrintContract, "Print or preview the check in invoice")
        Me.cmdPrintContract.UseVisualStyleBackColor = False
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.HelpProvider1.SetHelpString(Me.cmdCancel, "Cancel the checkout process.")
        Me.cmdCancel.Location = New System.Drawing.Point(456, 886)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.HelpProvider1.SetShowHelp(Me.cmdCancel, True)
        Me.cmdCancel.Size = New System.Drawing.Size(179, 32)
        Me.cmdCancel.TabIndex = 130
        Me.cmdCancel.Text = "&Cancel"
        Me.ToolTip1.SetToolTip(Me.cmdCancel, "Cancel the check-in process")
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'btnCheckOut
        '
        Me.btnCheckOut.BackColor = System.Drawing.SystemColors.Control
        Me.btnCheckOut.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnCheckOut.Enabled = False
        Me.btnCheckOut.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCheckOut.ForeColor = System.Drawing.SystemColors.ControlText
        Me.HelpProvider1.SetHelpString(Me.btnCheckOut, "Print contract and put on rent.")
        Me.btnCheckOut.Location = New System.Drawing.Point(456, 783)
        Me.btnCheckOut.Name = "btnCheckOut"
        Me.btnCheckOut.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.HelpProvider1.SetShowHelp(Me.btnCheckOut, True)
        Me.btnCheckOut.Size = New System.Drawing.Size(179, 33)
        Me.btnCheckOut.TabIndex = 135
        Me.btnCheckOut.Text = "Check &In"
        Me.ToolTip1.SetToolTip(Me.btnCheckOut, "Complete the check-in  process")
        Me.btnCheckOut.UseVisualStyleBackColor = False
        '
        'btnOtherCharges
        '
        Me.btnOtherCharges.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOtherCharges.Location = New System.Drawing.Point(456, 816)
        Me.btnOtherCharges.Name = "btnOtherCharges"
        Me.btnOtherCharges.Size = New System.Drawing.Size(179, 32)
        Me.btnOtherCharges.TabIndex = 136
        Me.btnOtherCharges.Text = "&Labor, Fuel, Supp"
        Me.ToolTip1.SetToolTip(Me.btnOtherCharges, "Add Labor, Fuel, or Misc items to invoice")
        '
        'btnDelete
        '
        Me.btnDelete.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDelete.Location = New System.Drawing.Point(770, 564)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(144, 34)
        Me.btnDelete.TabIndex = 155
        Me.btnDelete.Text = "&Delete Row"
        Me.ToolTip1.SetToolTip(Me.btnDelete, "Delete the selected grid item")
        '
        'mnuContext
        '
        Me.mnuContext.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.mnuAddHalfDay})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuDaily, Me.mnuWeek, Me.mnuWeekEnd, Me.mnuHalfDay, Me.mnuMonth, Me.mnuHourly})
        Me.MenuItem1.Text = "Change Period"
        '
        'mnuDaily
        '
        Me.mnuDaily.Index = 0
        Me.mnuDaily.Text = "Daily"
        '
        'mnuWeek
        '
        Me.mnuWeek.Index = 1
        Me.mnuWeek.Text = "Weekly"
        '
        'mnuWeekEnd
        '
        Me.mnuWeekEnd.Index = 2
        Me.mnuWeekEnd.Text = "Week End"
        '
        'mnuHalfDay
        '
        Me.mnuHalfDay.Index = 3
        Me.mnuHalfDay.Text = "Half Day"
        '
        'mnuMonth
        '
        Me.mnuMonth.Index = 4
        Me.mnuMonth.Text = "Monthly"
        '
        'mnuHourly
        '
        Me.mnuHourly.Index = 5
        Me.mnuHourly.Text = "Hourly"
        '
        'mnuAddHalfDay
        '
        Me.mnuAddHalfDay.Index = 1
        Me.mnuAddHalfDay.Text = "Add Half Day"
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.txtBillingAddress2)
        Me.Frame2.Controls.Add(Me.Label18)
        Me.Frame2.Controls.Add(Me.textDriversLicence)
        Me.Frame2.Controls.Add(Me.Label15)
        Me.Frame2.Controls.Add(Me.txtTaxID)
        Me.Frame2.Controls.Add(Me.Label9)
        Me.Frame2.Controls.Add(Me.Label7)
        Me.Frame2.Controls.Add(Me.txtInvoiceID)
        Me.Frame2.Controls.Add(Me.Label6)
        Me.Frame2.Controls.Add(Me.txtCustomerID)
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
        Me.Frame2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(5, 4)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(942, 244)
        Me.Frame2.TabIndex = 16
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Customer"
        '
        'txtBillingAddress2
        '
        Me.txtBillingAddress2.AcceptsReturn = True
        Me.txtBillingAddress2.BackColor = System.Drawing.SystemColors.Window
        Me.txtBillingAddress2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtBillingAddress2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBillingAddress2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtBillingAddress2.Location = New System.Drawing.Point(144, 77)
        Me.txtBillingAddress2.MaxLength = 0
        Me.txtBillingAddress2.Name = "txtBillingAddress2"
        Me.txtBillingAddress2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtBillingAddress2.Size = New System.Drawing.Size(418, 26)
        Me.txtBillingAddress2.TabIndex = 63
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.BackColor = System.Drawing.SystemColors.Control
        Me.Label18.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label18.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label18.Location = New System.Drawing.Point(6, 77)
        Me.Label18.Name = "Label18"
        Me.Label18.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label18.Size = New System.Drawing.Size(118, 18)
        Me.Label18.TabIndex = 64
        Me.Label18.Text = "BillingAddress2"
        '
        'textDriversLicence
        '
        Me.textDriversLicence.Location = New System.Drawing.Point(397, 208)
        Me.textDriversLicence.Name = "textDriversLicence"
        Me.textDriversLicence.Size = New System.Drawing.Size(166, 26)
        Me.textDriversLicence.TabIndex = 62
        Me.textDriversLicence.Tag = "(No Auto Formatting)"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(333, 216)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(38, 18)
        Me.Label15.TabIndex = 61
        Me.Label15.Text = "DL#"
        '
        'txtTaxID
        '
        Me.txtTaxID.Location = New System.Drawing.Point(141, 208)
        Me.txtTaxID.Name = "txtTaxID"
        Me.txtTaxID.Size = New System.Drawing.Size(181, 26)
        Me.txtTaxID.TabIndex = 60
        Me.txtTaxID.Tag = "(No Auto Formatting)"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(72, 213)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(50, 18)
        Me.Label9.TabIndex = 59
        Me.Label9.Text = "Tax ID"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(581, 208)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(45, 18)
        Me.Label7.TabIndex = 58
        Me.Label7.Text = "Inv ID"
        '
        'txtInvoiceID
        '
        Me.txtInvoiceID.Location = New System.Drawing.Point(638, 208)
        Me.txtInvoiceID.Name = "txtInvoiceID"
        Me.txtInvoiceID.Size = New System.Drawing.Size(218, 26)
        Me.txtInvoiceID.TabIndex = 57
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(568, 181)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(59, 18)
        Me.Label6.TabIndex = 56
        Me.Label6.Text = "Cust ID"
        '
        'txtCustomerID
        '
        Me.txtCustomerID.Location = New System.Drawing.Point(638, 178)
        Me.txtCustomerID.Name = "txtCustomerID"
        Me.txtCustomerID.Size = New System.Drawing.Size(218, 26)
        Me.txtCustomerID.TabIndex = 55
        '
        'lblLine
        '
        Me.lblLine.BackColor = System.Drawing.Color.Gray
        Me.lblLine.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLine.Location = New System.Drawing.Point(2, 113)
        Me.lblLine.Name = "lblLine"
        Me.lblLine.Size = New System.Drawing.Size(937, 1)
        Me.lblLine.TabIndex = 53
        '
        'txtPONbr
        '
        Me.txtPONbr.Location = New System.Drawing.Point(397, 178)
        Me.txtPONbr.Name = "txtPONbr"
        Me.txtPONbr.Size = New System.Drawing.Size(166, 26)
        Me.txtPONbr.TabIndex = 52
        Me.txtPONbr.Tag = "(No Auto Formatting)"
        '
        'lblPONbr
        '
        Me.lblPONbr.AutoSize = True
        Me.lblPONbr.Location = New System.Drawing.Point(333, 181)
        Me.lblPONbr.Name = "lblPONbr"
        Me.lblPONbr.Size = New System.Drawing.Size(44, 18)
        Me.lblPONbr.TabIndex = 51
        Me.lblPONbr.Text = "PO #"
        '
        'txtShipZip
        '
        Me.txtShipZip.AcceptsReturn = True
        Me.txtShipZip.BackColor = System.Drawing.SystemColors.Window
        Me.txtShipZip.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtShipZip.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtShipZip.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtShipZip.Location = New System.Drawing.Point(734, 151)
        Me.txtShipZip.MaxLength = 0
        Me.txtShipZip.Name = "txtShipZip"
        Me.txtShipZip.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtShipZip.Size = New System.Drawing.Size(146, 26)
        Me.txtShipZip.TabIndex = 47
        Me.txtShipZip.Tag = "(No Auto Formatting)"
        '
        'txtShipState
        '
        Me.txtShipState.AcceptsReturn = True
        Me.txtShipState.BackColor = System.Drawing.SystemColors.Window
        Me.txtShipState.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtShipState.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtShipState.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtShipState.Location = New System.Drawing.Point(638, 151)
        Me.txtShipState.MaxLength = 2
        Me.txtShipState.Name = "txtShipState"
        Me.txtShipState.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtShipState.Size = New System.Drawing.Size(44, 26)
        Me.txtShipState.TabIndex = 46
        Me.txtShipState.Tag = "(No Auto Formatting)"
        '
        'txtShipCity
        '
        Me.txtShipCity.AcceptsReturn = True
        Me.txtShipCity.BackColor = System.Drawing.SystemColors.Window
        Me.txtShipCity.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtShipCity.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtShipCity.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtShipCity.Location = New System.Drawing.Point(638, 121)
        Me.txtShipCity.MaxLength = 0
        Me.txtShipCity.Name = "txtShipCity"
        Me.txtShipCity.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtShipCity.Size = New System.Drawing.Size(277, 26)
        Me.txtShipCity.TabIndex = 45
        Me.txtShipCity.Tag = "(No Auto Formatting)"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(696, 153)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(34, 18)
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
        Me.Label4.Location = New System.Drawing.Point(578, 153)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(49, 18)
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
        Me.Label5.Location = New System.Drawing.Point(578, 126)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(39, 18)
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
        Me.txtContactName.Location = New System.Drawing.Point(141, 178)
        Me.txtContactName.MaxLength = 0
        Me.txtContactName.Name = "txtContactName"
        Me.txtContactName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtContactName.Size = New System.Drawing.Size(181, 26)
        Me.txtContactName.TabIndex = 38
        Me.txtContactName.Tag = "(No Auto Formatting)"
        '
        'txtPostalCode
        '
        Me.txtPostalCode.AcceptsReturn = True
        Me.txtPostalCode.BackColor = System.Drawing.SystemColors.Window
        Me.txtPostalCode.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPostalCode.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPostalCode.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPostalCode.Location = New System.Drawing.Point(736, 47)
        Me.txtPostalCode.MaxLength = 0
        Me.txtPostalCode.Name = "txtPostalCode"
        Me.txtPostalCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPostalCode.Size = New System.Drawing.Size(146, 26)
        Me.txtPostalCode.TabIndex = 7
        '
        'txtState
        '
        Me.txtState.AcceptsReturn = True
        Me.txtState.BackColor = System.Drawing.SystemColors.Window
        Me.txtState.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtState.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtState.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtState.Location = New System.Drawing.Point(640, 47)
        Me.txtState.MaxLength = 2
        Me.txtState.Name = "txtState"
        Me.txtState.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtState.Size = New System.Drawing.Size(43, 26)
        Me.txtState.TabIndex = 6
        '
        'txtShipAddress1
        '
        Me.txtShipAddress1.AcceptsReturn = True
        Me.txtShipAddress1.BackColor = System.Drawing.SystemColors.Window
        Me.txtShipAddress1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtShipAddress1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtShipAddress1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtShipAddress1.Location = New System.Drawing.Point(141, 151)
        Me.txtShipAddress1.MaxLength = 0
        Me.txtShipAddress1.Name = "txtShipAddress1"
        Me.txtShipAddress1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtShipAddress1.Size = New System.Drawing.Size(417, 26)
        Me.txtShipAddress1.TabIndex = 4
        Me.txtShipAddress1.Tag = "(No Auto Formatting)"
        '
        'txtCity
        '
        Me.txtCity.AcceptsReturn = True
        Me.txtCity.BackColor = System.Drawing.SystemColors.Window
        Me.txtCity.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCity.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCity.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCity.Location = New System.Drawing.Point(640, 18)
        Me.txtCity.MaxLength = 0
        Me.txtCity.Name = "txtCity"
        Me.txtCity.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCity.Size = New System.Drawing.Size(277, 26)
        Me.txtCity.TabIndex = 5
        '
        'txtShipToCustomer
        '
        Me.txtShipToCustomer.AcceptsReturn = True
        Me.txtShipToCustomer.BackColor = System.Drawing.SystemColors.Window
        Me.txtShipToCustomer.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtShipToCustomer.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtShipToCustomer.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtShipToCustomer.Location = New System.Drawing.Point(141, 121)
        Me.txtShipToCustomer.MaxLength = 0
        Me.txtShipToCustomer.Name = "txtShipToCustomer"
        Me.txtShipToCustomer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtShipToCustomer.Size = New System.Drawing.Size(417, 26)
        Me.txtShipToCustomer.TabIndex = 3
        Me.txtShipToCustomer.Tag = "(No Auto Formatting)"
        '
        'txtBillingAddress1
        '
        Me.txtBillingAddress1.AcceptsReturn = True
        Me.txtBillingAddress1.BackColor = System.Drawing.SystemColors.Window
        Me.txtBillingAddress1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtBillingAddress1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBillingAddress1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtBillingAddress1.Location = New System.Drawing.Point(144, 47)
        Me.txtBillingAddress1.MaxLength = 0
        Me.txtBillingAddress1.Name = "txtBillingAddress1"
        Me.txtBillingAddress1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtBillingAddress1.Size = New System.Drawing.Size(418, 26)
        Me.txtBillingAddress1.TabIndex = 2
        '
        'txtCompanyName
        '
        Me.txtCompanyName.AcceptsReturn = True
        Me.txtCompanyName.BackColor = System.Drawing.SystemColors.Window
        Me.txtCompanyName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCompanyName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCompanyName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCompanyName.Location = New System.Drawing.Point(144, 18)
        Me.txtCompanyName.MaxLength = 0
        Me.txtCompanyName.Name = "txtCompanyName"
        Me.txtCompanyName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCompanyName.Size = New System.Drawing.Size(418, 26)
        Me.txtCompanyName.TabIndex = 1
        '
        'lblContact
        '
        Me.lblContact.AutoSize = True
        Me.lblContact.BackColor = System.Drawing.SystemColors.Control
        Me.lblContact.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblContact.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblContact.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblContact.Location = New System.Drawing.Point(10, 180)
        Me.lblContact.Name = "lblContact"
        Me.lblContact.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblContact.Size = New System.Drawing.Size(112, 18)
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
        Me._lblLabels_6.Location = New System.Drawing.Point(698, 50)
        Me._lblLabels_6.Name = "_lblLabels_6"
        Me._lblLabels_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_6.Size = New System.Drawing.Size(34, 18)
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
        Me._lblLabels_5.Location = New System.Drawing.Point(579, 50)
        Me._lblLabels_5.Name = "_lblLabels_5"
        Me._lblLabels_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_5.Size = New System.Drawing.Size(49, 18)
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
        Me._lblLabels_4.Location = New System.Drawing.Point(579, 22)
        Me._lblLabels_4.Name = "_lblLabels_4"
        Me._lblLabels_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_4.Size = New System.Drawing.Size(39, 18)
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
        Me.lblLabels.SetIndex(Me._lblLabels_3, CType(3, Short))
        Me._lblLabels_3.Location = New System.Drawing.Point(13, 155)
        Me._lblLabels_3.Name = "_lblLabels_3"
        Me._lblLabels_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_3.Size = New System.Drawing.Size(79, 18)
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
        Me.lblLabels.SetIndex(Me._lblLabels_2, CType(2, Short))
        Me._lblLabels_2.Location = New System.Drawing.Point(13, 126)
        Me._lblLabels_2.Name = "_lblLabels_2"
        Me._lblLabels_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_2.Size = New System.Drawing.Size(64, 18)
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
        Me._lblLabels_1.Location = New System.Drawing.Point(8, 48)
        Me._lblLabels_1.Name = "_lblLabels_1"
        Me._lblLabels_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_1.Size = New System.Drawing.Size(118, 18)
        Me._lblLabels_1.TabIndex = 18
        Me._lblLabels_1.Text = "BillingAddress1"
        '
        '_lblLabels_0
        '
        Me._lblLabels_0.AutoSize = True
        Me._lblLabels_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblLabels_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_0, CType(0, Short))
        Me._lblLabels_0.Location = New System.Drawing.Point(8, 23)
        Me._lblLabels_0.Name = "_lblLabels_0"
        Me._lblLabels_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_0.Size = New System.Drawing.Size(121, 18)
        Me._lblLabels_0.TabIndex = 17
        Me._lblLabels_0.Text = "CompanyName:"
        '
        'txtDelivery
        '
        Me.txtDelivery.Location = New System.Drawing.Point(774, 642)
        Me.txtDelivery.Name = "txtDelivery"
        Me.txtDelivery.Size = New System.Drawing.Size(149, 26)
        Me.txtDelivery.TabIndex = 45
        Me.txtDelivery.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblDelivery
        '
        Me.lblDelivery.AutoSize = True
        Me.lblDelivery.Location = New System.Drawing.Point(698, 645)
        Me.lblDelivery.Name = "lblDelivery"
        Me.lblDelivery.Size = New System.Drawing.Size(64, 18)
        Me.lblDelivery.TabIndex = 46
        Me.lblDelivery.Text = "Delivery"
        '
        'lblTax
        '
        Me.lblTax.AutoSize = True
        Me.lblTax.Location = New System.Drawing.Point(683, 712)
        Me.lblTax.Name = "lblTax"
        Me.lblTax.Size = New System.Drawing.Size(75, 18)
        Me.lblTax.TabIndex = 47
        Me.lblTax.Text = "Sales Tax"
        '
        'txtSalesTax
        '
        Me.txtSalesTax.Location = New System.Drawing.Point(774, 707)
        Me.txtSalesTax.Name = "txtSalesTax"
        Me.txtSalesTax.Size = New System.Drawing.Size(149, 26)
        Me.txtSalesTax.TabIndex = 48
        Me.txtSalesTax.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtDeposit
        '
        Me.txtDeposit.Location = New System.Drawing.Point(774, 741)
        Me.txtDeposit.Name = "txtDeposit"
        Me.txtDeposit.Size = New System.Drawing.Size(149, 26)
        Me.txtDeposit.TabIndex = 49
        Me.txtDeposit.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblDeposit
        '
        Me.lblDeposit.AutoSize = True
        Me.lblDeposit.Location = New System.Drawing.Point(699, 744)
        Me.lblDeposit.Name = "lblDeposit"
        Me.lblDeposit.Size = New System.Drawing.Size(63, 18)
        Me.lblDeposit.TabIndex = 50
        Me.lblDeposit.Text = "Deposit"
        '
        'lblTotal
        '
        Me.lblTotal.AutoSize = True
        Me.lblTotal.Location = New System.Drawing.Point(722, 778)
        Me.lblTotal.Name = "lblTotal"
        Me.lblTotal.Size = New System.Drawing.Size(40, 18)
        Me.lblTotal.TabIndex = 51
        Me.lblTotal.Text = "Total"
        Me.lblTotal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtTotal
        '
        Me.txtTotal.Location = New System.Drawing.Point(774, 775)
        Me.txtTotal.Name = "txtTotal"
        Me.txtTotal.Size = New System.Drawing.Size(149, 26)
        Me.txtTotal.TabIndex = 52
        Me.txtTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblItemTotal
        '
        Me.lblItemTotal.AutoSize = True
        Me.lblItemTotal.Location = New System.Drawing.Point(682, 611)
        Me.lblItemTotal.Name = "lblItemTotal"
        Me.lblItemTotal.Size = New System.Drawing.Size(73, 18)
        Me.lblItemTotal.TabIndex = 53
        Me.lblItemTotal.Text = "Item Total"
        '
        'txtItemTotal
        '
        Me.txtItemTotal.Location = New System.Drawing.Point(774, 608)
        Me.txtItemTotal.Name = "txtItemTotal"
        Me.txtItemTotal.Size = New System.Drawing.Size(148, 26)
        Me.txtItemTotal.TabIndex = 0
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
        Me.GroupBox1.Location = New System.Drawing.Point(6, 627)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(404, 110)
        Me.GroupBox1.TabIndex = 55
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Payment Arrangement"
        '
        'optCash
        '
        Me.optCash.Location = New System.Drawing.Point(219, 82)
        Me.optCash.Name = "optCash"
        Me.optCash.Size = New System.Drawing.Size(154, 23)
        Me.optCash.TabIndex = 6
        Me.optCash.Text = "Paid by Cash"
        '
        'optLeftCardNumber
        '
        Me.optLeftCardNumber.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optLeftCardNumber.Location = New System.Drawing.Point(14, 82)
        Me.optLeftCardNumber.Name = "optLeftCardNumber"
        Me.optLeftCardNumber.Size = New System.Drawing.Size(176, 23)
        Me.optLeftCardNumber.TabIndex = 4
        Me.optLeftCardNumber.Text = "Left Card Number"
        '
        'optBillTo
        '
        Me.optBillTo.Location = New System.Drawing.Point(219, 54)
        Me.optBillTo.Name = "optBillTo"
        Me.optBillTo.Size = New System.Drawing.Size(167, 23)
        Me.optBillTo.TabIndex = 3
        Me.optBillTo.Text = "Bill To Customer"
        '
        'optLeftBlankCheck
        '
        Me.optLeftBlankCheck.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optLeftBlankCheck.Location = New System.Drawing.Point(219, 23)
        Me.optLeftBlankCheck.Name = "optLeftBlankCheck"
        Me.optLeftBlankCheck.Size = New System.Drawing.Size(170, 24)
        Me.optLeftBlankCheck.TabIndex = 2
        Me.optLeftBlankCheck.Text = "Left Blank Check"
        '
        'optPaidByCreditCard
        '
        Me.optPaidByCreditCard.Location = New System.Drawing.Point(14, 53)
        Me.optPaidByCreditCard.Name = "optPaidByCreditCard"
        Me.optPaidByCreditCard.Size = New System.Drawing.Size(189, 23)
        Me.optPaidByCreditCard.TabIndex = 1
        Me.optPaidByCreditCard.Text = "Paid by Credit Card"
        '
        'optPaidByCheck
        '
        Me.optPaidByCheck.Location = New System.Drawing.Point(13, 23)
        Me.optPaidByCheck.Name = "optPaidByCheck"
        Me.optPaidByCheck.Size = New System.Drawing.Size(150, 24)
        Me.optPaidByCheck.TabIndex = 0
        Me.optPaidByCheck.Text = "Paid by Check"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(11, 767)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(155, 18)
        Me.Label2.TabIndex = 56
        Me.Label2.Text = "Check/Card  Number"
        '
        'txtCheckNumber
        '
        Me.HelpProvider1.SetHelpString(Me.txtCheckNumber, "Enter check or credit card number if left.")
        Me.txtCheckNumber.Location = New System.Drawing.Point(11, 791)
        Me.txtCheckNumber.Name = "txtCheckNumber"
        Me.HelpProvider1.SetShowHelp(Me.txtCheckNumber, True)
        Me.txtCheckNumber.Size = New System.Drawing.Size(179, 26)
        Me.txtCheckNumber.TabIndex = 57
        '
        'lblAmtPaid
        '
        Me.lblAmtPaid.Location = New System.Drawing.Point(691, 810)
        Me.lblAmtPaid.Name = "lblAmtPaid"
        Me.lblAmtPaid.Size = New System.Drawing.Size(79, 25)
        Me.lblAmtPaid.TabIndex = 58
        Me.lblAmtPaid.Text = "Amt Paid At Rental"
        '
        'txtAmtPaid
        '
        Me.HelpProvider1.SetHelpString(Me.txtAmtPaid, "Enter 0 if no payment made.")
        Me.txtAmtPaid.Location = New System.Drawing.Point(776, 808)
        Me.txtAmtPaid.Name = "txtAmtPaid"
        Me.HelpProvider1.SetShowHelp(Me.txtAmtPaid, True)
        Me.txtAmtPaid.Size = New System.Drawing.Size(147, 26)
        Me.txtAmtPaid.TabIndex = 59
        Me.txtAmtPaid.Tag = "(No Auto Formatting)"
        Me.txtAmtPaid.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtBalDue
        '
        Me.txtBalDue.Location = New System.Drawing.Point(773, 884)
        Me.txtBalDue.Name = "txtBalDue"
        Me.txtBalDue.Size = New System.Drawing.Size(150, 26)
        Me.txtBalDue.TabIndex = 61
        Me.txtBalDue.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblBalDue
        '
        Me.lblBalDue.Location = New System.Drawing.Point(688, 889)
        Me.lblBalDue.Name = "lblBalDue"
        Me.lblBalDue.Size = New System.Drawing.Size(77, 22)
        Me.lblBalDue.TabIndex = 60
        Me.lblBalDue.Text = "Balance Due"
        '
        'lblLine2
        '
        Me.lblLine2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLine2.Location = New System.Drawing.Point(773, 875)
        Me.lblLine2.Name = "lblLine2"
        Me.lblLine2.Size = New System.Drawing.Size(147, 3)
        Me.lblLine2.TabIndex = 62
        '
        'txtAmtPaidAtCkIn
        '
        Me.HelpProvider1.SetHelpString(Me.txtAmtPaidAtCkIn, "Enter 0 if no payment made.")
        Me.txtAmtPaidAtCkIn.Location = New System.Drawing.Point(774, 842)
        Me.txtAmtPaidAtCkIn.Name = "txtAmtPaidAtCkIn"
        Me.HelpProvider1.SetShowHelp(Me.txtAmtPaidAtCkIn, True)
        Me.txtAmtPaidAtCkIn.Size = New System.Drawing.Size(149, 26)
        Me.txtAmtPaidAtCkIn.TabIndex = 133
        Me.txtAmtPaidAtCkIn.Tag = "(No Auto Formatting)"
        Me.txtAmtPaidAtCkIn.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'textManualPickup
        '
        Me.textManualPickup.AcceptsReturn = True
        Me.textManualPickup.BackColor = System.Drawing.SystemColors.Window
        Me.textManualPickup.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.textManualPickup.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.textManualPickup.ForeColor = System.Drawing.SystemColors.WindowText
        Me.HelpProvider1.SetHelpString(Me.textManualPickup, "Enter override delivery charge.")
        Me.textManualPickup.Location = New System.Drawing.Point(774, 675)
        Me.textManualPickup.MaxLength = 0
        Me.textManualPickup.Name = "textManualPickup"
        Me.textManualPickup.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.HelpProvider1.SetShowHelp(Me.textManualPickup, True)
        Me.textManualPickup.Size = New System.Drawing.Size(149, 26)
        Me.textManualPickup.TabIndex = 30
        Me.textManualPickup.Tag = "$#,##0.00;($#,##0.00)"
        Me.textManualPickup.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'btnManualRecalc
        '
        Me.btnManualRecalc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnManualRecalc.Location = New System.Drawing.Point(456, 851)
        Me.btnManualRecalc.Name = "btnManualRecalc"
        Me.btnManualRecalc.Size = New System.Drawing.Size(179, 32)
        Me.btnManualRecalc.TabIndex = 132
        Me.btnManualRecalc.Text = "&Manual Recalc"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(664, 840)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(101, 43)
        Me.Label8.TabIndex = 134
        Me.Label8.Text = "Due/Paid At Check In"
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFile, Me.mnuOptions})
        '
        'mnuFile
        '
        Me.mnuFile.Index = 0
        Me.mnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuPrint, Me.mnuCheckIn, Me.mnuCancel, Me.mnuVoidInvoice})
        Me.mnuFile.Text = "&File"
        '
        'mnuPrint
        '
        Me.mnuPrint.Index = 0
        Me.mnuPrint.Shortcut = System.Windows.Forms.Shortcut.CtrlP
        Me.mnuPrint.Text = "Print"
        '
        'mnuCheckIn
        '
        Me.mnuCheckIn.Index = 1
        Me.mnuCheckIn.Shortcut = System.Windows.Forms.Shortcut.CtrlI
        Me.mnuCheckIn.Text = "Check In"
        '
        'mnuCancel
        '
        Me.mnuCancel.Index = 2
        Me.mnuCancel.Shortcut = System.Windows.Forms.Shortcut.CtrlC
        Me.mnuCancel.Text = "Cancel"
        '
        'mnuVoidInvoice
        '
        Me.mnuVoidInvoice.Index = 3
        Me.mnuVoidInvoice.Text = "&Void Invoice"
        '
        'mnuOptions
        '
        Me.mnuOptions.Index = 1
        Me.mnuOptions.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuShowPriceTable})
        Me.mnuOptions.Text = "&Options"
        '
        'mnuShowPriceTable
        '
        Me.mnuShowPriceTable.Index = 0
        Me.mnuShowPriceTable.Shortcut = System.Windows.Forms.Shortcut.CtrlR
        Me.mnuShowPriceTable.Text = "Show &Rates"
        '
        'txtNotes
        '
        Me.txtNotes.Location = New System.Drawing.Point(8, 852)
        Me.txtNotes.MaxLength = 255
        Me.txtNotes.Multiline = True
        Me.txtNotes.Name = "txtNotes"
        Me.txtNotes.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtNotes.Size = New System.Drawing.Size(430, 64)
        Me.txtNotes.TabIndex = 137
        '
        'lblNotes
        '
        Me.lblNotes.AutoSize = True
        Me.lblNotes.Location = New System.Drawing.Point(8, 826)
        Me.lblNotes.Name = "lblNotes"
        Me.lblNotes.Size = New System.Drawing.Size(101, 18)
        Me.lblNotes.TabIndex = 138
        Me.lblNotes.Text = "Invoice Notes"
        '
        'txtCardId
        '
        Me.txtCardId.Location = New System.Drawing.Point(376, 791)
        Me.txtCardId.MaxLength = 4
        Me.txtCardId.Name = "txtCardId"
        Me.txtCardId.Size = New System.Drawing.Size(64, 26)
        Me.txtCardId.TabIndex = 152
        Me.txtCardId.Tag = "(No Auto Formatting)"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(363, 767)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(62, 18)
        Me.Label10.TabIndex = 151
        Me.Label10.Text = "Card ID"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(286, 767)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(53, 18)
        Me.Label11.TabIndex = 150
        Me.Label11.Text = "Exp Yr"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(194, 767)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(69, 18)
        Me.Label12.TabIndex = 147
        Me.Label12.Text = "Exp Mon"
        '
        'cboDeliveryDistance
        '
        Me.cboDeliveryDistance.BackColor = System.Drawing.SystemColors.Window
        Me.cboDeliveryDistance.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboDeliveryDistance.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboDeliveryDistance.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboDeliveryDistance.Location = New System.Drawing.Point(10, 586)
        Me.cboDeliveryDistance.Name = "cboDeliveryDistance"
        Me.cboDeliveryDistance.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboDeliveryDistance.Size = New System.Drawing.Size(411, 26)
        Me.cboDeliveryDistance.TabIndex = 29
        '
        '_Label1_1
        '
        Me._Label1_1.AutoSize = True
        Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_1.Location = New System.Drawing.Point(710, 675)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.Size = New System.Drawing.Size(56, 18)
        Me._Label1_1.TabIndex = 32
        Me._Label1_1.Text = "Pickup"
        '
        '_Label1_0
        '
        Me._Label1_0.AutoSize = True
        Me._Label1_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_0.Location = New System.Drawing.Point(8, 561)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_0.Size = New System.Drawing.Size(122, 18)
        Me._Label1_0.TabIndex = 31
        Me._Label1_0.Text = "Pickup Distance"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(432, 561)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(142, 18)
        Me.Label13.TabIndex = 158
        Me.Label13.Text = "Check In Employee"
        '
        'cbEmployees
        '
        Me.cbEmployees.Location = New System.Drawing.Point(434, 586)
        Me.cbEmployees.Name = "cbEmployees"
        Me.cbEmployees.Size = New System.Drawing.Size(216, 26)
        Me.cbEmployees.TabIndex = 157
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(251, 259)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(106, 20)
        Me.Label14.TabIndex = 159
        Me.Label14.Text = "Ck Out Time"
        '
        'lblCkOutDate
        '
        Me.lblCkOutDate.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCkOutDate.Location = New System.Drawing.Point(368, 265)
        Me.lblCkOutDate.Name = "lblCkOutDate"
        Me.lblCkOutDate.Size = New System.Drawing.Size(190, 23)
        Me.lblCkOutDate.TabIndex = 160
        Me.lblCkOutDate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(251, 295)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(106, 21)
        Me.Label16.TabIndex = 161
        Me.Label16.Text = "Ck In Time"
        '
        'Label17
        '
        Me.Label17.Location = New System.Drawing.Point(8, 260)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(61, 37)
        Me.Label17.TabIndex = 163
        Me.Label17.Text = "Ck'd Out By"
        '
        'lblWrittenBy
        '
        Me.lblWrittenBy.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblWrittenBy.Location = New System.Drawing.Point(82, 265)
        Me.lblWrittenBy.Name = "lblWrittenBy"
        Me.lblWrittenBy.Size = New System.Drawing.Size(156, 23)
        Me.lblWrittenBy.TabIndex = 164
        Me.lblWrittenBy.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblCkInDate
        '
        Me.lblCkInDate.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCkInDate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCkInDate.Location = New System.Drawing.Point(366, 294)
        Me.lblCkInDate.Name = "lblCkInDate"
        Me.lblCkInDate.Size = New System.Drawing.Size(192, 23)
        Me.lblCkInDate.TabIndex = 162
        Me.lblCkInDate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'dtpCkInDateReset
        '
        Me.dtpCkInDateReset.CustomFormat = "MM/dd/yyyy hh:mm tt"
        Me.dtpCkInDateReset.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpCkInDateReset.Location = New System.Drawing.Point(443, 646)
        Me.dtpCkInDateReset.Name = "dtpCkInDateReset"
        Me.dtpCkInDateReset.Size = New System.Drawing.Size(213, 26)
        Me.dtpCkInDateReset.TabIndex = 165
        '
        'Label19
        '
        Me.Label19.Location = New System.Drawing.Point(438, 626)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(156, 16)
        Me.Label19.TabIndex = 166
        Me.Label19.Text = "Ck In Date Reset"
        '
        'Label20
        '
        Me.Label20.Location = New System.Drawing.Point(571, 265)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(115, 46)
        Me.Label20.TabIndex = 167
        Me.Label20.Text = "Out Period / Hours"
        '
        'textElapsedTimeHours
        '
        Me.textElapsedTimeHours.Location = New System.Drawing.Point(686, 259)
        Me.textElapsedTimeHours.Multiline = True
        Me.textElapsedTimeHours.Name = "textElapsedTimeHours"
        Me.textElapsedTimeHours.Size = New System.Drawing.Size(256, 57)
        Me.textElapsedTimeHours.TabIndex = 168
        '
        'btnRerunAutoCalc
        '
        Me.btnRerunAutoCalc.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRerunAutoCalc.Location = New System.Drawing.Point(456, 716)
        Me.btnRerunAutoCalc.Name = "btnRerunAutoCalc"
        Me.btnRerunAutoCalc.Size = New System.Drawing.Size(179, 32)
        Me.btnRerunAutoCalc.TabIndex = 169
        Me.btnRerunAutoCalc.Text = "ReRun AutoCalc"
        '
        'chkPrintToFile
        '
        Me.chkPrintToFile.AutoSize = True
        Me.chkPrintToFile.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkPrintToFile.ForeColor = System.Drawing.Color.Red
        Me.chkPrintToFile.Location = New System.Drawing.Point(456, 684)
        Me.chkPrintToFile.Name = "chkPrintToFile"
        Me.chkPrintToFile.Size = New System.Drawing.Size(131, 23)
        Me.chkPrintToFile.TabIndex = 170
        Me.chkPrintToFile.Text = "Print to File?"
        Me.chkPrintToFile.UseVisualStyleBackColor = True
        '
        'chkSaveCreditCard
        '
        Me.chkSaveCreditCard.AutoSize = True
        Me.chkSaveCreditCard.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSaveCreditCard.ForeColor = System.Drawing.Color.Red
        Me.chkSaveCreditCard.Location = New System.Drawing.Point(13, 741)
        Me.chkSaveCreditCard.Name = "chkSaveCreditCard"
        Me.chkSaveCreditCard.Size = New System.Drawing.Size(256, 23)
        Me.chkSaveCreditCard.TabIndex = 171
        Me.chkSaveCreditCard.Text = "Save Credit Card and DL Info"
        Me.chkSaveCreditCard.UseVisualStyleBackColor = True
        '
        'cbExpYr
        '
        Me.cbExpYr.Items.AddRange(New Object() {"03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "40", "42", "42", "43", "44", "45", "46", "47", "48,", "49,", "50"})
        Me.cbExpYr.Location = New System.Drawing.Point(293, 788)
        Me.cbExpYr.Name = "cbExpYr"
        Me.cbExpYr.Size = New System.Drawing.Size(77, 26)
        Me.cbExpYr.TabIndex = 173
        '
        'cbExpMon
        '
        Me.cbExpMon.Items.AddRange(New Object() {"01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"})
        Me.cbExpMon.Location = New System.Drawing.Point(200, 788)
        Me.cbExpMon.Name = "cbExpMon"
        Me.cbExpMon.Size = New System.Drawing.Size(77, 26)
        Me.cbExpMon.TabIndex = 172
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.AutoScroll = True
        Me.Panel1.Controls.Add(Me.Frame2)
        Me.Panel1.Controls.Add(Me.lblLine2)
        Me.Panel1.Controls.Add(Me.lblNotes)
        Me.Panel1.Controls.Add(Me.txtBalDue)
        Me.Panel1.Controls.Add(Me.txtNotes)
        Me.Panel1.Controls.Add(Me.lblBalDue)
        Me.Panel1.Controls.Add(Me.cbExpYr)
        Me.Panel1.Controls.Add(Me.cmdCancel)
        Me.Panel1.Controls.Add(Me.Label17)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.cbExpMon)
        Me.Panel1.Controls.Add(Me.txtAmtPaidAtCkIn)
        Me.Panel1.Controls.Add(Me.lblCkOutDate)
        Me.Panel1.Controls.Add(Me.btnManualRecalc)
        Me.Panel1.Controls.Add(Me.Label10)
        Me.Panel1.Controls.Add(Me.chkSaveCreditCard)
        Me.Panel1.Controls.Add(Me.Label11)
        Me.Panel1.Controls.Add(Me.lblCkInDate)
        Me.Panel1.Controls.Add(Me.Label12)
        Me.Panel1.Controls.Add(Me.btnRerunAutoCalc)
        Me.Panel1.Controls.Add(Me.chkPrintToFile)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.txtCardId)
        Me.Panel1.Controls.Add(Me.Label14)
        Me.Panel1.Controls.Add(Me.Label16)
        Me.Panel1.Controls.Add(Me.Label19)
        Me.Panel1.Controls.Add(Me.lblTotal)
        Me.Panel1.Controls.Add(Me.textElapsedTimeHours)
        Me.Panel1.Controls.Add(Me.lblDeposit)
        Me.Panel1.Controls.Add(Me.dtpCkInDateReset)
        Me.Panel1.Controls.Add(Me.lblWrittenBy)
        Me.Panel1.Controls.Add(Me.txtCheckNumber)
        Me.Panel1.Controls.Add(Me.Label13)
        Me.Panel1.Controls.Add(Me.btnOtherCharges)
        Me.Panel1.Controls.Add(Me.lblAmtPaid)
        Me.Panel1.Controls.Add(Me.Label20)
        Me.Panel1.Controls.Add(Me.btnCheckOut)
        Me.Panel1.Controls.Add(Me.dbgShoppingList)
        Me.Panel1.Controls.Add(Me._Label1_0)
        Me.Panel1.Controls.Add(Me.cboDeliveryDistance)
        Me.Panel1.Controls.Add(Me.lblItemTotal)
        Me.Panel1.Controls.Add(Me.txtAmtPaid)
        Me.Panel1.Controls.Add(Me.btnDelete)
        Me.Panel1.Controls.Add(Me.lblTax)
        Me.Panel1.Controls.Add(Me.txtTotal)
        Me.Panel1.Controls.Add(Me.cbEmployees)
        Me.Panel1.Controls.Add(Me.txtDeposit)
        Me.Panel1.Controls.Add(Me.lblDelivery)
        Me.Panel1.Controls.Add(Me.txtItemTotal)
        Me.Panel1.Controls.Add(Me._Label1_1)
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Controls.Add(Me.textManualPickup)
        Me.Panel1.Controls.Add(Me.cmdPrintContract)
        Me.Panel1.Controls.Add(Me.txtDelivery)
        Me.Panel1.Controls.Add(Me.txtSalesTax)
        Me.Panel1.Location = New System.Drawing.Point(13, 1)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(982, 902)
        Me.Panel1.TabIndex = 174
        '
        'frmCheckinNew
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(8, 19)
        Me.ClientSize = New System.Drawing.Size(1002, 904)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.HelpButton = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(141, 141)
        Me.MaximizeBox = False
        Me.Menu = Me.MainMenu1
        Me.MinimizeBox = False
        Me.Name = "frmCheckinNew"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Customer Check-In"
        CType(Me.dbgShoppingList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Frame2.ResumeLayout(False)
        Me.Frame2.PerformLayout()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblFieldLable, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblLabels, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
#End Region

#Region " Module Variables "
    'Public dtFC As DataTable
    Dim mbFormLoading As Boolean = True
    Private oDA As CDataAccess
    Private m_CurrentInvoice As Integer
    Public dtList As New DataTable()
    Private m_bPrintedInvoice As Boolean = False
    Private oCG As New CGrid()
    Public origItemCount As Integer = 0
    Private ignoreKeyPreview As Boolean
    Private oCI As CCheckIn
    Public voidInvoice As Boolean = False
    Private listHitRow As Integer
    Public CheckOutEmployee As String


#End Region

#Region " No Longer Used "
    'Private Sub chkChargeTax_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
    '   If Not mbFormLoading Then ManualRecalc() 'LoadTheGrid()
    'End Sub

#End Region

#Region " No Longer Used "
    'Private Sub chkDepositRequired_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
    '   If Not mbFormLoading Then
    '      If eventSender.Checked Then
    '         ManualRecalc() 'LoadTheGrid()
    '      End If
    '   End If
    'End Sub

#End Region

#Region " No Longer Used "
    'Private Sub LoadTheGrid()
    '   ' for chek in we need to 
    '   ' 1) load the invoice header data
    '   ' 2) load the invoice data into the grid (record type 1)
    '   ' 3) load the total boxes
    '   Dim i As Short
    '   Dim o As CItems
    '   Dim lsLine As String
    '   Dim lcurDeposit As Decimal = 0
    '   Dim lcurDelivery As Decimal = 0
    '   Dim lcurItemTotal As Decimal = 0
    '   Dim lcurPrice As Decimal = 0
    '   Dim lcurTax As Decimal = 0
    '   Dim SQL As String
    '   Dim dt As New DataTable()
    '   Dim dr As DataRow()
    '   Dim lTotal As Decimal
    '   Static bGridloaded As Boolean
    '   Dim decLabor As Decimal

    '   Try
    '      lcurItemTotal = 0

    '      If Not bGridloaded Then
    '         dtFC = New DataTable()
    '         SQL = "select * from tempitems where user_id = '" & UserName & "' order by ItemId"
    '         oDA.SendQuery(SQL, dtFC, ConnectString)
    '      End If

    '      ' loop thru dt to accumulate and calc totals
    '      ' ItemTotal, Delivery, Sales Tax, Deposit, Total
    '      If dtFC.Rows.Count > 0 Then
    '         For i = 0 To dtFC.Rows.Count - 1
    '            With dtFC.Rows(i)
    '               ' now compute the running total
    '               lcurDeposit += .Item("ItemDeposit")
    '               lcurPrice += .Item("ItemExtendedPrice")
    '               'lcurItemTotal += .Item("ItemTotal")
    '            End With
    '         Next i

    '         If Not bGridloaded Then
    '            Me.dbgShoppingList.DataSource = dtFC
    '            bGridloaded = True
    '         End If

    '         ' item total
    '         Me.txtItemTotal.Text = FormatCurrency(lcurPrice)

    '         ' total line:
    '         ' Total      Price    Delivery   Deposit   Total
    '         Me.txtTotal.Text = FormatCurrency(lcurPrice + _
    '                            lcurTax + _
    '                            UnFormat(Me.txtDeposit.Text) + _
    '                            lcurDelivery)
    '         Me.txtAmtPaid.Text = Me.txtTotal.Text
    '         Me.txtBalDue.Text = FormatCurrency(UnFormat(Me.txtTotal.Text) - _
    '            UnFormat(Me.txtAmtPaid.Text))
    '      End If
    '   Catch ex As System.Exception
    '      StructuredErrorHandler(ex)
    '   End Try
    'End Sub
#End Region

#Region " Public Properties "
    Public Property CurrentInvoice() As Integer
        Get
            Return m_CurrentInvoice
        End Get
        Set(ByVal Value As Integer)
            m_CurrentInvoice = Value
        End Set
    End Property

#End Region

#Region " Private Methods "
    Private Function VerifySettingsForCkOut() As Boolean
        ' ck to see the correct options are turned on
        If Me.txtCompanyName.Text.Trim.Length = 0 Then
            MsgBox("You must select a customer to bill.", MsgBoxStyle.Exclamation)
            Return False
        End If
        If Not Me.optPaidByCheck.Checked AndAlso Not Me.optPaidByCreditCard.Checked AndAlso Not Me.optBillTo.Checked AndAlso Not Me.optCash.Checked AndAlso Not Me.optLeftBlankCheck.Checked AndAlso Not optLeftCardNumber.Checked Then
            MsgBox("You must select a Payment Arrangement Option", MsgBoxStyle.Exclamation)
            Return False
        End If
        If Me.cbEmployees.Text.Trim.Length = 0 Then
            MsgBox("Please select your employee name.", MsgBoxStyle.Exclamation)
            Return False
        End If
        If Me.optLeftBlankCheck.Checked Or Me.optLeftCardNumber.Checked Then
            MsgBox("Can't leave check or card at check in.", MsgBoxStyle.Exclamation)
            Return False
        End If

        If UnFormat(Me.txtDeposit.Text) > 0 Then
            Dim sMsg As String
            Dim iRV As Integer
            sMsg = "Deposit amount is greater than 0.  Are you going" & Chr(10)
            sMsg &= "to keep the deposit?" & Chr(10)
            sMsg &= "" & Chr(10)
            sMsg &= "Click Yes to continue.  Click No to cancel and then" & Chr(10)
            sMsg &= "clear the deposit amount to 0 and press Manual " & Chr(10)
            sMsg &= "Recalculate button before printing invoice." & Chr(10)
            sMsg &= "" & Chr(10)
            iRV = MsgBox(sMsg, CType(292, Microsoft.VisualBasic.MsgBoxStyle), "Confirm Keeping Deposit")

            If iRV = 6 Then
                ' Yes Code goes here
                Return True
            Else
                ' No code goes here
                Return False
            End If
        End If
        Return True
    End Function

    Private Sub CheckIn()
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
        Dim iInvID As Integer
        Dim o As CItems
        Dim sErr As String
        Dim iRows As Integer
        Dim shtPmtType As Short
        Dim iCustID As Integer = Val(Me.txtCustomerID.Text)
        Dim decTotal As Decimal
        Dim dr As DataRow
        Dim myTransaction As OleDb.OleDbTransaction
        Dim conn As OleDb.OleDbConnection
        Try
            If Not voidInvoice Then
                If Not Me.VerifySettingsForCkOut() Then Exit Sub
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

            iInvID = Val(Me.txtInvoiceID.Text)


            ' First, we mark the items off rent
            For i = 0 To dtList.Rows.Count - 1
                With dtList.Rows(i)
                    If dtList.Rows(i).RowState <> DataRowState.Deleted Then
                        dr = dtList.Rows(i)
                        If dr("rental_period") <> SALE AndAlso _
                           dr("equip_id") <> "Labor" AndAlso _
                           dr("equip_id") <> "Fuel" AndAlso _
                           dr("equip_id") <> "Misc" AndAlso _
                           dr("rental_period") <> "N/A" AndAlso _
                           dr("equip_id") <> RERENT Then
                            SQL = "update equipment "
                            SQL &= "set rented_date = Null, "
                            SQL &= "available='YES', "
                            SQL &= "renting_company_id = Null, "
                            SQL &= "available_date = #" & Me.dtpCkInDateReset.Value.ToString & "# "
                            ' update counter of times used
                            If MNSng(dr("meterin")) > 0 Then
                                SQL &= ", reserving_company_id = " & CType(dr("meterin"), Integer) & " "
                            Else
                                SQL &= ", reserving_company_id = reserving_company_id + 1 "
                            End If
                            SQL &= "where equip_id = '" & dr("equip_id") & "' "
                            cmd.CommandText = SQL
                            iRows = oDA.SendActionSql(cmd, sErr)
                            If iRows < 1 Then
                                StructuredErrorHandler("Ck In Equip not found, ck-in will continue " & Chr(10) & SQL)
                            End If

                            ' if meter used, update meter table
                            If MNSng(dr("meterin")) > 0 Then
                                SQL = "insert into  meter_reading"
                                SQL &= "(equip_id, meter_reading, date_entered, invoice_id,entry_type) "
                                SQL &= "values("
                                SQL &= "'" & dr("equip_id") & "', "
                                SQL &= dr("meterin") & ", "
                                SQL &= "#" & Me.dtpCkInDateReset.Value.ToString & "#, "
                                SQL &= Me.txtInvoiceID.Text & ", "
                                SQL &= "'IN' "
                                SQL &= ")"
                                cmd.CommandText = SQL
                                iRows = oDA.SendActionSql(cmd, sErr)
                                If iRows < 1 Then
                                    Throw New Exception("Update on Meter Reading Table Failed." & vbCrLf & SQL & vbCrLf)
                                End If
                            End If
                        End If

                        ' if we added fuel or labor rows, they must be inserted
                        ' rather than updated...
                        If i > (origItemCount - 1) Then
                            SQL = "insert into invoice_details "
                            SQL &= "(invoiceid,customer_id,quantity,priceperunit,"
                            SQL &= "equip_id,equip_name,rented_date,returned_date,record_type,"
                            SQL &= "deposit,rental_period,delivery,record_description) "
                            SQL &= "values("
                            SQL &= Me.txtInvoiceID.Text & ", "
                            SQL &= Me.txtCustomerID.Text & ", "
                            SQL &= dtList.Rows(i).Item("quantity") & ", "
                            SQL &= dtList.Rows(i).Item("priceperunit") & ", "
                            SQL &= "'" & dtList.Rows(i).Item("equip_id") & "', "
                            SQL &= "'" & dtList.Rows(i).Item("equip_name") & "', "
                            SQL &= "#" & Me.dtpCkInDateReset.Value.ToString & "#, "
                            SQL &= "#" & Me.dtpCkInDateReset.Value.ToString & "#, "
                            SQL &= "15, "
                            SQL &= "0, "
                            SQL &= "'N/A', "
                            SQL &= "0, "
                            SQL &= "'Rent/Sale Item')"
                        Else
                            SQL = "Update invoice_details "
                            SQL &= "set returned_date = #" & Me.dtpCkInDateReset.Value.ToString & "#, "
                            SQL &= "priceperunit = " & .Item("priceperunit") & ", "
                            SQL &= "quantity = " & .Item("quantity") & ", "
                            SQL &= "rental_period = '" & .Item("rental_period") & "' "
                            If .Item("meterin") > 0 Then
                                SQL &= ",meter_in = " & .Item("meterin") & " "
                            End If
                            SQL &= "where invoiceid = " & Val(Me.txtInvoiceID.Text) & " "
                            SQL &= "and record_type = 15 "
                            SQL &= "and equip_id = '" & .Item("equip_id") & "'"
                        End If
                        cmd.CommandText = SQL
                        If oDA.SendActionSql(cmd, sErr) = 0 Then
                            Throw New Exception("Update of Invoice Detail for invoice: " & Me.txtInvoiceID.Text & _
                               " EquipID: " & .Item("equip_id") & " failed.")
                        End If

                        ' if rerent, mark as returned in rerent table
                        If dr("equip_id") = RERENT Then
                            SQL = "update rerents set returned_date = #" & Now.ToString & "# "
                            SQL &= "where unique_id = " & dr("rerent_id")
                            cmd.CommandText = SQL
                            If oDA.SendActionSql(cmd, sErr) = 0 Then
                                Throw New Exception("Update of rerent table failed to set returned_date for item: " & dr("rerent_id"))
                            End If
                        End If
                    End If
                End With
            Next

            ' delete all of the invoice total items
            ' 0 rows deleted is ok here if there is no equipment
            ' on the invoice, which can be the case when an invoice is split...
            SQL = "delete from invoice_details "
            SQL &= "where invoiceid = " & Val(Me.txtInvoiceID.Text) & " "
            SQL &= "and record_type <> 15 "
            cmd.CommandText = SQL
            If oDA.SendActionSql(cmd, sErr) < 0 Then
                Throw New Exception("Failed to delete invoice details " & Chr(10) & SQL)
            End If

            ' if the invoice is to be voided, leave the invoice details
            ' alone and mark the invoice header
            If Not voidInvoice Then
                ' now reinsert the total records after the user has settled the deal
                If Val(UnFormat(Me.txtDeposit.Text)) > 0 Then
                    SQL = "Insert into invoice_details "
                    SQL &= "(invoiceid,record_type,Deposit, Customer_id,record_description) "
                    SQL &= "values("
                    SQL &= iInvID & ", "
                    SQL &= "25" & ", " 'manual deposit record
                    SQL &= UnFormat(Me.txtDeposit.Text) & ", "
                    'SQL &= shtPmtType.ToString & ", "
                    SQL &= iCustID.ToString
                    SQL &= ",'Equip/Sale Item'"
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
                    SQL &= "(invoiceid,record_type,delivery, customer_id,record_description) "
                    SQL &= "values("
                    SQL &= iInvID & ", "
                    SQL &= "45" & ", " 'delivery
                    SQL &= UnFormat(Me.txtDelivery.Text) & ", "
                    'SQL &= shtPmtType.ToString & ", "
                    SQL &= iCustID.ToString
                    SQL &= ",'Delivery'"
                    SQL &= ")"
                    cmd.CommandText = SQL
                    iRows = oDA.SendActionSql(cmd, sErr)
                    If iRows < 1 Then
                        Throw New Exception("Invoice Detail Update Failure: " & Chr(10) & _
                           sErr & Chr(10) & SQL)
                    End If
                End If

                ' determine pickup cost if applicable
                If Val(UnFormat(Me.textManualPickup.Text)) > 0 Then
                    SQL = "Insert into invoice_details "
                    SQL &= "(invoiceid,record_type,delivery, customer_id,record_description) "
                    SQL &= "values("
                    SQL &= iInvID & ", "
                    SQL &= "46" & ", " 'pickup
                    SQL &= UnFormat(Me.textManualPickup.Text) & ", "
                    'SQL &= shtPmtType.ToString & ", "
                    SQL &= iCustID.ToString
                    SQL &= ",'Pickup'"
                    SQL &= ")"
                    cmd.CommandText = SQL
                    iRows = oDA.SendActionSql(cmd, sErr)
                    If iRows < 1 Then
                        Throw New Exception("Invoice Detail Update Failure: " & Chr(10) & _
                           sErr & Chr(10) & SQL)
                    End If
                End If

                ' create sales tax record if applicable
                ' amtpaid will hold the itemTotal amount for the sales tax item
                ' create sale tax record, even if zero so we have the itemTotal
                'If Val(UnFormat(Me.txtSalesTax.Text)) > 0 Then
                SQL = "Insert into invoice_details "
                SQL &= "(invoiceid,record_type,salestax,amtpaid,customer_id,record_description) "
                SQL &= "values("
                SQL &= iInvID & ", "
                SQL &= "35" & ", " ' sales tax record
                SQL &= UnFormat(Me.txtSalesTax.Text) & ", "
                SQL &= UnFormat(Me.txtItemTotal.Text) & ", "
                SQL &= iCustID.ToString
                SQL &= ",'Sales Tax'"
                'SQL &= "'" & Me.txtTaxID.Text & "'"
                SQL &= ")"
                cmd.CommandText = SQL
                iRows = oDA.SendActionSql(cmd, sErr)
                If iRows < 1 Then
                    Throw New Exception("Invoice Detail Update Failure: " & Chr(10) & _
                       sErr & Chr(10) & SQL)
                End If
                'End If
                'End If

                ' create amt paid record if applicable
                If Val(UnFormat(Me.txtAmtPaid.Text)) > 0 Then
                    SQL = "Insert into invoice_details "
                    SQL &= "(invoiceid,record_type,amtpaid, customer_id,record_description) "
                    SQL &= "values("
                    SQL &= iInvID & ", "
                    SQL &= "55" & ", " ' amt paid record
                    SQL &= UnFormat(Me.txtAmtPaid.Text) + UnFormat(Me.txtAmtPaidAtCkIn.Text) & ", "
                    SQL &= iCustID.ToString
                    SQL &= ",'Paid at CkIn'"
                    SQL &= ")"
                    cmd.CommandText = SQL
                    iRows = oDA.SendActionSql(cmd, sErr)
                    If iRows < 1 Then
                        Throw New Exception("Invoice Detail Update Failure: " & Chr(10) & _
                           sErr & Chr(10) & SQL)
                    End If
                End If
                'End If

                ' create refund paid record if applicable
                Dim bal As Decimal = Val(UnFormat(Me.txtBalDue.Text))
                If bal <> 0 Then
                    SQL = "Insert into invoice_details "
                    SQL &= "(invoiceid,record_type,amtpaid, customer_id,record_description) "
                    SQL &= "values("
                    SQL &= iInvID & ", "
                    If bal > 0 Then
                        SQL &= "75," ' bal due
                    Else
                        SQL &= "65, " ' refund
                    End If

                    SQL &= bal & ", "
                    SQL &= iCustID.ToString
                    If bal > 0 Then
                        SQL &= ",'Bal Due'"
                    Else
                        SQL &= ",'Refund'"
                    End If
                    SQL &= ")"
                    cmd.CommandText = SQL
                    iRows = oDA.SendActionSql(cmd, sErr)
                    If iRows < 1 Then
                        Throw New Exception("Invoice Detail Update Failure: " & Chr(10) & _
                           sErr & Chr(10) & SQL)
                    End If
                End If

                ' close the invoice header
                Dim sPaidBy As String = ""
                Select Case True
                    Case Me.optBillTo.Checked : sPaidBy = "BT"
                    Case Me.optCash.Checked : sPaidBy = "CA"
                    Case Me.optLeftBlankCheck.Checked : sPaidBy = "BC"
                    Case Me.optLeftCardNumber.Checked : sPaidBy = "LC"
                    Case Me.optPaidByCheck.Checked : sPaidBy = "CK"
                    Case Me.optPaidByCreditCard.Checked : sPaidBy = "CC"
                End Select
                SQL = "update invoices set "
                If bal <> 0 Then
                    SQL &= "balancedue = " & UnFormat(Me.txtBalDue.Text) & ", "
                    SQL &= "status = 'OPEN' "
                Else
                    SQL &= "balancedue= 0, "
                    SQL &= "status = 'CLOSED' "
                End If
                SQL &= ",PaidOption = '" & sPaidBy & "', "
                SQL &= "notes = '" & Me.txtNotes.Text.Replace("'", "''") & "', "
                SQL &= "ponumber = '" & MNS(Me.txtPONbr.Text) & "', "
                SQL &= "elapsed_time = '" & MNS(Me.textElapsedTimeHours.Text) & "' "
            Else
                SQL = "update invoices set notes = 'INVOICE IS VOID', "
                SQL &= "balancedue = 0, status = 'CLOSED' "
            End If
            SQL &= ",check_in_employee = '" & Me.cbEmployees.Text & "' "
            SQL &= "where invoiceid = " & iInvID
            cmd.CommandText = SQL
            iRows = oDA.SendActionSql(cmd, sErr)
            If iRows <> 1 Then
                Throw New System.Exception("update of invoice header failed: " & iInvID.ToString & vbCrLf & SQL & vbCrLf)
            End If

            ' save customer data
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

        Catch ex As System.Exception
            myTransaction.Rollback()
            StructuredErrorHandler(ex)
            MsgBox("Please email lesterlsmith@gmail.com and attach the lastest errlog txt file.", MsgBoxStyle.Exclamation)
        End Try
    End Sub

    Public Sub PrintInvoice(ByVal InvoiceId As Integer)
        ' format the print line
        Dim ps As New System.Text.StringBuilder()
        Dim ps2 As New System.Text.StringBuilder()
        Dim decEP As Decimal
        Dim SQL As String
        Dim i As Integer
        Dim dt As New DataTable()
        Dim oUtil As New CUtilities()
        Dim decTotal As Decimal
        Dim sName As String
        ' get customer data and print
        '#If CustomerApp = RELIABLE Then
        '      Dim oPR As New CReliablePrint.CReliablePrint(Me)
        '#Else
        Dim oPR As New CPioneerInvoice.CPioneerPrepInvoice(Me)
        '#End If
        Try
            oPR.PrintCheckInInvoice(Me.txtInvoiceID.Text)
            '         Dim trash As String = Me.cbEmployees.Text
            '         If PrintInitialsOnly Then
            '            CheckOutEmployee = oUtil.GetToken(trash)
            '         Else
            '            CheckOutEmployee = oUtil.GetToken(trash, " ")
            '            CheckOutEmployee = oUtil.GetToken(trash, " ")
            '         End If


            '         For i = 0 To dtList.Rows.Count - 1
            '            With dtList.Rows(i)
            '               ' qty
            '               ps.Append(CType(.Item("Quantity"), String).PadLeft(6))
            '               ' skip 3 spaces and print id - name
            '               ps.Append(Space(3) & CType(.Item("Equip_Id"), String).PadRight(10) & " - ")
            '               sName = .Item("equip_name")
            '               If sName.Length > 27 Then
            '                  sName = sName.Substring(0, 27)
            '               End If
            '               ' qty
            '               ps.Append(sName.PadRight(28))

            '               ' rental period (Daily...)
            '               ps.Append(Space(9) & CType(.Item("Rental_Period"), String).PadRight(10))
            '               ' price per unit
            '               ps.Append(Format(.Item("priceperunit"), "#,##0.00").PadLeft(10))
            '               decEP = .Item("PriceperUnit") * _
            '                      .Item("Quantity") ' + _
            '               ps.Append(Space(2) & Format(decEP, "#,##0.00").PadLeft(10) & vbCrLf)

            '               ' if meter_required, print the meter reading at checkout
            '               If .Item("meterin") > 0 Then
            '                  ps.Append(Space(9) & "Meter Out: " & Format(.Item("meterout"), "0.00") & _
            '                                       " In: " & Format(.Item("meterin"), "0.00") & vbCrLf)
            '               End If

            '            End With
            '         Next i

            '         Const DTSP = 70
            '         If Not voidInvoice Then
            '            ' print the totals
            '            ps.Append(vbCrLf & Space(DTSP) & "Item Total".PadRight(11) & Me.txtItemTotal.Text.PadLeft(10) & vbCrLf)
            '            If UnFormat(Me.txtDeposit.Text) <> 0 Then
            '               ps.Append(vbCrLf & Space(DTSP) & "Deposit".PadRight(11) & Me.txtDeposit.Text.PadLeft(10) & vbCrLf)
            '            End If

            '            If UnFormat(Me.txtSalesTax.Text) > 0 Then
            '               ps.Append(vbCrLf & Space(DTSP) & "Sales Tax".PadRight(11) & Me.txtSalesTax.Text.PadLeft(10) & vbCrLf)
            '            End If

            '            If UnFormat(Me.txtDelivery.Text) <> 0 Then
            '               ps.Append(vbCrLf & Space(DTSP) & "Delivery".PadRight(11) & Me.txtDelivery.Text.PadLeft(10) & vbCrLf)
            '            End If

            '            If UnFormat(Me.textManualPickup.Text) <> 0 Then
            '               ps.Append(vbCrLf & Space(DTSP) & "Pickup".PadRight(11) & Me.textManualPickup.Text.PadLeft(10) & vbCrLf)
            '            End If
            '            ps.Append(vbCrLf & Space(DTSP) & "Total".PadRight(11) & Me.txtTotal.Text.PadLeft(10) & vbCrLf)
            '            If UnFormat(Me.txtAmtPaid.Text) <> 0 Then
            '               ps.Append(vbCrLf & Space(DTSP) & "Paid/CkOut".PadRight(11) & Me.txtAmtPaid.Text.PadLeft(10) & vbCrLf)
            '            End If

            '            If UnFormat(Me.txtAmtPaidAtCkIn.Text) > 0 Then
            '               ps.Append(vbCrLf & Space(DTSP) & "Paid/CkIn".PadRight(11) & Me.txtAmtPaidAtCkIn.Text.PadLeft(10) & vbCrLf)
            '            End If
            '            If UnFormat(Me.txtBalDue.Text) < 0 Then
            '               Dim valu As Decimal = UnFormat(Me.txtBalDue.Text)
            '               ps.Append(vbCrLf & Space(DTSP) & "Refund Due".PadRight(11) & _
            '                  FormatCurrency(valu * -1).PadLeft(10) & vbCrLf)
            '            Else
            '               ps.Append(vbCrLf & Space(DTSP) & "Bal Due".PadRight(11) & Me.txtBalDue.Text.PadLeft(10) & vbCrLf)
            '            End If

            '            If Me.txtNotes.Text.Trim.Length > 0 Then
            '               Dim sMemo As String = Me.txtNotes.Text
            '               Dim iNL As Integer = oUtil.MLCount(sMemo, 60)
            '               Dim k As Integer
            '               ps.Append(vbCrLf & vbCrLf & "Notes:" & vbCrLf)

            '               For k = 1 To iNL
            '                  ps.Append(oUtil.MemoLine(sMemo, 60, k) & vbCrLf)
            '               Next
            '            End If
            '         Else
            '            ps.Append(vbCrLf & Space(DTSP) & "Bal Due".PadRight(11) & FormatCurrency(0).PadLeft(10) & vbCrLf)
            '            ps.Append(vbCrLf & vbCrLf & "Notes:" & " INVOICE IS VOID" & vbCrLf)
            '         End If

            '#If Reliable Then
            '         Dim dti As New DataTable()
            '         SQL = "select * from invoices where invoiceid = '" & Me.txtInvoiceID.Text & "'"
            '         If oDA.SendQuery(SQL, dti, ConnectString) = 0 Then
            '            MsgBox("Can't read invoice header record.", MsgBoxStyle.Critical)
            '            Exit Sub
            '         End If
            '         Dim dr As DataRow = dti.Rows(0)

            '         Dim billTo As String = Me.txtCompanyName.Text & vbCrLf & _
            '                                Me.txtBillingAddress1.Text & vbCrLf & _
            '                                Me.txtCity.Text & ", " & Me.txtState.Text & " " & Me.txtPostalCode.Text
            '         Dim shipTo As String = Me.txtShipToCustomer.Text & vbCrLf & _
            '                                Me.txtShipAddress1.Text & vbCrLf & _
            '                                Me.txtShipCity.Text & vbCrLf

            '         Dim invoice As String = Me.txtInvoiceID.Text
            '         Dim timeOut As String = Me.lblCkOutDate.Text
            '         Dim timeIn As String = Me.lblCkInDate.Text
            '         Dim elapsedHours As String = _
            '            Format(DateDiff(DateInterval.Hour, CType(Me.lblCkOutDate.Text, DateTime), CType(Me.lblCkInDate.Text, DateTime)), "0")
            '         Dim jobPhone As String = "None"
            '         Dim ckInBy As String = Me.cbEmployees.Text
            '         Dim dlnbr As String = MNS(Me.textDriversLicence.Text)
            '         Dim agent As String = MNS(dr("contactname"))
            '         Dim poNbr As String = MNS(dr("ponumber"))
            '         Dim writtenBy As String = MNS(dr("check_out_employee"))
            '         Dim detailLines As String = ps.ToString
            '         Dim dueIn As String = ""
            '         Dim jobLocation As String = ""
            '         Dim totalDesc As String = ps2.ToString()

            '         Dim oPD As New CRelialblePrintObject(billTo, _
            '            shipTo, invoice, timeOut, timeIn, elapsedHours, _
            '            jobPhone, ckInBy, dlnbr, agent, poNbr, writtenBy, _
            '            detailLines, dueIn, jobLocation, totalDesc)

            '#Else
            '         Dim oPD As New CPioneerPrint(Me)
            '#End If

            '         If modMain.fMainForm.mnuPreviewBeforePrint.Checked Then
            '            oPD.Preview()
            '         Else
            '            oPD.Print()
            '         End If
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

    Private Sub PrintHandler()
        '  Print the checkout bill
        If Not Me.VerifySettingsForCkOut Then Exit Sub
        m_bPrintedInvoice = True
        PrintInvoice(Val(Me.txtInvoiceID.Text))
        Me.btnCheckOut.Enabled = True
    End Sub

    Private Sub LoadPickupCombo()
        Dim SQL As String = ""
        Dim dt As New DataTable()
        Dim i As Integer

        SQL &= "select delivery_price,delivery_desc "
        SQL &= "from delivery_pickup "
        SQL &= "where delivery_pickup ='P' "
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
        UnFmt_T_B = Val(Replace(Replace(Replace(Replace(Replace(roTB.Text, "$", ""), ",", ""), ")", ""), "(", ""), "%", ""))
        If InStr(roTB.Text, "%") Then
            UnFmt_T_B = UnFmt_T_B / 100
        End If
        If InStr(roTB.Text, "(") > 0 And InStr(roTB.Text, ")") > 0 Then
            UnFmt_T_B = UnFmt_T_B * -1
        End If
    End Function
    Public Function Fmt_T_B(ByRef roTB As System.Windows.Forms.TextBox) As String
        On Error Resume Next
        If InStr(1, roTB.Tag, ";", 1) > 0 Then
            If InStr(roTB.Text, "-") > 0 Or (InStr(roTB.Text, "(") > 0 And InStr(roTB.Text, ")") > 0) Then
                Fmt_T_B = VB6.Format(System.Math.Abs(CDbl(roTB.Text)), Mid(roTB.Tag, InStr(roTB.Tag, ";") + 1))
            Else
                Fmt_T_B = VB6.Format(roTB.Text, VB.Left(roTB.Tag, InStr(roTB.Tag, ";") - 1))
            End If
        ElseIf InStr(1, roTB.Tag, "%", 1) > 0 Then
            Fmt_T_B = VB6.Format(roTB.Text, roTB.Tag)
        Else
            Fmt_T_B = VB6.Format(roTB.Text, roTB.Tag)
        End If
    End Function
    Public Function Fmt_D_F(ByRef rsTxt As Object, ByRef roTB As System.Windows.Forms.TextBox) As String
        On Error Resume Next

        If InStr(1, roTB.Tag, ";", 1) > 0 Then
            If InStr(rsTxt, "-") Then

                Fmt_D_F = VB6.Format(Replace(rsTxt, "-", ""), Mid(roTB.Tag, InStr(roTB.Tag, ";") + 1))
            Else
                Fmt_D_F = VB6.Format(rsTxt, VB.Left(roTB.Tag, InStr(roTB.Tag, ";") - 1))
            End If
        ElseIf InStr(1, roTB.Tag, "%", 1) > 0 Then
            Fmt_D_F = VB6.Format(rsTxt, roTB.Tag)
        Else
            Fmt_D_F = VB6.Format(rsTxt, roTB.Tag)
        End If
    End Function


    Private Sub CancelHandler()
        ' No code goes here
        If Me.m_bPrintedInvoice Then
            Dim sMsg As String
            Dim iRV As Integer
            sMsg = "You have previewed or printed the invoice, do " & Chr(10)
            sMsg &= "you want to close without checking in the equipment?" & Chr(10)
            sMsg &= "" & Chr(10)
            sMsg &= "Click Yes to close the form without checking out;" & Chr(10)
            sMsg &= "Click No to checkout, or Cancel to leave the checkin" & Chr(10)
            sMsg &= "form open." & Chr(10)
            sMsg &= "" & Chr(10)
            iRV = MsgBox(sMsg, CType(35, Microsoft.VisualBasic.MsgBoxStyle), "Confirm Cancel")

            If iRV = 6 Then
                ' Yes Code goes here
            ElseIf iRV = 7 Then
                ' No code goes here
                CheckIn()
            Else
                ' Cancel code goes here
                Exit Sub
            End If
        End If
        Me.Close()
        DoEvents()
    End Sub

    Private Sub InitialLoadData()
        Dim dt As New DataTable()
        Dim SQL As String
        Dim oDA As New CDataAccess()
        ' for chek in we need to 
        ' 1) load the invoice header data

        SQL = "select * from invoices "
        SQL &= "where invoiceid = " & m_CurrentInvoice & " "
        If oDA.SendQuery(SQL, dt, ConnectString) > 0 Then
            Dim dr As DataRow = dt.Rows(0)
            Me.txtNotes.Text = IIf(IsDBNull(dr("Notes")), "", dr("Notes"))
            Me.txtInvoiceID.Text = dr("invoiceid")
            Me.txtCustomerID.Text = dr("customerid")
            Me.txtPONbr.Text = dr("ponumber")
            Me.txtCheckNumber.Text = dr("ckcardnumber")
            Me.cbExpMon.Text = MNS(dr("exp_month"))
            Me.cbExpYr.Text = MNS(dr("exp_yr"))
            Me.txtCardId.Text = MNS(dr("card_id"))
            Me.txtContactName.Text = dr("contactname")
            Me.lblWrittenBy.Text = MNS(dr("check_out_employee"))
            'Select Case dr("paidoption")
            '   Case "CK" : Me.optPaidByCheck.Checked = True
            '   Case "CA" : Me.optCash.Checked = True
            '   Case "CC" : Me.optPaidByCreditCard.Checked = True
            '   Case "LC" : Me.optLeftCardNumber.Checked = True
            '   Case "BC" : Me.optLeftBlankCheck.Checked = True
            '   Case "BT" : Me.optBillTo.Checked = True
            'End Select
            Me.txtShipToCustomer.Text = dr("shiptocustomer")
            Me.txtShipAddress1.Text = dr("shiptoaddress")
            Me.txtShipCity.Text = dr("shiptocity")
            Me.txtShipState.Text = dr("shiptostate")
            Me.txtShipZip.Text = dr("shiptozip")
            Me.textDriversLicence.Text = MNS(dr("drivers_license"))
            Me.lblCkOutDate.Text = Format(dr("invoicedate"), "MM/dd/yy hh:mm tt")
            Me.lblCkInDate.Text = Format(Now, "MM/dd/yy hh:mm tt")
        End If

        ' 2) load the customer data
        SQL = "select companyname,billingaddress1,billingaddress2, "
        SQL &= "city,state,postalcode,tax_id "
        SQL &= " from customers "
        SQL &= "where customerid = " & dt.Rows(0).Item("customerid") & " "
        dt.Reset()
        If oDA.SendQuery(SQL, dt, ConnectString) > 0 Then
            With dt.Rows(0)
                Me.txtCompanyName.Text = MNS(.Item("companyname"))
                Me.txtBillingAddress1.Text = MNS(.Item("billingaddress1"))
                Me.txtBillingAddress2.Text = MNS(.Item("billingaddress2"))
                Me.txtCity.Text = MNS(.Item("city"))
                Me.txtState.Text = MNS(.Item("state"))
                Me.txtPostalCode.Text = MNS(.Item("postalcode"))
                Try
                    Me.txtTaxID.Text = StringEncryption.DecryptString(MNS(.Item("tax_id")))
                Catch ex As System.Exception
                    Me.txtTaxID.Text = MNS(.Item("tax_id"))
                End Try
            End With
        End If

        ' 3) load the total boxes
        SQL = "select record_type,priceperunit,deposit,salestax,delivery,amtpaid "
        SQL &= "from invoice_details "
        SQL &= "where invoiceid = " & m_CurrentInvoice & " "
        SQL &= "and record_type <> 15 "
        SQL &= "order by record_type"
        dt.Reset()
        Dim i As Integer
        If oDA.SendQuery(SQL, dt, ConnectString) > 0 Then
            For i = 0 To dt.Rows.Count - 1
                With dt.Rows(i)
                    Select Case .Item("record_type")
                        Case 15 ' details
                            Me.txtTotal.Text = FormatCurrency(.Item("priceperunit"))
                        Case 25 ' deposit
                            Me.txtDeposit.Text = FormatCurrency(.Item("deposit"))
                        Case 35 ' sales tax
                            ' if there was a sales tax amt and taxid is extant we have an issue
                            If .Item("Salestax") > 0 AndAlso Not String.IsNullOrEmpty(txtTaxID.Text) Then
                                Dim msg As String = "You had sales tax at Checkout but you have a Tax Id. Click Yes to clear the sales tax or No to allow the tax. If you click No, you should go to the Customer record and remove the Tax Id."
                                Dim ans As DialogResult = MessageBox.Show(msg, "Tax Id - Checkout Tax Inconsistency", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
                                If ans <> DialogResult.Yes Then
                                    Me.txtSalesTax.Text = FormatCurrency(.Item("salestax"))
                                Else
                                    Me.txtSalesTax.Text = FormatCurrency(0)
                                End If
                            End If
                        Case 45 ' deleivery
                            Me.txtDelivery.Text = FormatCurrency(.Item("delivery"))
                        Case 55 ' amt paid
                            Me.txtAmtPaid.Text = FormatCurrency(.Item("amtpaid"))
                    End Select
                End With
            Next i
        End If

        ' 2) load the invoice data into the grid (record type 1)
        SQL = "select Equip_Id,Equip_Name,Quantity,Rental_Period,"
        SQL &= "PricePerUnit,Rented_Date,"
        SQL &= "iif(isnull(meter_out),0,meter_out) as MeterOut, "
        SQL &= "iif(isnull(meter_in),0,meter_in) as MeterIn,rentalduetoreturn "
        SQL &= ",hourrate,halfday,daily,weekly,monthly,weekend,newprices,rerent_id "
        SQL &= "from invoice_details "
        SQL &= "where invoiceid = " & m_CurrentInvoice & " "
        SQL &= "and record_type = 15 "
        oCG.ClearDataTableForRebinding(dtList)

        If oDA.SendQuery(SQL, dtList, ConnectString, "dt") > 0 Then
            Dim formats() As String = _
               {"", "60", "T", "L", _
               "", "200", "T", "L", _
               "", "60", "F", "R", _
               "", "60", "F", "L", _
               "$#,#0.00", "100", "F", "R", _
               "MM/dd/yyyy HH:mm tt", "100", "T", "L", _
               "0.00", "60", "T", "R", _
               "0.00", "60", "F", "R", _
               "MM/dd/yyyy HH:mm tt", "100", "T", "L"}
            oCG.SetTablesStyle(dtList, Me.dbgShoppingList, formats)
            Me.dbgShoppingList.SetDataBinding(dtList, "")
            oCG.DisableAddNew(Me.dbgShoppingList, Me)

            TotalTheItems()
            ' remember how many items we had to start
            origItemCount = dtList.Rows.Count
        End If
    End Sub

    Private Sub TotalTheItems()
        Dim i As Integer
        Dim total As Decimal = 0
        Dim laborCost As Decimal

        For i = 0 To dtList.Rows.Count - 1
            With dtList.Rows(i)
                total += dtList.Rows(i).Item("priceperunit") * Val(dtList.Rows(i).Item("quantity"))
                If .Item("equip_id") = "Labor" Then
                    laborCost += .Item("priceperunit") * .Item("priceperunit")
                End If
            End With
        Next
        Me.txtItemTotal.Text = FormatCurrency(total)
        If Me.txtSalesTax.Visible Then
            total += UnFormat(Me.txtSalesTax.Text)
        End If
        total += UnFormat(Me.txtDelivery.Text)
        total += UnFormat(Me.txtDeposit.Text)
        Me.txtTotal.Text = FormatCurrency(total)
        total -= UnFormat(Me.txtAmtPaid.Text)
        Me.txtBalDue.Text = FormatCurrency(total)
    End Sub

    Private Sub MovePaidToDue()
        'Me.txtBalDue.Text = Me.txtAmtPaid.Text
        'Me.txtAmtPaid.Text = FormatCurrency(0)
    End Sub
    Private Sub MoveDueToPaid()
        'Me.txtAmtPaid.Text = Me.txtBalDue.Text
        'Me.txtBalDue.Text = FormatCurrency(0)
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

    Private Sub ManualRecalc()
        Dim laborCost As Decimal

        Try
            Dim amt As Decimal = 0
            Dim i As Integer

            With Me
                ' total up the detail rows
                For i = 0 To dtList.Rows.Count - 1
                    If Not (dtList.Rows(i).RowState = DataRowState.Deleted) Then
                        With dtList.Rows(i)
                            amt += MND(.Item("priceperunit")) * MNI(.Item("quantity"))

                            If MNS(dtList.Rows(i).Item("equip_id")) = "Labor" Then 'Or dtList.Rows(i).Item("equip_id") = "Fuel" Then
                                ' total up the labor costs so we can subtract
                                ' it out before computing tax
                                ' right now we are charging sales tax on fuel...
                                laborCost += MND(.Item("priceperunit")) * MNI(.Item("quantity"))
                            End If
                        End With
                    End If
                Next

                .txtItemTotal.Text = FormatCurrency(amt)

                If String.IsNullOrEmpty(.txtTaxID.Text) Then
                    .txtSalesTax.Text = FormatCurrency((UnFormat(.txtItemTotal.Text) + _
                                                        UnFormat(.txtDelivery.Text) + _
                                                        UnFormat(textManualPickup.Text) - laborCost) * TaxRate)
                Else
                    .txtSalesTax.Text = FormatCurrency(0)
                End If

                amt = FormatCurrency(UnFormat(.txtItemTotal.Text) +
                      UnFormat(.txtDelivery.Text) +
                      UnFormat(.textManualPickup.Text) +
                      UnFormat(Me.txtDeposit.Text) +
                      IIf(String.IsNullOrEmpty(txtTaxID.Text), UnFormat(.txtSalesTax.Text), 0))
                'IIf(Me.txtTaxID.Visible, UnFormat(.txtSalesTax.Text), 0))

                Me.txtTotal.Text = FormatCurrency(amt)

                If .optBillTo.Checked Then
                    Me.txtBalDue.Text = FormatCurrency(UnFormat(Me.txtTotal.Text) - _
                                   UnFormat(Me.txtAmtPaid.Text))
                    Me.txtAmtPaidAtCkIn.Text = FormatCurrency(0)
                ElseIf Me.optPaidByCheck.Checked Or _
                   Me.optPaidByCreditCard.Checked Or _
                   Me.optCash.Checked Or _
                   Me.optLeftBlankCheck.Checked _
                   Then
                    Dim apckin As Decimal = amt - UnFormat(Me.txtAmtPaid.Text)
                    If apckin > 0 Then
                        .txtAmtPaidAtCkIn.Text = FormatCurrency(apckin)
                        Me.txtBalDue.Text = FormatCurrency(0)
                    ElseIf apckin < 0 Then
                        .txtBalDue.Text = FormatCurrency(apckin)
                        Me.lblBalDue.Text = "Refund Due"
                    Else
                        .txtBalDue.Text = FormatCurrency(0)
                    End If
                Else
                    Me.txtBalDue.Text = FormatCurrency(UnFormat(Me.txtTotal.Text) - _
                                   UnFormat(Me.txtAmtPaid.Text))
                    Me.txtAmtPaidAtCkIn.Text = FormatCurrency(0)
                End If
            End With
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

    Private Sub CheckInHandler()
        CheckIn()
        Me.Close()
        System.Windows.Forms.Application.DoEvents()
        modMain.fMainForm.LoadEquipGridFromType(0)

    End Sub

#End Region

#Region " Form & Control Events "
    Private Sub frmCheckinNew_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 27 Then
            Me.Close()
            System.Windows.Forms.Application.DoEvents()
        End If

        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub chkManualDepositOverride_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        'LoadTheGrid()
        ManualRecalc()
    End Sub


    Public Sub cmdPrintContract_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPrintContract.Click
        PrintHandler()
    End Sub


    Public Sub btnDelete_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnDelete.Click
        Dim sMsg As String
        Dim iRV As Integer
        sMsg = "Are you absolutely sure you want to delete the " & Chr(10)
        sMsg &= "selected row?" & Chr(10)
        sMsg &= "" & Chr(10)
        sMsg &= "If you delete the row, and you later determine that " & Chr(10)
        sMsg &= "you want it back in the grid, you will have to cancel" & Chr(10)
        sMsg &= "out of the Check In form and start the Check In " & Chr(10)
        sMsg &= "process over." & Chr(10)
        sMsg &= "" & Chr(10)
        sMsg &= "Click Yes to Delete, No to cancel the deletion." & Chr(10)
        sMsg &= "" & Chr(10)
        iRV = MsgBox(sMsg, CType(308, Microsoft.VisualBasic.MsgBoxStyle), "Confirm Deletion of Selected Row")

        If iRV = 6 Then
            ' Yes Code goes here
        Else
            ' No code goes here
            Exit Sub
        End If

        ' add the deleted sale items back to the inventory
        Dim i As Integer = Me.dbgShoppingList.CurrentCell.RowNumber

        Dim s As String = MNS(Me.dtList.Rows(i).Item("rental_period"))
        Dim units As Integer = MNI(Me.dtList.Rows(i).Item("quantity"))
        Dim id As String = MNS(Me.dtList.Rows(i).Item("equip_id"))
        Dim errMsg As String = String.Empty

        If s.Trim.ToUpper = "SALE" Then
            Dim sql As String = "Update products "
            sql &= "set unitsinstock = unitsinstock + " & units & " "
            sql &= "where ProductId = '" & id & "'"
            Dim oda As New CDataAccess()
            If oda.SendActionSql(sql, ConnectString, errMsg) <> 1 Then
                MsgBox("Update of sale item inventory failed.", MsgBoxStyle.Exclamation)
                Exit Sub
            End If
        End If
        Me.dtList.Rows(listHitRow).Delete()
        ManualRecalc()
        'LoadTheGrid()
    End Sub

    Private Sub frmCheckinNew_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        If mbFormLoading Then
            mbFormLoading = False
            Me.Top = 0
            oCI.GetMeterHours()
            If modMain.AutoCalcOn Then
                Me.lblCkInDate.Text = oCI.AutoCalcCheckIn().ToString
                ManualRecalc()
                TotalTheItems()
            End If
            Me.textElapsedTimeHours.Text = modAutoCalc.ElapsedTime
        End If
    End Sub



    Private Sub cmdClose_Click()
        Me.Close()
        System.Windows.Forms.Application.DoEvents()
    End Sub


    Public Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
        CancelHandler()
    End Sub

    Private Sub frmCheckinNew_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        CenterForm(Me)
        Me.cbEmployees.Text = GetSetting(RENTALPRO, SETTINGS, "CKINEMP", "")
        mbFormLoading = True
        LoadPickupCombo()
        LoadEmployeeCombo()
        InitialLoadData()
        Me.optLeftBlankCheck.Enabled = False
        Me.optLeftCardNumber.Enabled = False
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
    End Sub
    Private Sub txtCompanyName_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCompanyName.Enter
        txtCompanyName.SelectionStart = 0
        txtCompanyName.SelectionLength = Len(Trim(txtCompanyName.Text))
    End Sub

    Private Sub LoadCustomerBoxes(ByRef rsCustomerName As String)
        Dim SQL As String
        Dim dt As New DataTable()

        SQL = ""
        SQL = SQL & "select customerid,companyname,billingaddress1,billingaddress2, "
        SQL = SQL & "billingaddress3,city,state,postalcode,phonenumber, "
        SQL = SQL & "contactname,contacttitle,customerid,tax_id "
        SQL = SQL & "from customers "
        SQL = SQL & "where companyname = '" & rsCustomerName & "'"
        oDA.SendQuery(SQL, dt, ConnectString)

        If dt.Rows.Count > 0 Then
            With dt.Rows(0)
                Me.txtCustomerID.Text = MNS(.Item("customerid"))
                Me.txtCompanyName.Text = MNS(.Item("companyname"))
                Me.txtBillingAddress1.Text = MNS(.Item("billingaddress1"))
                Me.txtBillingAddress2.Text = MNS(.Item("bilingaddress2"))
                Me.txtCity.Text = MNS(.Item("city"))
                Me.txtState.Text = MNS(.Item("state"))
                Me.txtPostalCode.Text = MNS(.Item("postalcode"))
                Try
                    Me.txtTaxID.Text = StringEncryption.DecryptString(MNS(.Item("tax_id")))
                Catch ex As System.Exception
                    Me.txtTaxID.Text = MNS(.Item("tax_id"))
                End Try
            End With
        End If
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
            ManualRecalc()
        End If
    End Sub

    Private Sub optLeftBlankCheck_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optLeftBlankCheck.CheckedChanged
        If Me.optLeftBlankCheck.Checked Then
            'MovePaidToDue()
            ManualRecalc()
        End If
    End Sub

    Private Sub optLeftCardNumber_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optLeftCardNumber.CheckedChanged
        'MovePaidToDue()
        ManualRecalc()
    End Sub
    Private Sub optCash_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles optCash.CheckedChanged
        ManualRecalc()
    End Sub

    Private Sub optPaidByCheck_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optPaidByCheck.CheckedChanged
        If Me.optPaidByCheck.Checked Then
            'MoveDueToPaid()
            ManualRecalc()
        End If
    End Sub
    Private Sub optPaidByCreditCard_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles optPaidByCreditCard.CheckedChanged
        ManualRecalc()
    End Sub



    ''' <summary>
    ''' User wants to add a new row to the grid for a half day
    ''' </summary>
    ''' <param name = "sender"></param>
    ''' <param name = "e"></param>
    Private Sub mnuAddHalfDay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddHalfDay.Click
        Dim hdRate As Decimal
        Dim eID As String
        Dim dt As New DataTable()
        Dim iPid As Integer

        Try
            Dim iRow As Integer = Me.dbgShoppingList.CurrentCell.RowNumber

            ' get the half day rate
            eID = dtList.Rows(Me.dbgShoppingList.CurrentRowIndex).Item("equip_id")
            Dim SQL As String = ""
            SQL &= "select price_id "
            SQL &= "from equipment "
            SQL &= "where equip_id = '" & eID & "' "
            If oDA.SendQuery(SQL, dt, ConnectString) < 1 OrElse _
               IsDBNull(dt.Rows(0).Item("Price_id")) Then
                MsgBox("Left click on the row you want to change to select it, and then right click to display popup menu.", MsgBoxStyle.Information)
                Exit Sub
            End If

            iPid = dt.Rows(0).Item("Price_id")

            dt.Reset()
            SQL = "select * from rental_rates where price_id = " & iPid
            If oDA.SendQuery(SQL, dt, ConnectString) < 1 Then
                Throw New System.Exception("Can't read rental rates table")
            End If
            hdRate = dt.Rows(0).Item("HalfDay")
            Dim newRow As String() = {dtList.Rows(iRow).Item("Equip_id"), _
                                    dtList.Rows(iRow).Item("equip_name"), _
                                    1, _
                                    "Half Day", _
                                    hdRate, _
                                    Now.ToString}
            oCG.AddRowToTable(dtList, newRow)
            oCG.BindDataTableToGrid(dtList, Me.dbgShoppingList)
            oCG.DisableAddNew(Me.dbgShoppingList, Me)
            ManualRecalc()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

    ''' <summary>
    '' here the user has chosen to change the time period type
    ''' of the rental.  So we have to go to the price table for 
    ''' the piece of equipment and plug in the price per unit of that equip
    ''' first, we have to get to the price table by picking up the equip_id
    ''' and using that to get the price_id
    ''' </summary>
    ''' <param name = "sender"></param>
    ''' <param name = "e"></param>
    Private Sub mnuDaily_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
Handles mnuDaily.Click, mnuHalfDay.Click, mnuMonth.Click, mnuWeek.Click, mnuWeekEnd.Click
        Dim eid As String
        Dim dt As New DataTable()
        Dim iPid As Integer
        Dim price As Decimal

        Try
            eid = dtList.Rows(Me.dbgShoppingList.CurrentRowIndex).Item("equip_id")
            Dim SQL As String = ""
            SQL &= "select price_id "
            SQL &= "from equipment "
            SQL &= "where equip_id = '" & eid & "' "
            If oDA.SendQuery(SQL, dt, ConnectString) < 1 OrElse _
               IsDBNull(dt.Rows(0).Item("Price_id")) Then
                MsgBox("Left click on the row you want to change to select it, and then right click to display popup menu.", MsgBoxStyle.Information)
                Exit Sub
            End If

            iPid = dt.Rows(0).Item("Price_id")

            dt.Reset()
            SQL = "select * from rental_rates where price_id = " & iPid
            If oDA.SendQuery(SQL, dt, ConnectString) < 1 Then
                Throw New System.Exception("Can't read rental rates table")
            End If

            Dim sNewRP As String
            Dim o As New MenuItem()
            o = CType(sender, MenuItem)
            sNewRP = o.Text

            With dt.Rows(0)
                Select Case sNewRP
                    Case DAILY : price = .Item("daily") : sNewRP = DAILY
                    Case WEEKLY : price = .Item("Weekly") : sNewRP = WEEKLY
                    Case MONTHLY : price = .Item("Monthly") : sNewRP = MONTHLY
                    Case WEEK_END : price = .Item("WeekEnd") : sNewRP = WEEK_END
                    Case HALF_DAY : price = .Item("HalfDay") : sNewRP = HALF_DAY
                    Case Else
                        price = 0
                End Select
            End With

            If price <> 0 Then
                dtList.Rows(Me.dbgShoppingList.CurrentRowIndex).Item("rental_period") = sNewRP
                dtList.Rows(Me.dbgShoppingList.CurrentRowIndex).Item("priceperunit") = price
            End If

        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try

    End Sub

    Private Sub btnManualRecalc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnManualRecalc.Click
        ManualRecalc()
    End Sub


    Private Sub btnCheckOut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCheckOut.Click
        CheckInHandler()
    End Sub

    Private Sub mnuPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPrint.Click
        PrintHandler()
    End Sub

    Private Sub mnuCheckIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCheckIn.Click
        CheckInHandler()
    End Sub

    Private Sub mnuCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCancel.Click
        CancelHandler()
    End Sub

    Private Sub btnOtherCharges_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOtherCharges.Click
        Dim oFrm As New frmMiscCheckInItems(dtList)
        oFrm.ShowDialog()
        oCG.BindDataTableToGrid(dtList, Me.dbgShoppingList)
        oCG.DisableAddNew(Me.dbgShoppingList, Me)
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

    Private Sub mnuShowPriceTable_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuShowPriceTable.Click
        Dim ofrm As New frmRentalRates()
        ofrm.ShowDialog()
    End Sub

    Private Sub txtNotes_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNotes.Enter
        ignoreKeyPreview = True
    End Sub

    Private Sub txtNotes_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNotes.Leave
        ignoreKeyPreview = False
    End Sub

    Private Sub mnuVoidInvoice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuVoidInvoice.Click
        Dim sMsg As String
        Dim iRV As Integer
        sMsg = "Are you absolutely sure that you want to VOID this Invoice?" & Chr(10)
        sMsg &= "" & Chr(10)
        iRV = MsgBox(sMsg, CType(292, Microsoft.VisualBasic.MsgBoxStyle), "Confirm Voiding Invoice")

        If iRV = 6 Then
            ' Yes Code goes here
        Else
            ' No code goes here
            Exit Sub
        End If
        voidInvoice = True
        PrintInvoice(Me.txtInvoiceID.Text)
        CheckIn()
        Me.Close()
    End Sub

    Private Sub dbgShoppingList_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dbgShoppingList.MouseUp

        Try
            Static busy As Boolean
            If busy Then Exit Sub
            busy = True
            Try
                listHitRow = dbgShoppingList.CurrentCell.RowNumber
                dbgShoppingList.Select(listHitRow)
            Catch
                MsgBox("Selected Row has been deleted.", MsgBoxStyle.Exclamation)
                Exit Sub
            End Try

            'If e.Button = MouseButtons.Right Then
            '   'If hti.Type = DataGrid.HitTestType.RowHeader Then
            '   Dim ocg As New CGrid()
            '   Dim miHitRow As Integer = ocg.SelectCkBoxRow(Me.dbgShoppingList, e)
            '   'End If
            'End If

            If e.Button = MouseButtons.Right Then
                Dim s As String = MNS(Me.dtList.Rows(Me.dbgShoppingList.CurrentCell.RowNumber).Item("rental_period"))
                If "Daily_Hourly_Half Day_Weekly_Monthly_Week End".IndexOf(s) > -1 Then
                    Dim oFrm As New frmCheckinMods()
                    oFrm.frm = Me
                    oFrm.ShowDialog()
                    ManualRecalc()
                ElseIf s.Trim.ToUpper = "SALE" Then
                    Dim ofrm As New frmCkInModifySaleItem()
                    ofrm.frm = Me
                    ofrm.ShowDialog()
                    ManualRecalc()
                    'Dim sMsg As String
                    'sMsg = "If you want to modify the sale item selected, " & Chr(10)
                    'sMsg &= "press the Delete Button, which will add the items" & Chr(10)
                    'sMsg &= "back to inventory and then add the number of " & Chr(10)
                    'sMsg &= "items actually used by clicking thee Labor, Fuel, " & Chr(10)
                    'sMsg &= "and Sale Items Button." & Chr(10)
                    'sMsg &= "" & Chr(10)
                    'MsgBox(sMsg, CType(64, Microsoft.VisualBasic.MsgBoxStyle), "Delete Sale Item")

                End If
            End If
            busy = False
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

    Private Sub textManualPickup_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles textManualPickup.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Exit Sub
        End If
        e.Handled = CkKeyPressNumeric(Asc(Val(e.KeyChar)), textManualPickup) = 0
    End Sub
    Private Sub textManualPickup_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles textManualPickup.KeyDown
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
        If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
        If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
    End Sub
    Private Sub textManualPickup_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles textManualPickup.Enter
        textManualPickup.Text = UnFmt_T_B(textManualPickup)
        textManualPickup.SelectionStart = 0
        textManualPickup.SelectionLength = textManualPickup.Text.Trim.Length
    End Sub
    Private Sub textManualPickup_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles textManualPickup.Leave
        textManualPickup.Text = Fmt_T_B(textManualPickup)
    End Sub

    Private Sub cboDeliveryDistance_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDeliveryDistance.SelectedIndexChanged
        Dim s As String = Me.cboDeliveryDistance.Text
        Me.textManualPickup.Text = FormatCurrency(UnFormat(s.Substring(0, s.IndexOf(" "))))
        ManualRecalc()
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
    Private Sub textDriversLicence_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles textDriversLicence.Enter
        textDriversLicence.SelectionStart = 0
        textDriversLicence.SelectionLength = textDriversLicence.Text.Trim.Length
    End Sub

    Private Sub dtpCkInDateReset_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpCkInDateReset.ValueChanged
        With dtpCkInDateReset
            Me.lblCkInDate.Text = .Value.ToString
        End With
    End Sub

    Private Sub cbEmployees_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbEmployees.SelectedIndexChanged
        If Me.cbEmployees.Text.Trim.Length > 0 Then
            SaveSetting(RENTALPRO, SETTINGS, "CKINEMP", Me.cbEmployees.Text.Trim)
        End If
    End Sub

    Private Sub btnRerunAutoCalc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRerunAutoCalc.Click
        Me.lblCkInDate.Text = oCI.AutoCalcCheckIn(CType(Me.lblCkInDate.Text, DateTime)).ToString
        Me.textElapsedTimeHours.Text = modAutoCalc.ElapsedTime
        ManualRecalc()
        Me.TotalTheItems()
    End Sub


#End Region

End Class
