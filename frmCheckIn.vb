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
Public Class frmCheckIn
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
   Friend WithEvents txtExpMonth As System.Windows.Forms.TextBox
   Friend WithEvents txtExpYr As System.Windows.Forms.TextBox
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCheckIn))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.dbgShoppingList = New System.Windows.Forms.DataGrid()
        Me.mnuContext = New System.Windows.Forms.ContextMenu()
        Me.MenuItem1 = New System.Windows.Forms.MenuItem()
        Me.mnuDaily = New System.Windows.Forms.MenuItem()
        Me.mnuWeek = New System.Windows.Forms.MenuItem()
        Me.mnuWeekEnd = New System.Windows.Forms.MenuItem()
        Me.mnuHalfDay = New System.Windows.Forms.MenuItem()
        Me.mnuMonth = New System.Windows.Forms.MenuItem()
        Me.mnuHourly = New System.Windows.Forms.MenuItem()
        Me.mnuAddHalfDay = New System.Windows.Forms.MenuItem()
        Me.cmdPrintContract = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
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
        Me.btnCheckOut = New System.Windows.Forms.Button()
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
        Me.btnOtherCharges = New System.Windows.Forms.Button()
        Me.txtNotes = New System.Windows.Forms.TextBox()
        Me.lblNotes = New System.Windows.Forms.Label()
        Me.txtCardId = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.txtExpMonth = New System.Windows.Forms.TextBox()
        Me.txtExpYr = New System.Windows.Forms.TextBox()
        CType(Me.dbgShoppingList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Frame2.SuspendLayout()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblFieldLable, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblLabels, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'dbgShoppingList
        '
        Me.dbgShoppingList.AllowSorting = False
        Me.dbgShoppingList.AlternatingBackColor = System.Drawing.Color.MediumSeaGreen
        Me.dbgShoppingList.CaptionVisible = False
        Me.dbgShoppingList.ContextMenu = Me.mnuContext
        Me.dbgShoppingList.DataMember = ""
        Me.dbgShoppingList.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.HelpProvider1.SetHelpString(Me.dbgShoppingList, "To change period or add a half day, right click on the row you desire to change.")
        Me.dbgShoppingList.Location = New System.Drawing.Point(8, 172)
        Me.dbgShoppingList.Name = "dbgShoppingList"
        Me.HelpProvider1.SetShowHelp(Me.dbgShoppingList, True)
        Me.dbgShoppingList.Size = New System.Drawing.Size(591, 148)
        Me.dbgShoppingList.TabIndex = 44
        Me.ToolTip1.SetToolTip(Me.dbgShoppingList, "Right click on desired row to change period or add half day")
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
        'cmdPrintContract
        '
        Me.cmdPrintContract.BackColor = System.Drawing.SystemColors.Control
        Me.cmdPrintContract.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdPrintContract.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPrintContract.ForeColor = System.Drawing.SystemColors.ControlText
        Me.HelpProvider1.SetHelpString(Me.cmdPrintContract, "Print contract and put on rent.")
        Me.cmdPrintContract.Location = New System.Drawing.Point(313, 327)
        Me.cmdPrintContract.Name = "cmdPrintContract"
        Me.cmdPrintContract.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.HelpProvider1.SetShowHelp(Me.cmdPrintContract, True)
        Me.cmdPrintContract.Size = New System.Drawing.Size(107, 27)
        Me.cmdPrintContract.TabIndex = 0
        Me.cmdPrintContract.Text = "&Print Contract"
        Me.cmdPrintContract.UseVisualStyleBackColor = False
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.HelpProvider1.SetHelpString(Me.cmdCancel, "Cancel the checkout process.")
        Me.cmdCancel.Location = New System.Drawing.Point(313, 390)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.HelpProvider1.SetShowHelp(Me.cmdCancel, True)
        Me.cmdCancel.Size = New System.Drawing.Size(107, 27)
        Me.cmdCancel.TabIndex = 130
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
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
        Me.Frame2.Location = New System.Drawing.Point(8, 1)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(589, 167)
        Me.Frame2.TabIndex = 16
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Customer"
        '
        'txtTaxID
        '
        Me.txtTaxID.Location = New System.Drawing.Point(88, 140)
        Me.txtTaxID.Name = "txtTaxID"
        Me.txtTaxID.Size = New System.Drawing.Size(120, 20)
        Me.txtTaxID.TabIndex = 60
        Me.txtTaxID.Tag = "(No Auto Formatting)"
        Me.ToolTip1.SetToolTip(Me.txtTaxID, "A personal SSN is not a valid Federal Taxid")
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(45, 142)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(36, 14)
        Me.Label9.TabIndex = 59
        Me.Label9.Text = "Tax ID"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(363, 139)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(33, 14)
        Me.Label7.TabIndex = 58
        Me.Label7.Text = "Inv ID"
        '
        'txtInvoiceID
        '
        Me.txtInvoiceID.Location = New System.Drawing.Point(399, 139)
        Me.txtInvoiceID.Name = "txtInvoiceID"
        Me.txtInvoiceID.Size = New System.Drawing.Size(136, 20)
        Me.txtInvoiceID.TabIndex = 57
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(355, 120)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(41, 14)
        Me.Label6.TabIndex = 56
        Me.Label6.Text = "Cust ID"
        '
        'txtCustomerID
        '
        Me.txtCustomerID.Location = New System.Drawing.Point(399, 118)
        Me.txtCustomerID.Name = "txtCustomerID"
        Me.txtCustomerID.Size = New System.Drawing.Size(136, 20)
        Me.txtCustomerID.TabIndex = 55
        '
        'lblLine
        '
        Me.lblLine.BackColor = System.Drawing.Color.Gray
        Me.lblLine.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLine.Location = New System.Drawing.Point(1, 62)
        Me.lblLine.Name = "lblLine"
        Me.lblLine.Size = New System.Drawing.Size(586, 1)
        Me.lblLine.TabIndex = 53
        '
        'txtPONbr
        '
        Me.txtPONbr.Location = New System.Drawing.Point(248, 116)
        Me.txtPONbr.Name = "txtPONbr"
        Me.txtPONbr.Size = New System.Drawing.Size(104, 20)
        Me.txtPONbr.TabIndex = 52
        Me.txtPONbr.Tag = "(No Auto Formatting)"
        '
        'lblPONbr
        '
        Me.lblPONbr.AutoSize = True
        Me.lblPONbr.Location = New System.Drawing.Point(208, 119)
        Me.lblPONbr.Name = "lblPONbr"
        Me.lblPONbr.Size = New System.Drawing.Size(30, 14)
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
        Me.txtShipZip.Location = New System.Drawing.Point(459, 92)
        Me.txtShipZip.MaxLength = 0
        Me.txtShipZip.Name = "txtShipZip"
        Me.txtShipZip.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtShipZip.Size = New System.Drawing.Size(91, 21)
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
        Me.txtShipState.Location = New System.Drawing.Point(399, 92)
        Me.txtShipState.MaxLength = 2
        Me.txtShipState.Name = "txtShipState"
        Me.txtShipState.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtShipState.Size = New System.Drawing.Size(27, 21)
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
        Me.txtShipCity.Location = New System.Drawing.Point(399, 68)
        Me.txtShipCity.MaxLength = 0
        Me.txtShipCity.Name = "txtShipCity"
        Me.txtShipCity.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtShipCity.Size = New System.Drawing.Size(173, 21)
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
        Me.Label3.Location = New System.Drawing.Point(435, 94)
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
        Me.Label4.Location = New System.Drawing.Point(361, 94)
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
        Me.Label5.Location = New System.Drawing.Point(361, 71)
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
        Me.txtContactName.Location = New System.Drawing.Point(89, 116)
        Me.txtContactName.MaxLength = 0
        Me.txtContactName.Name = "txtContactName"
        Me.txtContactName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtContactName.Size = New System.Drawing.Size(113, 21)
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
        Me.txtPostalCode.Location = New System.Drawing.Point(460, 37)
        Me.txtPostalCode.MaxLength = 0
        Me.txtPostalCode.Name = "txtPostalCode"
        Me.txtPostalCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPostalCode.Size = New System.Drawing.Size(91, 21)
        Me.txtPostalCode.TabIndex = 7
        '
        'txtState
        '
        Me.txtState.AcceptsReturn = True
        Me.txtState.BackColor = System.Drawing.SystemColors.Window
        Me.txtState.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtState.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtState.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtState.Location = New System.Drawing.Point(400, 37)
        Me.txtState.MaxLength = 2
        Me.txtState.Name = "txtState"
        Me.txtState.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtState.Size = New System.Drawing.Size(27, 21)
        Me.txtState.TabIndex = 6
        '
        'txtShipAddress1
        '
        Me.txtShipAddress1.AcceptsReturn = True
        Me.txtShipAddress1.BackColor = System.Drawing.SystemColors.Window
        Me.txtShipAddress1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtShipAddress1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtShipAddress1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtShipAddress1.Location = New System.Drawing.Point(90, 92)
        Me.txtShipAddress1.MaxLength = 0
        Me.txtShipAddress1.Name = "txtShipAddress1"
        Me.txtShipAddress1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtShipAddress1.Size = New System.Drawing.Size(261, 21)
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
        Me.txtCity.Location = New System.Drawing.Point(400, 13)
        Me.txtCity.MaxLength = 0
        Me.txtCity.Name = "txtCity"
        Me.txtCity.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCity.Size = New System.Drawing.Size(173, 21)
        Me.txtCity.TabIndex = 5
        '
        'txtShipToCustomer
        '
        Me.txtShipToCustomer.AcceptsReturn = True
        Me.txtShipToCustomer.BackColor = System.Drawing.SystemColors.Window
        Me.txtShipToCustomer.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtShipToCustomer.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtShipToCustomer.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtShipToCustomer.Location = New System.Drawing.Point(90, 68)
        Me.txtShipToCustomer.MaxLength = 0
        Me.txtShipToCustomer.Name = "txtShipToCustomer"
        Me.txtShipToCustomer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtShipToCustomer.Size = New System.Drawing.Size(261, 21)
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
        Me.txtBillingAddress1.Location = New System.Drawing.Point(90, 37)
        Me.txtBillingAddress1.MaxLength = 0
        Me.txtBillingAddress1.Name = "txtBillingAddress1"
        Me.txtBillingAddress1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtBillingAddress1.Size = New System.Drawing.Size(261, 21)
        Me.txtBillingAddress1.TabIndex = 2
        '
        'txtCompanyName
        '
        Me.txtCompanyName.AcceptsReturn = True
        Me.txtCompanyName.BackColor = System.Drawing.SystemColors.Window
        Me.txtCompanyName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCompanyName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCompanyName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCompanyName.Location = New System.Drawing.Point(90, 13)
        Me.txtCompanyName.MaxLength = 0
        Me.txtCompanyName.Name = "txtCompanyName"
        Me.txtCompanyName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCompanyName.Size = New System.Drawing.Size(261, 21)
        Me.txtCompanyName.TabIndex = 1
        '
        'lblContact
        '
        Me.lblContact.AutoSize = True
        Me.lblContact.BackColor = System.Drawing.SystemColors.Control
        Me.lblContact.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblContact.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblContact.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblContact.Location = New System.Drawing.Point(6, 118)
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
        Me._lblLabels_6.Location = New System.Drawing.Point(436, 39)
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
        Me._lblLabels_5.Location = New System.Drawing.Point(362, 39)
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
        Me._lblLabels_4.Location = New System.Drawing.Point(362, 16)
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
        Me.lblLabels.SetIndex(Me._lblLabels_3, CType(3, Short))
        Me._lblLabels_3.Location = New System.Drawing.Point(8, 95)
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
        Me.lblLabels.SetIndex(Me._lblLabels_2, CType(2, Short))
        Me._lblLabels_2.Location = New System.Drawing.Point(8, 71)
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
        Me._lblLabels_1.Location = New System.Drawing.Point(8, 41)
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
        Me._lblLabels_0.Location = New System.Drawing.Point(8, 17)
        Me._lblLabels_0.Name = "_lblLabels_0"
        Me._lblLabels_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_0.Size = New System.Drawing.Size(82, 14)
        Me._lblLabels_0.TabIndex = 17
        Me._lblLabels_0.Text = "CompanyName:"
        '
        'txtDelivery
        '
        Me.txtDelivery.Location = New System.Drawing.Point(488, 376)
        Me.txtDelivery.Name = "txtDelivery"
        Me.txtDelivery.Size = New System.Drawing.Size(91, 20)
        Me.txtDelivery.TabIndex = 45
        Me.txtDelivery.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblDelivery
        '
        Me.lblDelivery.AutoSize = True
        Me.lblDelivery.Location = New System.Drawing.Point(440, 378)
        Me.lblDelivery.Name = "lblDelivery"
        Me.lblDelivery.Size = New System.Drawing.Size(46, 14)
        Me.lblDelivery.TabIndex = 46
        Me.lblDelivery.Text = "Delivery"
        '
        'lblTax
        '
        Me.lblTax.AutoSize = True
        Me.lblTax.Location = New System.Drawing.Point(432, 403)
        Me.lblTax.Name = "lblTax"
        Me.lblTax.Size = New System.Drawing.Size(54, 14)
        Me.lblTax.TabIndex = 47
        Me.lblTax.Text = "Sales Tax"
        '
        'txtSalesTax
        '
        Me.txtSalesTax.Location = New System.Drawing.Point(488, 400)
        Me.txtSalesTax.Name = "txtSalesTax"
        Me.txtSalesTax.Size = New System.Drawing.Size(91, 20)
        Me.txtSalesTax.TabIndex = 48
        Me.txtSalesTax.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtDeposit
        '
        Me.txtDeposit.Location = New System.Drawing.Point(488, 424)
        Me.txtDeposit.Name = "txtDeposit"
        Me.txtDeposit.Size = New System.Drawing.Size(91, 20)
        Me.txtDeposit.TabIndex = 49
        Me.txtDeposit.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblDeposit
        '
        Me.lblDeposit.AutoSize = True
        Me.lblDeposit.Location = New System.Drawing.Point(440, 426)
        Me.lblDeposit.Name = "lblDeposit"
        Me.lblDeposit.Size = New System.Drawing.Size(43, 14)
        Me.lblDeposit.TabIndex = 50
        Me.lblDeposit.Text = "Deposit"
        '
        'lblTotal
        '
        Me.lblTotal.AutoSize = True
        Me.lblTotal.Location = New System.Drawing.Point(456, 448)
        Me.lblTotal.Name = "lblTotal"
        Me.lblTotal.Size = New System.Drawing.Size(29, 14)
        Me.lblTotal.TabIndex = 51
        Me.lblTotal.Text = "Total"
        Me.lblTotal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtTotal
        '
        Me.txtTotal.Location = New System.Drawing.Point(488, 448)
        Me.txtTotal.Name = "txtTotal"
        Me.txtTotal.Size = New System.Drawing.Size(91, 20)
        Me.txtTotal.TabIndex = 52
        Me.txtTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblItemTotal
        '
        Me.lblItemTotal.AutoSize = True
        Me.lblItemTotal.Location = New System.Drawing.Point(432, 352)
        Me.lblItemTotal.Name = "lblItemTotal"
        Me.lblItemTotal.Size = New System.Drawing.Size(51, 14)
        Me.lblItemTotal.TabIndex = 53
        Me.lblItemTotal.Text = "Item Total"
        '
        'txtItemTotal
        '
        Me.txtItemTotal.Location = New System.Drawing.Point(488, 352)
        Me.txtItemTotal.Name = "txtItemTotal"
        Me.txtItemTotal.Size = New System.Drawing.Size(91, 20)
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
        Me.GroupBox1.Location = New System.Drawing.Point(8, 322)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(272, 79)
        Me.GroupBox1.TabIndex = 55
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Payment Arrangement"
        '
        'optCash
        '
        Me.optCash.Location = New System.Drawing.Point(137, 56)
        Me.optCash.Name = "optCash"
        Me.optCash.Size = New System.Drawing.Size(96, 16)
        Me.optCash.TabIndex = 6
        Me.optCash.Text = "Paid by Cash"
        '
        'optLeftCardNumber
        '
        Me.optLeftCardNumber.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
        Me.optLeftBlankCheck.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
        Me.optPaidByCheck.Checked = True
        Me.optPaidByCheck.Location = New System.Drawing.Point(8, 16)
        Me.optPaidByCheck.Name = "optPaidByCheck"
        Me.optPaidByCheck.Size = New System.Drawing.Size(94, 16)
        Me.optPaidByCheck.TabIndex = 0
        Me.optPaidByCheck.TabStop = True
        Me.optPaidByCheck.Text = "Paid by Check"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 408)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(106, 14)
        Me.Label2.TabIndex = 56
        Me.Label2.Text = "Check/Card  Number"
        '
        'txtCheckNumber
        '
        Me.HelpProvider1.SetHelpString(Me.txtCheckNumber, "Enter check or credit card number if left.")
        Me.txtCheckNumber.Location = New System.Drawing.Point(8, 424)
        Me.txtCheckNumber.Name = "txtCheckNumber"
        Me.HelpProvider1.SetShowHelp(Me.txtCheckNumber, True)
        Me.txtCheckNumber.Size = New System.Drawing.Size(112, 20)
        Me.txtCheckNumber.TabIndex = 57
        '
        'lblAmtPaid
        '
        Me.lblAmtPaid.Location = New System.Drawing.Point(431, 468)
        Me.lblAmtPaid.Name = "lblAmtPaid"
        Me.lblAmtPaid.Size = New System.Drawing.Size(56, 25)
        Me.lblAmtPaid.TabIndex = 58
        Me.lblAmtPaid.Text = "Amt Paid At Rental"
        '
        'txtAmtPaid
        '
        Me.HelpProvider1.SetHelpString(Me.txtAmtPaid, "Enter 0 if no payment made.")
        Me.txtAmtPaid.Location = New System.Drawing.Point(488, 472)
        Me.txtAmtPaid.Name = "txtAmtPaid"
        Me.HelpProvider1.SetShowHelp(Me.txtAmtPaid, True)
        Me.txtAmtPaid.Size = New System.Drawing.Size(91, 20)
        Me.txtAmtPaid.TabIndex = 59
        Me.txtAmtPaid.Tag = "(No Auto Formatting)"
        Me.txtAmtPaid.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtBalDue
        '
        Me.txtBalDue.Location = New System.Drawing.Point(488, 533)
        Me.txtBalDue.Name = "txtBalDue"
        Me.txtBalDue.Size = New System.Drawing.Size(91, 20)
        Me.txtBalDue.TabIndex = 61
        Me.txtBalDue.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblBalDue
        '
        Me.lblBalDue.Location = New System.Drawing.Point(438, 533)
        Me.lblBalDue.Name = "lblBalDue"
        Me.lblBalDue.Size = New System.Drawing.Size(45, 25)
        Me.lblBalDue.TabIndex = 60
        Me.lblBalDue.Text = "Balance Due"
        '
        'lblLine2
        '
        Me.lblLine2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLine2.Location = New System.Drawing.Point(488, 525)
        Me.lblLine2.Name = "lblLine2"
        Me.lblLine2.Size = New System.Drawing.Size(95, 2)
        Me.lblLine2.TabIndex = 62
        '
        'txtAmtPaidAtCkIn
        '
        Me.HelpProvider1.SetHelpString(Me.txtAmtPaidAtCkIn, "Enter 0 if no payment made.")
        Me.txtAmtPaidAtCkIn.Location = New System.Drawing.Point(488, 496)
        Me.txtAmtPaidAtCkIn.Name = "txtAmtPaidAtCkIn"
        Me.HelpProvider1.SetShowHelp(Me.txtAmtPaidAtCkIn, True)
        Me.txtAmtPaidAtCkIn.Size = New System.Drawing.Size(91, 20)
        Me.txtAmtPaidAtCkIn.TabIndex = 133
        Me.txtAmtPaidAtCkIn.Tag = "(No Auto Formatting)"
        Me.txtAmtPaidAtCkIn.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'btnCheckOut
        '
        Me.btnCheckOut.BackColor = System.Drawing.SystemColors.Control
        Me.btnCheckOut.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnCheckOut.Enabled = False
        Me.btnCheckOut.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCheckOut.ForeColor = System.Drawing.SystemColors.ControlText
        Me.HelpProvider1.SetHelpString(Me.btnCheckOut, "Print contract and put on rent.")
        Me.btnCheckOut.Location = New System.Drawing.Point(313, 358)
        Me.btnCheckOut.Name = "btnCheckOut"
        Me.btnCheckOut.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.HelpProvider1.SetShowHelp(Me.btnCheckOut, True)
        Me.btnCheckOut.Size = New System.Drawing.Size(107, 27)
        Me.btnCheckOut.TabIndex = 135
        Me.btnCheckOut.Text = "Check &In"
        Me.btnCheckOut.UseVisualStyleBackColor = False
        '
        'btnManualRecalc
        '
        Me.btnManualRecalc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnManualRecalc.Location = New System.Drawing.Point(313, 462)
        Me.btnManualRecalc.Name = "btnManualRecalc"
        Me.btnManualRecalc.Size = New System.Drawing.Size(107, 27)
        Me.btnManualRecalc.TabIndex = 132
        Me.btnManualRecalc.Text = "&Manual Recalc"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(433, 494)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(54, 36)
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
        'btnOtherCharges
        '
        Me.btnOtherCharges.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOtherCharges.Location = New System.Drawing.Point(313, 423)
        Me.btnOtherCharges.Name = "btnOtherCharges"
        Me.btnOtherCharges.Size = New System.Drawing.Size(107, 34)
        Me.btnOtherCharges.TabIndex = 136
        Me.btnOtherCharges.Text = "&Labor, Fuel, Tools, & Supplies"
        '
        'txtNotes
        '
        Me.txtNotes.Location = New System.Drawing.Point(8, 471)
        Me.txtNotes.MaxLength = 255
        Me.txtNotes.Multiline = True
        Me.txtNotes.Name = "txtNotes"
        Me.txtNotes.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtNotes.Size = New System.Drawing.Size(269, 95)
        Me.txtNotes.TabIndex = 137
        '
        'lblNotes
        '
        Me.lblNotes.AutoSize = True
        Me.lblNotes.Location = New System.Drawing.Point(8, 453)
        Me.lblNotes.Name = "lblNotes"
        Me.lblNotes.Size = New System.Drawing.Size(35, 14)
        Me.lblNotes.TabIndex = 138
        Me.lblNotes.Text = "Notes"
        '
        'txtCardId
        '
        Me.txtCardId.Location = New System.Drawing.Point(236, 424)
        Me.txtCardId.MaxLength = 4
        Me.txtCardId.Name = "txtCardId"
        Me.txtCardId.Size = New System.Drawing.Size(40, 20)
        Me.txtCardId.TabIndex = 152
        Me.txtCardId.Tag = "(No Auto Formatting)"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(228, 408)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(42, 14)
        Me.Label10.TabIndex = 151
        Me.Label10.Text = "Card ID"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(180, 408)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(40, 14)
        Me.Label11.TabIndex = 150
        Me.Label11.Text = "Exp Yr"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(122, 408)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(48, 14)
        Me.Label12.TabIndex = 147
        Me.Label12.Text = "Exp Mon"
        '
        'txtExpMonth
        '
        Me.txtExpMonth.Location = New System.Drawing.Point(128, 424)
        Me.txtExpMonth.Name = "txtExpMonth"
        Me.txtExpMonth.Size = New System.Drawing.Size(32, 20)
        Me.txtExpMonth.TabIndex = 153
        '
        'txtExpYr
        '
        Me.txtExpYr.Location = New System.Drawing.Point(180, 424)
        Me.txtExpYr.Name = "txtExpYr"
        Me.txtExpYr.Size = New System.Drawing.Size(32, 20)
        Me.txtExpYr.TabIndex = 154
        '
        'frmCheckIn
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(610, 571)
        Me.Controls.Add(Me.txtExpYr)
        Me.Controls.Add(Me.txtExpMonth)
        Me.Controls.Add(Me.txtCardId)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.lblNotes)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lblItemTotal)
        Me.Controls.Add(Me.lblTotal)
        Me.Controls.Add(Me.lblDeposit)
        Me.Controls.Add(Me.lblTax)
        Me.Controls.Add(Me.lblDelivery)
        Me.Controls.Add(Me.txtNotes)
        Me.Controls.Add(Me.btnOtherCharges)
        Me.Controls.Add(Me.btnCheckOut)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.lblAmtPaid)
        Me.Controls.Add(Me.txtAmtPaidAtCkIn)
        Me.Controls.Add(Me.btnManualRecalc)
        Me.Controls.Add(Me.lblLine2)
        Me.Controls.Add(Me.txtBalDue)
        Me.Controls.Add(Me.lblBalDue)
        Me.Controls.Add(Me.txtAmtPaid)
        Me.Controls.Add(Me.txtCheckNumber)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.txtItemTotal)
        Me.Controls.Add(Me.txtTotal)
        Me.Controls.Add(Me.txtDeposit)
        Me.Controls.Add(Me.txtSalesTax)
        Me.Controls.Add(Me.txtDelivery)
        Me.Controls.Add(Me.dbgShoppingList)
        Me.Controls.Add(Me.cmdPrintContract)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.Frame2)
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.HelpButton = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(141, 141)
        Me.MaximizeBox = False
        Me.Menu = Me.MainMenu1
        Me.MinimizeBox = False
        Me.Name = "frmCheckIn"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Customer CheckIn"
        CType(Me.dbgShoppingList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Frame2.ResumeLayout(False)
        Me.Frame2.PerformLayout()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblFieldLable, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblLabels, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
   Public dtFC As DataTable
   Dim mbFormLoading As Boolean = True
   Private oDA As CDataAccess
   Private m_CurrentInvoice As Integer
   Private dtList As New DataTable()
   Private m_bPrintedInvoice As Boolean = False
   Private oCG As New CGrid()
   Private origItemCount As Integer = 0
   Private ignoreKeyPreview As Boolean
   'Public MyCombo As New ComboBox()
   Private voidInvoice As Boolean = False
#If Not Reliable Then
   Private Sub frmCheckIn_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
      Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
      If KeyAscii = 27 Then
         Me.Close()
         System.Windows.Forms.Application.DoEvents()
      End If
      'If Not ignoreKeyPreview Then
      '   Select Case UCase(Chr(KeyAscii))
      '      Case "D"
      '         'SendKeys.Send("%D")
      '         'Me.cmdDelete_Click(New Object(), New Object())
      '      Case "P"
      '         Me.cmdPrintContract_Click(New Object(), New Object())
      '      Case "C"
      '         Me.cmdCancel_Click(New Object(), New Object())
      '      Case Else
      '         Exit Sub
      '   End Select
      '   '   KeyAscii = 0
      'End If

      If KeyAscii = 0 Then
         eventArgs.Handled = True
      End If
   End Sub

   Private Function VerifySettingsForCkOut() As Boolean
      ' ck to see the correct options are turned on
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

      Try
         If Not voidInvoice Then
            If Not Me.VerifySettingsForCkOut() Then Exit Sub
         End If

         iInvID = Val(Me.txtInvoiceID.Text)


         ' First, we mark the items off rent
         For i = 0 To dtList.Rows.Count - 1
            With dtList.Rows(i)
               If .Item("rental_period") <> SALE AndAlso _
                  .Item("equip_id") <> "Labor" AndAlso _
                  .Item("equip_id") <> "Fuel" AndAlso _
                  .Item("equip_id") <> "Misc" Then
                  SQL = "update equipment "
                  SQL &= "set rented_date = Null, "
                  SQL &= "available='YES', "
                  SQL &= "renting_company_id = Null, "
                  SQL &= "available_date = #" & Now.ToString & "# "
                  SQL &= "where equip_id = '" & .Item("equip_id") & "' "
                  iRows = oDA.SendActionSql(SQL, ConnectString, sErr)
                  'If iRows <> 1 Then
                  '   MsgBox("Equipment: " & CType(.Item("equip_id"), String) & " could not be taken off rent", MsgBoxStyle.Critical)
                  'End If
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
                  SQL &= .Item("quantity") & ", "
                  SQL &= .Item("priceperunit") & ", "
                  SQL &= "'" & .Item("equip_id") & "', "
                  SQL &= "'" & .Item("equip_name") & "', "
                  SQL &= "#" & Now.ToString & "#, "
                  SQL &= "#" & Now.ToString & "#, "
                  SQL &= "15, "
                  SQL &= "0, "
                  SQL &= "'N/A', "
                  SQL &= "0, "
                  SQL &= "'Rent/Sale Item')"
               Else
                  SQL = "Update invoice_details "
                  SQL &= "set returned_date = #" & Now.ToString & "#, "
                  SQL &= "priceperunit = " & .Item("priceperunit") & ", "
                  SQL &= "quantity = " & .Item("quantity") & ", "
                  SQL &= "rental_period = '" & .Item("rental_period") & "' "
                  SQL &= "where invoiceid = " & Val(Me.txtInvoiceID.Text) & " "
                  SQL &= "and record_type = 15 "
                  SQL &= "and equip_id = '" & .Item("equip_id") & "'"
               End If
               If oDA.SendActionSql(SQL, ConnectString, sErr) = 0 Then
                  MsgBox("Update of Invoice Detail for invoice: " & Me.txtInvoiceID.Text & _
                     " EquipID: " & .Item("equip_id") & " failed.", MsgBoxStyle.Critical)
               End If

            End With
         Next

         ' delete all of the invoice total items
         SQL = "delete from invoice_details "
         SQL &= "where invoiceid = " & Val(Me.txtInvoiceID.Text) & " "
         SQL &= "and record_type <> 15 "
         oDA.SendActionSql(SQL, ConnectString, sErr)

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
               iRows = oDA.SendActionSql(SQL, ConnectString, sErr)
               If iRows < 1 Then
                  MsgBox("Invoice Detail Update Failure: " & Chr(10) & _
                     sErr & Chr(10) & SQL, MsgBoxStyle.Critical)
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
               iRows = oDA.SendActionSql(SQL, ConnectString, sErr)
               If iRows < 1 Then
                  MsgBox("Invoice Detail Update Failure: " & Chr(10) & _
                     sErr & Chr(10) & SQL, MsgBoxStyle.Critical)
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
            iRows = oDA.SendActionSql(SQL, ConnectString, sErr)
            If iRows < 1 Then
               MsgBox("Invoice Detail Update Failure: " & Chr(10) & _
                  sErr & Chr(10) & SQL, MsgBoxStyle.Critical)
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
               iRows = oDA.SendActionSql(SQL, ConnectString, sErr)
               If iRows < 1 Then
                  MsgBox("Invoice Detail Update Failure: " & Chr(10) & _
                     sErr & Chr(10) & SQL, MsgBoxStyle.Critical)
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
               iRows = oDA.SendActionSql(SQL, ConnectString, sErr)
               If iRows < 1 Then
                  MsgBox("Invoice Detail Update Failure: " & Chr(10) & _
                     sErr & Chr(10) & SQL, MsgBoxStyle.Critical)
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
            SQL &= "notes = '" & Me.txtNotes.Text & "' "
         Else
            SQL = "update invoices set notes = 'INVOICE IS VOID', "
            SQL &= "balancedue = 0, status = 'CLOSED' "
         End If
         SQL &= "where invoiceid = " & iInvID
         iRows = oDA.SendActionSql(SQL, ConnectString, sErr)


      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

   Private Sub PrintInvoice(ByVal InvoiceId As Integer)
      ' format the print line
      Dim ps As New System.Text.StringBuilder()
      Dim decEP As Decimal
      Dim SQL As String
      Dim i As Integer
      Dim dt As New DataTable()
#If Reliable Then
      Dim oPD As New CReliablePrint(Me)
#Else
      'Dim oPD As New CPioneerPrint(Me)
#End If
      Dim oUtil As New CUtilities()
      Dim decTotal As Decimal
      Dim sName As String
      ' get customer data and print

      Try

         For i = 0 To dtList.Rows.Count - 1
            With dtList.Rows(i)
               ' qty
               ps.Append(CType(.Item("Quantity"), String).PadLeft(4))
               ' skip 3 spaces and print id - name
               ps.Append(Space(3) & CType(.Item("Equip_Id"), String).PadRight(10) & " - ")
               sName = .Item("equip_name")
               If sName.Length > 29 Then
                  sName = sName.Substring(0, 29)
               End If
               ' qty
               ps.Append(sName.PadRight(30))

               ' rental period (Daily...)
               ps.Append(Space(9) & CType(.Item("Rental_Period"), String).PadRight(10))
               ' price per unit
               ps.Append(Format(.Item("priceperunit"), "#,##0.00").PadLeft(10))
               decEP = .Item("PriceperUnit") * _
                      .Item("Quantity") ' + _
               ps.Append(Space(2) & Format(decEP, "#,##0.00").PadLeft(10) & vbCrLf)

            End With
         Next i

         Const DTSP = 70
         If Not voidInvoice Then
            ' print the totals
            ps.Append(vbCrLf & Space(DTSP) & "Item Total".PadRight(11) & Me.txtItemTotal.Text.PadLeft(10) & vbCrLf)
            If UnFormat(Me.txtDeposit.Text) <> 0 Then
               ps.Append(vbCrLf & Space(DTSP) & "Deposit".PadRight(11) & Me.txtDeposit.Text.PadLeft(10) & vbCrLf)
            End If

            If UnFormat(Me.txtSalesTax.Text) > 0 Then
               ps.Append(vbCrLf & Space(DTSP) & "Sales Tax".PadRight(11) & Me.txtSalesTax.Text.PadLeft(10) & vbCrLf)
            End If

            If UnFormat(Me.txtDelivery.Text) <> 0 Then
               ps.Append(vbCrLf & Space(DTSP) & "Delivery".PadRight(11) & Me.txtDelivery.Text.PadLeft(10) & vbCrLf)
            End If

            ps.Append(vbCrLf & Space(DTSP) & "Total".PadRight(11) & Me.txtTotal.Text.PadLeft(10) & vbCrLf)
            If UnFormat(Me.txtAmtPaid.Text) <> 0 Then
               ps.Append(vbCrLf & Space(DTSP) & "Paid/CkOut".PadRight(11) & Me.txtAmtPaid.Text.PadLeft(10) & vbCrLf)
            End If

            If UnFormat(Me.txtAmtPaidAtCkIn.Text) > 0 Then
               ps.Append(vbCrLf & Space(DTSP) & "Paid/CkIn".PadRight(11) & Me.txtAmtPaidAtCkIn.Text.PadLeft(10) & vbCrLf)
            End If
            If UnFormat(Me.txtBalDue.Text) < 0 Then
               Dim valu As Decimal = UnFormat(Me.txtBalDue.Text)
               ps.Append(vbCrLf & Space(DTSP) & "Refund Due".PadRight(11) & _
                  FormatCurrency(valu * -1).PadLeft(10) & vbCrLf)
            Else
               ps.Append(vbCrLf & Space(DTSP) & "Bal Due".PadRight(11) & Me.txtBalDue.Text.PadLeft(10) & vbCrLf)
            End If

            If Me.txtNotes.Text.Trim.Length > 0 Then
               Dim sMemo As String = Me.txtNotes.Text
               Dim iNL As Integer = oUtil.MLCount(sMemo, 60)
               Dim k As Integer
               ps.Append(vbCrLf & vbCrLf & "Notes:" & vbCrLf)

               For k = 1 To iNL
                  ps.Append(oUtil.MemoLine(sMemo, 60, k) & vbCrLf)
               Next
            End If
         Else
            ps.Append(vbCrLf & Space(DTSP) & "Bal Due".PadRight(11) & FormatCurrency(0).PadLeft(10) & vbCrLf)
            ps.Append(vbCrLf & vbCrLf & "Notes:" & " INVOICE IS VOID" & vbCrLf)
         End If
         If modMain.fMainForm.mnuPreviewBeforePrint.Checked Then
            'oPD.PrintPreview(ps, InvoiceId)
         Else
            'oPD.StartPrint(ps, InvoiceId)
         End If
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub



   Private Sub chkManualDepositOverride_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
      LoadTheGrid()
   End Sub

   ''Public Sub cmdDelete_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
   ''   Dim nTotal As Integer
   ''   Dim nTotalSelRows As Short
   ''   Dim i As Short
   ''   Dim bkmrk As Object ' Bookmarks are always defined as variants
   ''   Dim oItem As CItems
   ''   Dim oFrm As frmAddToList
   ''   Dim j As Short

   ''   If MsgBox("Are you sure you want to delete the selected items?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
   ''      Exit Sub
   ''   End If

   ''   ' In the following, get the bookmark of the selected rows
   ''   ' loop thru the selected rows, find the item in the collection
   ''   ' and delete it.
   ''   For i = 0 To nTotalSelRows

   ''      For j = 0 To modMain.colItems.Count() - 1
   ''         With modMain.colItems(j)
   ''            'If .ItemID = Me.dbgShoppingList.Columns(0).CellValue(bkmrk) Then
   ''            '   colItems.Remove(j)
   ''            '   Exit For
   ''            'End If
   ''         End With
   ''      Next j
   ''   Next i


   ''   LoadTheGrid()
   ''End Sub
   Private Sub PrintHandler()
      '  Print the checkout bill
      If Not Me.VerifySettingsForCkOut Then Exit Sub
      m_bPrintedInvoice = True
      PrintInvoice(Val(Me.txtInvoiceID.Text))
      Me.btnCheckOut.Enabled = True
   End Sub

   Public Sub cmdPrintContract_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPrintContract.Click
      PrintHandler()
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

   Private Sub chkChargeTax_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
      If Not mbFormLoading Then LoadTheGrid()
   End Sub


   Private Sub chkDepositRequired_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
      If Not mbFormLoading Then
         If eventSender.Checked Then
            'If Me.chkDepositRequired.Value = vbChecked Then
            '   Me.cboDeliveryDistance.ListIndex = 0
            'End If
            LoadTheGrid()
         End If
      End If
   End Sub

   Private Sub frmCheckIn_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
      If mbFormLoading Then
         mbFormLoading = False
         LoadTheGrid()
      End If
   End Sub



   Private Sub cmdClose_Click()
      Me.Close()
      System.Windows.Forms.Application.DoEvents()
   End Sub
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
   Public Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
      CancelHandler()
   End Sub

   Private Sub frmCheckIn_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
      CenterForm(Me)
      mbFormLoading = True
      InitialLoadData()
      'AddHandler MyCombo.TextChanged, AddressOf Ctrls_TextChanged
      'SetUpComboColumn()
      Me.optLeftBlankCheck.Enabled = False
      Me.optLeftCardNumber.Enabled = False
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
         With dt.Rows(0)
            Me.txtNotes.Text = IIf(IsDBNull(.Item("Notes")), "", .Item("Notes"))
            Me.txtInvoiceID.Text = .Item("invoiceid")
            Me.txtCustomerID.Text = .Item("customerid")
            Me.txtPONbr.Text = .Item("ponumber")
            Me.txtCheckNumber.Text = .Item("ckcardnumber")
            Me.txtExpMonth.Text = MNS(.Item("exp_month"))
            Me.txtExpYr.Text = MNS(.Item("exp_yr"))
            Me.txtCardId.Text = MNS(.Item("card_id"))
            Me.txtContactName.Text = .Item("contactname")
            Select Case .Item("paidoption")
               Case "CK" : Me.optPaidByCheck.Checked = True
               Case "CA" : Me.optCash.Checked = True
               Case "CC" : Me.optPaidByCreditCard.Checked = True
               Case "LC" : Me.optLeftCardNumber.Checked = True
               Case "BC" : Me.optLeftBlankCheck.Checked = True
               Case "BT" : Me.optBillTo.Checked = True
            End Select
            Me.txtShipToCustomer.Text = .Item("shiptocustomer")
            Me.txtShipAddress1.Text = .Item("shiptoaddress")
            Me.txtShipCity.Text = .Item("shiptocity")
            Me.txtShipState.Text = .Item("shiptostate")
            Me.txtShipZip.Text = .Item("shiptozip")
         End With
      End If

      ' 2) load the customer data
      SQL = "select * from customers "
      SQL &= "where customerid = " & dt.Rows(0).Item("customerid") & " "
      dt.Reset()
      If oDA.SendQuery(SQL, dt, ConnectString) > 0 Then
         With dt.Rows(0)
            Me.txtCompanyName.Text = IIf(IsDBNull(.Item("companyname")), "", .Item("companyname"))
            Me.txtBillingAddress1.Text = IIf(IsDBNull(.Item("billingaddress1")), "", .Item("billingaddress1"))
            Me.txtCity.Text = IIf(IsDBNull(.Item("city")), "", .Item("city"))
            Me.txtState.Text = IIf(IsDBNull(.Item("state")), "", .Item("state"))
            Me.txtPostalCode.Text = IIf(IsDBNull(.Item("postalcode")), "", .Item("postalcode"))
            Me.txtTaxID.Text = IIf(IsDBNull(.Item("tax_id")), "", .Item("tax_id"))
         End With
      End If

      ' 3) load the total boxes
      SQL = "select * from invoice_details "
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
                     Me.txtSalesTax.Text = FormatCurrency(.Item("salestax"))
                     'If Not IsDBNull(.Item("equip_id")) AndAlso _
                     '   CType(.Item("equip_id"), String).Trim.Length > 0 Then
                     '   Me.txtSalesTax.Visible = False
                     '   Me.lblTax.Visible = False
                     'End If
                  Case 45 ' deleivery
                     Me.txtDelivery.Text = FormatCurrency(.Item("delivery"))
                  Case 55 ' amt paid
                     Me.txtAmtPaid.Text = FormatCurrency(.Item("amtpaid"))
               End Select
            End With
         Next i
      End If

      ' 2) load the invoice data into the grid (record type 1)
      SQL = "select Equip_Id,Equip_Name,Quantity,Rental_Period,PricePerUnit,Rented_Date "
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
            "MM/dd/yyyy HH:mm tt", "100", "T", "L"}
         oCG.SetTablesStyle(dtList, Me.dbgShoppingList, formats)
         'oCG.BindDataTableToGrid(dtList, dbgShoppingList)
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
      For i = 0 To dtList.Rows.Count - 1
         With dtList.Rows(i)
            total += .Item("priceperunit") * Val(.Item("quantity"))
         End With
      Next
      Me.txtItemTotal.Text = FormatCurrency(total)
      total += UnFormat(Me.txtDelivery.Text)
      total += UnFormat(Me.txtDeposit.Text)
      If Me.txtSalesTax.Visible Then
         total += UnFormat(Me.txtSalesTax.Text)
      End If
      Me.txtTotal.Text = FormatCurrency(total)
      total -= UnFormat(Me.txtAmtPaid.Text)
      Me.txtBalDue.Text = FormatCurrency(total)
   End Sub
   Private Sub LoadTheGrid()
      ' for chek in we need to 
      ' 1) load the invoice header data
      ' 2) load the invoice data into the grid (record type 1)
      ' 3) load the total boxes
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
      Dim decLabor As Decimal

      Try
         '' get the highest customer number 
         'SQL = "select max(customerid) from customers"
         'oDA.SendQuery(SQL, dt, ConnectString)

         'If dt.Rows.Count > 0 Then
         '   Me.txtCustomerID.Text = dt.Rows(0).Item(0)
         'End If

         lcurItemTotal = 0



         If Not bGridloaded Then
            dtFC = New DataTable()
            SQL = "select * from tempitems where user_id = '" & UserName & "' order by ItemId"
            oDA.SendQuery(SQL, dtFC, ConnectString)
         End If

         ' loop thru dt to accumulate and calc totals
         ' ItemTotal, Delivery, Sales Tax, Deposit, Total
         If dtFC.Rows.Count > 0 Then
            For i = 0 To dtFC.Rows.Count - 1
               With dtFC.Rows(i)
                  ' now compute the running total
                  lcurDeposit += .Item("ItemDeposit")
                  lcurPrice += .Item("ItemExtendedPrice")
                  'lcurItemTotal += .Item("ItemTotal")
               End With
            Next i

            If Not bGridloaded Then
               Me.dbgShoppingList.DataSource = dtFC
               bGridloaded = True
            End If

            ' item total
            Me.txtItemTotal.Text = FormatCurrency(lcurPrice)

            ' total line:
            ' Total      Price    Delivery   Deposit   Total
            Me.txtTotal.Text = FormatCurrency(lcurPrice + _
                               lcurTax + _
                               UnFormat(Me.txtDeposit.Text) + _
                               lcurDelivery)
            Me.txtAmtPaid.Text = Me.txtTotal.Text
            Me.txtBalDue.Text = FormatCurrency(UnFormat(Me.txtTotal.Text) - _
               UnFormat(Me.txtAmtPaid.Text))
         End If
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
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
            Me.txtCustomerID.Text = IIf(IsDBNull(.Item("customerid")), "", .Item("customerid"))
            Me.txtCompanyName.Text = IIf(IsDBNull(.Item("companyname")), "", .Item("companyname"))
            Me.txtBillingAddress1.Text = IIf(IsDBNull(.Item("billingaddress1")), "", .Item("billingaddress1"))
            Me.txtCity.Text = IIf(IsDBNull(.Item("city")), "", .Item("city"))
            Me.txtState.Text = IIf(IsDBNull(.Item("state")), "", .Item("state"))
            Me.txtPostalCode.Text = IIf(IsDBNull(.Item("postalcode")), "", .Item("postalcode"))
            Me.txtTaxID.Text = IIf(IsDBNull(.Item("tax_id")), "", .Item("tax_id"))
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


   Public Property CurrentInvoice() As Integer
      Get
         Return m_CurrentInvoice
      End Get
      Set(ByVal Value As Integer)
         m_CurrentInvoice = Value
      End Set
   End Property

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
         MovePaidToDue()
         ManualRecalc()
      End If
   End Sub

   Private Sub optLeftCardNumber_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optLeftCardNumber.CheckedChanged
      MovePaidToDue()
      ManualRecalc()
   End Sub
   Private Sub optCash_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles optCash.CheckedChanged
      ManualRecalc()
   End Sub

   Private Sub optPaidByCheck_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optPaidByCheck.CheckedChanged
      If Me.optPaidByCheck.Checked Then
         MoveDueToPaid()
         ManualRecalc()
      End If
   End Sub
   Private Sub optPaidByCreditCard_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles optPaidByCreditCard.CheckedChanged
      ManualRecalc()
   End Sub

   Private Sub MovePaidToDue()
      'Me.txtBalDue.Text = Me.txtAmtPaid.Text
      'Me.txtAmtPaid.Text = FormatCurrency(0)
   End Sub
   Private Sub MoveDueToPaid()
      'Me.txtAmtPaid.Text = Me.txtBalDue.Text
      'Me.txtBalDue.Text = FormatCurrency(0)
   End Sub

    Private Sub ManualRecalc()
        Dim laborCost As Decimal

        Try
            Dim amt As Decimal = 0
            Dim i As Integer

            With Me
                ' total up the detail rows
                For i = 0 To dtList.Rows.Count - 1
                    With dtList.Rows(i)
                        amt += .Item("priceperunit") * Val(.Item("quantity"))
                        If .Item("equip_id") = "Labor" Then
                            ' total up the labor costs so we can subtract
                            ' it out before computing tax
                            laborCost += .Item("priceperunit") * Val(.Item("quantity"))
                        End If
                    End With
                Next
                .txtItemTotal.Text = FormatCurrency(amt)

                If String.IsNullOrEmpty(.txtTaxID.Text) Then
                    .txtSalesTax.Text = FormatCurrency((UnFormat(.txtItemTotal.Text) + _
                                UnFormat(.txtDelivery.Text) - laborCost) * TaxRate)
                End If

                amt = FormatCurrency(UnFormat(.txtItemTotal.Text) + _
                      UnFormat(.txtDelivery.Text) + _
                      UnFormat(Me.txtDeposit.Text) + _
                      IIf(String.IsNullOrEmpty(Me.txtTaxID.Text), UnFormat(.txtSalesTax.Text), 0))

                'IIf(Me.txtTaxID.Visible, UnFormat(.txtSalesTax.Text), 0))
                Me.txtTotal.Text = FormatCurrency(amt)

                If .optBillTo.Checked Then
                    Me.txtBalDue.Text = FormatCurrency(UnFormat(Me.txtTotal.Text) - _
                                   UnFormat(Me.txtAmtPaid.Text))
                    Me.txtAmtPaidAtCkIn.Text = FormatCurrency(0)
                ElseIf Me.optPaidByCheck.Checked Or _
                   Me.optPaidByCreditCard.Checked Or _
                   Me.optCash.Checked Then
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
                End If
            End With
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
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
     Handles mnuDaily.Click, _
      mnuHalfDay.Click, _
      mnuMonth.Click, _
      mnuWeek.Click, _
      mnuWeekEnd.Click
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
   Private Sub CheckInHandler()
      CheckIn()
      Me.Close()
      System.Windows.Forms.Application.DoEvents()
      modMain.fMainForm.LoadEquipGridFromType(0)

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
      voidInvoice = True
      PrintInvoice(Me.txtInvoiceID.Text)
      CheckIn()
   End Sub
#End If
End Class
