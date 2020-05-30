' This form provides the ability to add fuel, labor,
' and tools and supplies to an invoice at checkin time
Imports System.Windows.Forms.Application

Public Class frmMiscCheckInItems
   Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

   Public Sub New(ByRef dt As DataTable, Optional ByVal CkOut As Boolean = False)
      MyBase.New()

      'This call is required by the Windows Form Designer.
      InitializeComponent()

      'Add any initialization after the InitializeComponent() call
      dtList = dt
      CheckingOut = CkOut
      colSaleItems = New Collection()
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
   Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
   Friend WithEvents Labor As System.Windows.Forms.TabPage
   Friend WithEvents Fuel As System.Windows.Forms.TabPage
   Friend WithEvents btnCancel As System.Windows.Forms.Button
   Friend WithEvents btnSave As System.Windows.Forms.Button
   Friend WithEvents dgLaborTypes As System.Windows.Forms.DataGrid
   Friend WithEvents lblFuel As System.Windows.Forms.Label
   Friend WithEvents cboNumberGallons As System.Windows.Forms.ComboBox
   Friend WithEvents lblPerGallon As System.Windows.Forms.Label
   Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
   Friend WithEvents txtGasPrice As System.Windows.Forms.TextBox
   Friend WithEvents Label1 As System.Windows.Forms.Label
   Friend WithEvents cboGas As System.Windows.Forms.ComboBox
   Friend WithEvents Label2 As System.Windows.Forms.Label
   Friend WithEvents txtPropanePrice As System.Windows.Forms.TextBox
   Friend WithEvents Label3 As System.Windows.Forms.Label
   Friend WithEvents cboPropane As System.Windows.Forms.ComboBox
   Friend WithEvents Label4 As System.Windows.Forms.Label
   Friend WithEvents lblMC As System.Windows.Forms.Label
   Friend WithEvents txtMiscChg As System.Windows.Forms.TextBox
   Friend WithEvents lblMiscPrice As System.Windows.Forms.Label
   Friend WithEvents txtMiscPrice As System.Windows.Forms.TextBox
   Friend WithEvents Supplies As System.Windows.Forms.TabPage
   Friend WithEvents HelpProvider1 As System.Windows.Forms.HelpProvider
   Friend WithEvents Label5 As System.Windows.Forms.Label
   Friend WithEvents Label6 As System.Windows.Forms.Label
   Friend WithEvents Label7 As System.Windows.Forms.Label
   Friend WithEvents Label8 As System.Windows.Forms.Label
   Friend WithEvents btnAddItemsAndClose As System.Windows.Forms.Button
   Friend WithEvents btnAddSelectedItems As System.Windows.Forms.Button
   Friend WithEvents txtExtendedPrice As System.Windows.Forms.TextBox
   Friend WithEvents txtPricePerUnit As System.Windows.Forms.TextBox
   Friend WithEvents cboQty As System.Windows.Forms.ComboBox
   Friend WithEvents dbgToolnSupplies As System.Windows.Forms.DataGrid
   Friend WithEvents Button1 As System.Windows.Forms.Button
   Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
   Friend WithEvents ReRent As System.Windows.Forms.TabPage
   Friend WithEvents cbNbrPeriods As System.Windows.Forms.ComboBox
   Friend WithEvents Label9 As System.Windows.Forms.Label
   Friend WithEvents Label10 As System.Windows.Forms.Label
   Friend WithEvents cbPeriod As System.Windows.Forms.ComboBox
   Friend WithEvents btnAddReRent As System.Windows.Forms.Button
   Friend WithEvents textReRentPrice As System.Windows.Forms.TextBox
   Friend WithEvents Label11 As System.Windows.Forms.Label
   Friend WithEvents textReRentEquip As System.Windows.Forms.TextBox
   Friend WithEvents Label12 As System.Windows.Forms.Label
   Friend WithEvents textPO As System.Windows.Forms.TextBox
   Friend WithEvents Label13 As System.Windows.Forms.Label
   Friend WithEvents Label14 As System.Windows.Forms.Label
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
      Me.components = New System.ComponentModel.Container()
      Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMiscCheckInItems))
      Me.TabControl1 = New System.Windows.Forms.TabControl()
      Me.Supplies = New System.Windows.Forms.TabPage()
      Me.Label5 = New System.Windows.Forms.Label()
      Me.Label6 = New System.Windows.Forms.Label()
      Me.Label7 = New System.Windows.Forms.Label()
      Me.Label8 = New System.Windows.Forms.Label()
      Me.btnAddItemsAndClose = New System.Windows.Forms.Button()
      Me.btnAddSelectedItems = New System.Windows.Forms.Button()
      Me.txtExtendedPrice = New System.Windows.Forms.TextBox()
      Me.txtPricePerUnit = New System.Windows.Forms.TextBox()
      Me.cboQty = New System.Windows.Forms.ComboBox()
      Me.dbgToolnSupplies = New System.Windows.Forms.DataGrid()
      Me.Labor = New System.Windows.Forms.TabPage()
      Me.Button1 = New System.Windows.Forms.Button()
      Me.txtMiscPrice = New System.Windows.Forms.TextBox()
      Me.lblMiscPrice = New System.Windows.Forms.Label()
      Me.txtMiscChg = New System.Windows.Forms.TextBox()
      Me.lblMC = New System.Windows.Forms.Label()
      Me.dgLaborTypes = New System.Windows.Forms.DataGrid()
      Me.Fuel = New System.Windows.Forms.TabPage()
      Me.txtPropanePrice = New System.Windows.Forms.TextBox()
      Me.Label3 = New System.Windows.Forms.Label()
      Me.cboPropane = New System.Windows.Forms.ComboBox()
      Me.Label4 = New System.Windows.Forms.Label()
      Me.txtGasPrice = New System.Windows.Forms.TextBox()
      Me.Label1 = New System.Windows.Forms.Label()
      Me.cboGas = New System.Windows.Forms.ComboBox()
      Me.Label2 = New System.Windows.Forms.Label()
      Me.TextBox1 = New System.Windows.Forms.TextBox()
      Me.lblPerGallon = New System.Windows.Forms.Label()
      Me.cboNumberGallons = New System.Windows.Forms.ComboBox()
      Me.lblFuel = New System.Windows.Forms.Label()
      Me.ReRent = New System.Windows.Forms.TabPage()
      Me.textPO = New System.Windows.Forms.TextBox()
      Me.Label13 = New System.Windows.Forms.Label()
      Me.cbNbrPeriods = New System.Windows.Forms.ComboBox()
      Me.Label9 = New System.Windows.Forms.Label()
      Me.Label10 = New System.Windows.Forms.Label()
      Me.cbPeriod = New System.Windows.Forms.ComboBox()
      Me.btnAddReRent = New System.Windows.Forms.Button()
      Me.textReRentPrice = New System.Windows.Forms.TextBox()
      Me.Label11 = New System.Windows.Forms.Label()
      Me.textReRentEquip = New System.Windows.Forms.TextBox()
      Me.Label12 = New System.Windows.Forms.Label()
      Me.btnCancel = New System.Windows.Forms.Button()
      Me.btnSave = New System.Windows.Forms.Button()
      Me.HelpProvider1 = New System.Windows.Forms.HelpProvider()
      Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
      Me.Label14 = New System.Windows.Forms.Label()
      Me.TabControl1.SuspendLayout()
      Me.Supplies.SuspendLayout()
      CType(Me.dbgToolnSupplies, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.Labor.SuspendLayout()
      CType(Me.dgLaborTypes, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.Fuel.SuspendLayout()
      Me.ReRent.SuspendLayout()
      Me.SuspendLayout()
      '
      'TabControl1
      '
      Me.TabControl1.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                  Or System.Windows.Forms.AnchorStyles.Left) _
                  Or System.Windows.Forms.AnchorStyles.Right)
      Me.TabControl1.Controls.AddRange(New System.Windows.Forms.Control() {Me.Supplies, Me.Labor, Me.Fuel, Me.ReRent})
      Me.TabControl1.Name = "TabControl1"
      Me.TabControl1.SelectedIndex = 0
      Me.TabControl1.Size = New System.Drawing.Size(518, 318)
      Me.TabControl1.TabIndex = 0
      '
      'Supplies
      '
      Me.Supplies.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label5, Me.Label6, Me.Label7, Me.Label8, Me.btnAddItemsAndClose, Me.btnAddSelectedItems, Me.txtExtendedPrice, Me.txtPricePerUnit, Me.cboQty, Me.dbgToolnSupplies})
      Me.Supplies.Location = New System.Drawing.Point(4, 22)
      Me.Supplies.Name = "Supplies"
      Me.Supplies.Size = New System.Drawing.Size(510, 292)
      Me.Supplies.TabIndex = 2
      Me.Supplies.Text = "Tools & Supplies"
      '
      'Label5
      '
      Me.Label5.AutoSize = True
      Me.Label5.Location = New System.Drawing.Point(8, 8)
      Me.Label5.Name = "Label5"
      Me.Label5.Size = New System.Drawing.Size(135, 13)
      Me.Label5.TabIndex = 18
      Me.Label5.Text = "Select Tools and Supplies"
      '
      'Label6
      '
      Me.Label6.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
      Me.Label6.AutoSize = True
      Me.Label6.Location = New System.Drawing.Point(192, 216)
      Me.Label6.Name = "Label6"
      Me.Label6.Size = New System.Drawing.Size(81, 13)
      Me.Label6.TabIndex = 16
      Me.Label6.Text = "Extended Price"
      '
      'Label7
      '
      Me.Label7.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
      Me.Label7.AutoSize = True
      Me.Label7.Location = New System.Drawing.Point(96, 216)
      Me.Label7.Name = "Label7"
      Me.Label7.Size = New System.Drawing.Size(53, 13)
      Me.Label7.TabIndex = 14
      Me.Label7.Text = "Unit Price"
      '
      'Label8
      '
      Me.Label8.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
      Me.Label8.AutoSize = True
      Me.Label8.Location = New System.Drawing.Point(24, 216)
      Me.Label8.Name = "Label8"
      Me.Label8.Size = New System.Drawing.Size(46, 13)
      Me.Label8.TabIndex = 13
      Me.Label8.Text = "Quantity"
      '
      'btnAddItemsAndClose
      '
      Me.btnAddItemsAndClose.Location = New System.Drawing.Point(192, 232)
      Me.btnAddItemsAndClose.Name = "btnAddItemsAndClose"
      Me.btnAddItemsAndClose.Size = New System.Drawing.Size(112, 27)
      Me.btnAddItemsAndClose.TabIndex = 20
      Me.btnAddItemsAndClose.Text = "Add Items to Cart"
      Me.btnAddItemsAndClose.Visible = False
      '
      'btnAddSelectedItems
      '
      Me.btnAddSelectedItems.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
      Me.btnAddSelectedItems.BackColor = System.Drawing.SystemColors.Control
      Me.btnAddSelectedItems.Location = New System.Drawing.Point(386, 224)
      Me.btnAddSelectedItems.Name = "btnAddSelectedItems"
      Me.btnAddSelectedItems.Size = New System.Drawing.Size(104, 48)
      Me.btnAddSelectedItems.TabIndex = 19
      Me.btnAddSelectedItems.Text = "Add Selected Item To List"
      '
      'txtExtendedPrice
      '
      Me.txtExtendedPrice.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
      Me.txtExtendedPrice.Location = New System.Drawing.Point(192, 232)
      Me.txtExtendedPrice.Name = "txtExtendedPrice"
      Me.txtExtendedPrice.ReadOnly = True
      Me.txtExtendedPrice.Size = New System.Drawing.Size(80, 20)
      Me.txtExtendedPrice.TabIndex = 17
      Me.txtExtendedPrice.Text = ""
      Me.txtExtendedPrice.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      '
      'txtPricePerUnit
      '
      Me.txtPricePerUnit.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
      Me.txtPricePerUnit.Location = New System.Drawing.Point(96, 232)
      Me.txtPricePerUnit.Name = "txtPricePerUnit"
      Me.txtPricePerUnit.Size = New System.Drawing.Size(80, 20)
      Me.txtPricePerUnit.TabIndex = 15
      Me.txtPricePerUnit.Tag = "(No Auto Formatting)"
      Me.txtPricePerUnit.Text = ""
      Me.txtPricePerUnit.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      Me.ToolTip1.SetToolTip(Me.txtPricePerUnit, "Change this price manually to give a price break")
      '
      'cboQty
      '
      Me.cboQty.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
      Me.cboQty.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "30", "40", "50", "75", "100"})
      Me.cboQty.Location = New System.Drawing.Point(24, 232)
      Me.cboQty.Name = "cboQty"
      Me.cboQty.Size = New System.Drawing.Size(63, 21)
      Me.cboQty.TabIndex = 12
      Me.cboQty.Text = "1"
      '
      'dbgToolnSupplies
      '
      Me.dbgToolnSupplies.AllowSorting = False
      Me.dbgToolnSupplies.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                  Or System.Windows.Forms.AnchorStyles.Left) _
                  Or System.Windows.Forms.AnchorStyles.Right)
      Me.dbgToolnSupplies.CaptionVisible = False
      Me.dbgToolnSupplies.DataMember = ""
      Me.dbgToolnSupplies.HeaderForeColor = System.Drawing.SystemColors.ControlText
      Me.dbgToolnSupplies.Location = New System.Drawing.Point(8, 24)
      Me.dbgToolnSupplies.Name = "dbgToolnSupplies"
      Me.dbgToolnSupplies.ReadOnly = True
      Me.dbgToolnSupplies.Size = New System.Drawing.Size(493, 179)
      Me.dbgToolnSupplies.TabIndex = 11
      '
      'Labor
      '
      Me.Labor.Controls.AddRange(New System.Windows.Forms.Control() {Me.Button1, Me.txtMiscPrice, Me.lblMiscPrice, Me.txtMiscChg, Me.lblMC, Me.dgLaborTypes})
      Me.Labor.Location = New System.Drawing.Point(4, 22)
      Me.Labor.Name = "Labor"
      Me.Labor.Size = New System.Drawing.Size(510, 292)
      Me.Labor.TabIndex = 0
      Me.Labor.Text = "Labor"
      '
      'Button1
      '
      Me.Button1.Location = New System.Drawing.Point(304, 168)
      Me.Button1.Name = "Button1"
      Me.Button1.Size = New System.Drawing.Size(112, 27)
      Me.Button1.TabIndex = 21
      Me.Button1.Text = "Add Items to Cart"
      '
      'txtMiscPrice
      '
      Me.txtMiscPrice.Location = New System.Drawing.Point(84, 167)
      Me.txtMiscPrice.Name = "txtMiscPrice"
      Me.txtMiscPrice.Size = New System.Drawing.Size(64, 20)
      Me.txtMiscPrice.TabIndex = 4
      Me.txtMiscPrice.Tag = "$#,##0.00;($#,##0.00)"
      Me.txtMiscPrice.Text = ""
      Me.txtMiscPrice.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      '
      'lblMiscPrice
      '
      Me.lblMiscPrice.AutoSize = True
      Me.lblMiscPrice.Location = New System.Drawing.Point(40, 168)
      Me.lblMiscPrice.Name = "lblMiscPrice"
      Me.lblMiscPrice.Size = New System.Drawing.Size(30, 13)
      Me.lblMiscPrice.TabIndex = 3
      Me.lblMiscPrice.Text = "Price"
      '
      'txtMiscChg
      '
      Me.txtMiscChg.Location = New System.Drawing.Point(85, 142)
      Me.txtMiscChg.MaxLength = 29
      Me.txtMiscChg.Name = "txtMiscChg"
      Me.txtMiscChg.Size = New System.Drawing.Size(312, 20)
      Me.txtMiscChg.TabIndex = 2
      Me.txtMiscChg.Text = ""
      '
      'lblMC
      '
      Me.lblMC.AutoSize = True
      Me.lblMC.Location = New System.Drawing.Point(8, 144)
      Me.lblMC.Name = "lblMC"
      Me.lblMC.Size = New System.Drawing.Size(68, 13)
      Me.lblMC.TabIndex = 1
      Me.lblMC.Text = "Misc Charge"
      '
      'dgLaborTypes
      '
      Me.dgLaborTypes.CaptionText = "Select Labor Charges"
      Me.dgLaborTypes.DataMember = ""
      Me.dgLaborTypes.HeaderForeColor = System.Drawing.SystemColors.ControlText
      Me.dgLaborTypes.Location = New System.Drawing.Point(8, 8)
      Me.dgLaborTypes.Name = "dgLaborTypes"
      Me.dgLaborTypes.Size = New System.Drawing.Size(416, 128)
      Me.dgLaborTypes.TabIndex = 0
      '
      'Fuel
      '
      Me.Fuel.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtPropanePrice, Me.Label3, Me.cboPropane, Me.Label4, Me.txtGasPrice, Me.Label1, Me.cboGas, Me.Label2, Me.TextBox1, Me.lblPerGallon, Me.cboNumberGallons, Me.lblFuel})
      Me.Fuel.Location = New System.Drawing.Point(4, 22)
      Me.Fuel.Name = "Fuel"
      Me.Fuel.Size = New System.Drawing.Size(510, 292)
      Me.Fuel.TabIndex = 1
      Me.Fuel.Text = "Fuel"
      '
      'txtPropanePrice
      '
      Me.txtPropanePrice.Location = New System.Drawing.Point(314, 69)
      Me.txtPropanePrice.Name = "txtPropanePrice"
      Me.txtPropanePrice.Size = New System.Drawing.Size(56, 20)
      Me.txtPropanePrice.TabIndex = 11
      Me.txtPropanePrice.Tag = "$#,##0.00;($#,##0.00)"
      Me.txtPropanePrice.Text = "0"
      Me.txtPropanePrice.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      '
      'Label3
      '
      Me.Label3.AutoSize = True
      Me.Label3.Location = New System.Drawing.Point(242, 69)
      Me.Label3.Name = "Label3"
      Me.Label3.Size = New System.Drawing.Size(58, 13)
      Me.Label3.TabIndex = 10
      Me.Label3.Text = "Per Gallon"
      '
      'cboPropane
      '
      Me.cboPropane.Items.AddRange(New Object() {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "25", "30", "35", "40", "45", "50", "55", "60", "65", "70", "75", "80", "85", "90", "95", "100", "200", "300"})
      Me.cboPropane.Location = New System.Drawing.Point(179, 69)
      Me.cboPropane.Name = "cboPropane"
      Me.cboPropane.Size = New System.Drawing.Size(56, 21)
      Me.cboPropane.TabIndex = 9
      Me.cboPropane.Text = "0"
      '
      'Label4
      '
      Me.Label4.AutoSize = True
      Me.Label4.Location = New System.Drawing.Point(9, 69)
      Me.Label4.Name = "Label4"
      Me.Label4.Size = New System.Drawing.Size(158, 13)
      Me.Label4.TabIndex = 8
      Me.Label4.Text = "Select # Propane Gallons Sold"
      '
      'txtGasPrice
      '
      Me.txtGasPrice.Location = New System.Drawing.Point(314, 40)
      Me.txtGasPrice.Name = "txtGasPrice"
      Me.txtGasPrice.Size = New System.Drawing.Size(56, 20)
      Me.txtGasPrice.TabIndex = 7
      Me.txtGasPrice.Tag = "$#,##0.00;($#,##0.00)"
      Me.txtGasPrice.Text = "0"
      Me.txtGasPrice.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      '
      'Label1
      '
      Me.Label1.AutoSize = True
      Me.Label1.Location = New System.Drawing.Point(242, 40)
      Me.Label1.Name = "Label1"
      Me.Label1.Size = New System.Drawing.Size(58, 13)
      Me.Label1.TabIndex = 6
      Me.Label1.Text = "Per Gallon"
      '
      'cboGas
      '
      Me.cboGas.Items.AddRange(New Object() {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "25", "30", "35", "40", "45", "50", "55", "60", "65", "70", "75", "80", "85", "90", "95", "100", "200", "300"})
      Me.cboGas.Location = New System.Drawing.Point(179, 40)
      Me.cboGas.Name = "cboGas"
      Me.cboGas.Size = New System.Drawing.Size(56, 21)
      Me.cboGas.TabIndex = 5
      Me.cboGas.Text = "0"
      '
      'Label2
      '
      Me.Label2.AutoSize = True
      Me.Label2.Location = New System.Drawing.Point(9, 40)
      Me.Label2.Name = "Label2"
      Me.Label2.Size = New System.Drawing.Size(160, 13)
      Me.Label2.TabIndex = 4
      Me.Label2.Text = "Select # Gasoline Gallons Sold"
      '
      'TextBox1
      '
      Me.TextBox1.Location = New System.Drawing.Point(314, 10)
      Me.TextBox1.Name = "TextBox1"
      Me.TextBox1.Size = New System.Drawing.Size(56, 20)
      Me.TextBox1.TabIndex = 3
      Me.TextBox1.Tag = "$#,##0.00;($#,##0.00)"
      Me.TextBox1.Text = "0"
      Me.TextBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      '
      'lblPerGallon
      '
      Me.lblPerGallon.AutoSize = True
      Me.lblPerGallon.Location = New System.Drawing.Point(242, 13)
      Me.lblPerGallon.Name = "lblPerGallon"
      Me.lblPerGallon.Size = New System.Drawing.Size(58, 13)
      Me.lblPerGallon.TabIndex = 2
      Me.lblPerGallon.Text = "Per Gallon"
      '
      'cboNumberGallons
      '
      Me.cboNumberGallons.Items.AddRange(New Object() {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "25", "30", "35", "40", "45", "50", "55", "60", "65", "70", "75", "80", "85", "90", "95", "100", "200", "300"})
      Me.cboNumberGallons.Location = New System.Drawing.Point(178, 11)
      Me.cboNumberGallons.Name = "cboNumberGallons"
      Me.cboNumberGallons.Size = New System.Drawing.Size(56, 21)
      Me.cboNumberGallons.TabIndex = 1
      Me.cboNumberGallons.Text = "0"
      '
      'lblFuel
      '
      Me.lblFuel.AutoSize = True
      Me.lblFuel.Location = New System.Drawing.Point(8, 13)
      Me.lblFuel.Name = "lblFuel"
      Me.lblFuel.Size = New System.Drawing.Size(150, 13)
      Me.lblFuel.TabIndex = 0
      Me.lblFuel.Text = "Select # Diesel Gallons Sold "
      '
      'ReRent
      '
      Me.ReRent.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label14, Me.textPO, Me.Label13, Me.cbNbrPeriods, Me.Label9, Me.Label10, Me.cbPeriod, Me.btnAddReRent, Me.textReRentPrice, Me.Label11, Me.textReRentEquip, Me.Label12})
      Me.ReRent.Location = New System.Drawing.Point(4, 22)
      Me.ReRent.Name = "ReRent"
      Me.ReRent.Size = New System.Drawing.Size(510, 292)
      Me.ReRent.TabIndex = 3
      Me.ReRent.Text = "Re-Rent"
      '
      'textPO
      '
      Me.textPO.Location = New System.Drawing.Point(104, 206)
      Me.textPO.Name = "textPO"
      Me.textPO.Size = New System.Drawing.Size(111, 20)
      Me.textPO.TabIndex = 24
      Me.textPO.Tag = "(No Auto Formatting)"
      Me.textPO.Text = ""
      Me.textPO.Visible = False
      '
      'Label13
      '
      Me.Label13.AutoSize = True
      Me.Label13.Location = New System.Drawing.Point(8, 206)
      Me.Label13.Name = "Label13"
      Me.Label13.Size = New System.Drawing.Size(84, 13)
      Me.Label13.TabIndex = 23
      Me.Label13.Text = "Purchase Order"
      Me.Label13.Visible = False
      '
      'cbNbrPeriods
      '
      Me.cbNbrPeriods.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6", "7", "8", "9", "10"})
      Me.cbNbrPeriods.Location = New System.Drawing.Point(104, 182)
      Me.cbNbrPeriods.Name = "cbNbrPeriods"
      Me.cbNbrPeriods.Size = New System.Drawing.Size(88, 21)
      Me.cbNbrPeriods.TabIndex = 22
      Me.cbNbrPeriods.Visible = False
      '
      'Label9
      '
      Me.Label9.AutoSize = True
      Me.Label9.Location = New System.Drawing.Point(8, 182)
      Me.Label9.Name = "Label9"
      Me.Label9.Size = New System.Drawing.Size(86, 13)
      Me.Label9.TabIndex = 21
      Me.Label9.Text = "Number Periods"
      Me.Label9.Visible = False
      '
      'Label10
      '
      Me.Label10.AutoSize = True
      Me.Label10.Location = New System.Drawing.Point(16, 158)
      Me.Label10.Name = "Label10"
      Me.Label10.Size = New System.Drawing.Size(73, 13)
      Me.Label10.TabIndex = 20
      Me.Label10.Text = "Rental Period"
      Me.Label10.Visible = False
      '
      'cbPeriod
      '
      Me.cbPeriod.Items.AddRange(New Object() {"Half Day", "Daily", "Weekly", "Monthly", "Week End"})
      Me.cbPeriod.Location = New System.Drawing.Point(104, 158)
      Me.cbPeriod.Name = "cbPeriod"
      Me.cbPeriod.Size = New System.Drawing.Size(88, 21)
      Me.cbPeriod.TabIndex = 19
      Me.cbPeriod.Visible = False
      '
      'btnAddReRent
      '
      Me.btnAddReRent.Location = New System.Drawing.Point(248, 206)
      Me.btnAddReRent.Name = "btnAddReRent"
      Me.btnAddReRent.Size = New System.Drawing.Size(96, 24)
      Me.btnAddReRent.TabIndex = 18
      Me.btnAddReRent.Text = "Add Item to Grid"
      Me.btnAddReRent.Visible = False
      '
      'textReRentPrice
      '
      Me.textReRentPrice.Location = New System.Drawing.Point(104, 134)
      Me.textReRentPrice.Name = "textReRentPrice"
      Me.textReRentPrice.Size = New System.Drawing.Size(64, 20)
      Me.textReRentPrice.TabIndex = 17
      Me.textReRentPrice.Tag = "$#,##0.00;($#,##0.00)"
      Me.textReRentPrice.Text = ""
      Me.textReRentPrice.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      Me.textReRentPrice.Visible = False
      '
      'Label11
      '
      Me.Label11.AutoSize = True
      Me.Label11.Location = New System.Drawing.Point(56, 134)
      Me.Label11.Name = "Label11"
      Me.Label11.Size = New System.Drawing.Size(30, 13)
      Me.Label11.TabIndex = 16
      Me.Label11.Text = "Price"
      Me.Label11.Visible = False
      '
      'textReRentEquip
      '
      Me.textReRentEquip.Location = New System.Drawing.Point(104, 110)
      Me.textReRentEquip.MaxLength = 29
      Me.textReRentEquip.Name = "textReRentEquip"
      Me.textReRentEquip.Size = New System.Drawing.Size(225, 20)
      Me.textReRentEquip.TabIndex = 15
      Me.textReRentEquip.Tag = "(No Auto Formatting)"
      Me.textReRentEquip.Text = ""
      Me.textReRentEquip.Visible = False
      '
      'Label12
      '
      Me.Label12.AutoSize = True
      Me.Label12.Location = New System.Drawing.Point(16, 110)
      Me.Label12.Name = "Label12"
      Me.Label12.Size = New System.Drawing.Size(83, 13)
      Me.Label12.TabIndex = 14
      Me.Label12.Text = "ReRent Charge"
      Me.Label12.Visible = False
      '
      'btnCancel
      '
      Me.btnCancel.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
      Me.btnCancel.Location = New System.Drawing.Point(422, 325)
      Me.btnCancel.Name = "btnCancel"
      Me.btnCancel.Size = New System.Drawing.Size(88, 32)
      Me.btnCancel.TabIndex = 1
      Me.btnCancel.Text = "&Cancel"
      '
      'btnSave
      '
      Me.btnSave.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
      Me.btnSave.Location = New System.Drawing.Point(326, 326)
      Me.btnSave.Name = "btnSave"
      Me.btnSave.Size = New System.Drawing.Size(88, 32)
      Me.btnSave.TabIndex = 2
      Me.btnSave.Text = "&Save"
      '
      'Label14
      '
      Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.Label14.Location = New System.Drawing.Point(16, 16)
      Me.Label14.Name = "Label14"
      Me.Label14.Size = New System.Drawing.Size(360, 40)
      Me.Label14.TabIndex = 25
      Me.Label14.Text = "The ReRent feature is now  activated from the ReRent Button on the main forms too" & _
      "lbar."
      '
      'frmMiscCheckInItems
      '
      Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
      Me.ClientSize = New System.Drawing.Size(520, 366)
      Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnSave, Me.btnCancel, Me.TabControl1})
      Me.HelpButton = True
      Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
      Me.MaximizeBox = False
      Me.MinimizeBox = False
      Me.Name = "frmMiscCheckInItems"
      Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "Other Charges"
      Me.TabControl1.ResumeLayout(False)
      Me.Supplies.ResumeLayout(False)
      CType(Me.dbgToolnSupplies, System.ComponentModel.ISupportInitialize).EndInit()
      Me.Labor.ResumeLayout(False)
      CType(Me.dgLaborTypes, System.ComponentModel.ISupportInitialize).EndInit()
      Me.Fuel.ResumeLayout(False)
      Me.ReRent.ResumeLayout(False)
      Me.ResumeLayout(False)

   End Sub

#End Region

#Region " Module Variables "
   Private oDA As New CDataAccess()
   Private SQL As String
   Private dtLabor As New DataTable()
   Dim oCG As New CGrid()
   Private dtList As DataTable
   Private CheckingOut As Boolean ' implies different datatable
   Dim dtTS As DataTable
   Dim iHitRow As Integer
   Private colSaleItems As Collection
   Private mbFormLoading As Boolean = True
   Private mbDirty As Boolean


#End Region

#Region " Form & Control Events "
   Private Sub txtGasPrice_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtGasPrice.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
      e.Handled = CkKeyPressNumeric(Asc(Val(e.KeyChar)), txtGasPrice) = 0
   End Sub
   Private Sub txtGasPrice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtGasPrice.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub txtGasPrice_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGasPrice.Enter
      txtGasPrice.Text = UnFmt_T_B(txtGasPrice)
      txtGasPrice.SelectionStart = 0
      txtGasPrice.SelectionLength = txtGasPrice.Text.Trim.Length
   End Sub
   Private Sub txtGasPrice_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGasPrice.Leave
      txtGasPrice.Text = Fmt_T_B(txtGasPrice)
   End Sub
   Private Sub txtPropanePrice_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPropanePrice.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
      e.Handled = CkKeyPressNumeric(Asc(Val(e.KeyChar)), txtPropanePrice) = 0
   End Sub
   Private Sub txtPropanePrice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPropanePrice.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub txtPropanePrice_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPropanePrice.Enter
      txtPropanePrice.Text = UnFmt_T_B(txtPropanePrice)
      txtPropanePrice.SelectionStart = 0
      txtPropanePrice.SelectionLength = txtPropanePrice.Text.Trim.Length
   End Sub
   Private Sub txtPropanePrice_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPropanePrice.Leave
      txtPropanePrice.Text = Fmt_T_B(txtPropanePrice)
   End Sub

   Public Sub New()
      colSaleItems = New Collection()

   End Sub
   Private Sub txtMiscPrice_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMiscPrice.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
      e.Handled = CkKeyPressNumeric(Asc(Val(e.KeyChar)), txtMiscPrice) = 0
   End Sub
   Private Sub txtMiscPrice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtMiscPrice.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub txtMiscPrice_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMiscPrice.Enter
      txtMiscPrice.Text = UnFmt_T_B(txtMiscPrice)
      txtMiscPrice.SelectionStart = 0
      txtMiscPrice.SelectionLength = txtMiscPrice.Text.Trim.Length
   End Sub
   Private Sub txtMiscPrice_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMiscPrice.Leave
      txtMiscPrice.Text = Fmt_T_B(txtMiscPrice)
   End Sub

   Private Sub btnAddSelectedItems_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddSelectedItems.Click
      Dim o As New CToolsAndSupplies()
      Dim shtUnitsAvail As Short

      Try
         shtUnitsAvail = dtTS.Rows(iHitRow).Item("UnitsInStock")
         With o
            If Val(Me.cboQty.Text) > shtUnitsAvail Then
               MsgBox("There are only " & shtUnitsAvail & " available!", MsgBoxStyle.Information)
               Exit Sub
            End If
            .ItemCount = Val(Me.cboQty.Text)
            .ItemId = dtTS.Rows(iHitRow).Item("Productid")
            .ItemPrice = UnFormat(Me.txtPricePerUnit.Text) '* discount
            .ItemExtendedPrice = UnFormat(Me.txtExtendedPrice.Text)
            .Deposit = 0
            .RentOrSale = SALE
            .ItemName = dtTS.Rows(iHitRow).Item("productname")
         End With
         Me.colSaleItems.Add(o)
         Me.cboQty.Text = "1"
         mbDirty = True
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

   Private Sub txtPricePerUnit_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPricePerUnit.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
   End Sub

   Private Sub txtPricePerUnit_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPricePerUnit.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub

   Private Sub txtPricePerUnit_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPricePerUnit.Enter
      txtPricePerUnit.SelectionStart = 0
      txtPricePerUnit.SelectionLength = txtPricePerUnit.Text.Trim.Length
   End Sub

   Private Sub Recompute()
      With Me
         .txtExtendedPrice.Text = FormatCurrency(UnFormat(Me.txtPricePerUnit.Text) * Val(Me.cboQty.Text))
      End With
   End Sub

   Private Sub txtPricePerUnit_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPricePerUnit.Leave
      With Me
         txtPricePerUnit.Text = FormatCurrency(txtPricePerUnit.Text)
         Recompute()
      End With
   End Sub

   Private Sub cboQty_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboQty.SelectedIndexChanged
      If Not mbFormLoading Then
         'Me.LoadTextBoxes()
         Me.Recompute()
      End If
   End Sub

   Private Sub dbgToolnSupplies_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dbgToolnSupplies.MouseUp
      iHitRow = oCG.GetClickedRow(e, Me.dbgToolnSupplies)
      Try
         dbgToolnSupplies.Select(iHitRow)
         LoadTextBoxes()
      Catch
      End Try
   End Sub

   Private Sub btnAddReRent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddReRent.Click
      If Not CheckingOut Then
         Dim dr() As Object = {"ReRent", Me.textReRentEquip.Text, Me.cbNbrPeriods.Text, Me.cbPeriod.Text, UnFormat(Me.textReRentPrice.Text), Now.ToString, 0, 0}
         oCG.AddRowToTable(dtList, dr)
      Else
         Dim dr() As Object = {"ReRent", Me.textReRentEquip.Text, Me.cbNbrPeriods.Text, Me.cbPeriod.Text, UnFormat(Me.textReRentPrice.Text), Me.cbNbrPeriods.Text * UnFormat(Me.textReRentPrice.Text), 0, Me.textPO.Text, False, 0, UserName}
         oCG.AddRowToTable(dtList, dr)
      End If
   End Sub
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
   Private Sub textReRentPrice_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles textReRentPrice.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
      e.Handled = CkKeyPressNumeric(Asc(Val(e.KeyChar)), textReRentPrice) = 0
   End Sub
   Private Sub textReRentPrice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles textReRentPrice.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub textReRentPrice_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles textReRentPrice.Enter
      textReRentPrice.Text = UnFmt_T_B(textReRentPrice)
      textReRentPrice.SelectionStart = 0
      textReRentPrice.SelectionLength = textReRentPrice.Text.Trim.Length
   End Sub
   Private Sub textReRentPrice_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles textReRentPrice.Leave
      textReRentPrice.Text = Fmt_T_B(textReRentPrice)
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


   Private Sub dgLaborTypes_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgLaborTypes.MouseUp
      Dim bChecked As Boolean
      Dim i As Integer = oCG.SelectCkBoxRow(dtLabor, Me.dgLaborTypes, e, "Charge", bChecked)
   End Sub

   Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
      Me.Close()
      DoEvents()
   End Sub

   Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
      ' add the selected item to the passed datatable
      ' Equip_Id,Equip_Name,Quantity,Rental_Period,PricePerUnit,Rented_Date
      Dim i As Integer
      ' loop thru dt to find checked items
      For i = 0 To dtLabor.Rows.Count - 1
         With dtLabor.Rows(i)
            If .Item("Charge") = "true" Then
               ' add a row to the passed dtList
               Dim laborType As String = .Item("labor_type")
               laborType = Replace(laborType, "'", "''")
               If Not CheckingOut Then
                  Dim dr() As Object = {"Labor", laborType, 1, "Sale", .Item("Price"), Now.ToString, 0, 0}
                  oCG.AddRowToTable(dtList, dr)
               Else
                  Dim dr() As Object = {"Labor", laborType, 1, "N/A", .Item("Price"), .Item("Price"), 0, "Sale", False, 0, UserName}
                  oCG.AddRowToTable(dtList, dr)
               End If
            End If
         End With
      Next

      If Me.txtMiscChg.Text.Trim.Length > 0 Then
         If Not CheckingOut Then
            Dim dr() As Object = {"Misc", Replace(txtMiscChg.Text, "'", "''"), 1, "Sale", UnFormat(Me.txtMiscPrice.Text), Now.ToString, 0, 0}
            oCG.AddRowToTable(dtList, dr)
         Else
            Dim dr() As Object = {"Misc", Replace(txtMiscChg.Text, "'", "''"), 1, "N/A", UnFormat(Me.txtMiscPrice.Text), UnFormat(Me.txtMiscPrice.Text), 0, "Sale", False, 0, UserName}
            oCG.AddRowToTable(dtList, dr)
         End If
      End If

      If Val(Me.cboNumberGallons.Text) > 0 Then
         If Not CheckingOut Then
            Dim dr() As Object = {"Fuel", "Diesel", Val(Me.cboNumberGallons.Text), "Sale", UnFormat(Me.TextBox1.Text), Now.ToString, 0, 0}
            oCG.AddRowToTable(dtList, dr)
         Else
            Dim dr() As Object = {"Fuel", "Diesel", Val(Me.cboNumberGallons.Text), "N/A", UnFormat(Me.TextBox1.Text), Val(Me.cboNumberGallons.Text) * UnFormat(Me.TextBox1.Text), 0, "Sale", False, 0, UserName}
            oCG.AddRowToTable(dtList, dr)
         End If
      End If

      If Val(Me.cboGas.Text) > 0 Then
         If Not CheckingOut Then
            Dim dr() As Object = {"Fuel", "Gasoline", Val(Me.cboGas.Text), "Sale", UnFormat(Me.txtGasPrice.Text), Now.ToString, 0, 0}
            oCG.AddRowToTable(dtList, dr)
         Else
            Dim dr() As Object = {"Fuel", "Gasoline", Val(Me.cboGas.Text), "N/A", UnFormat(Me.txtGasPrice.Text), Val(Me.cboGas.Text) * UnFormat(Me.txtGasPrice.Text), 0, "Sale", False, 0, UserName}
            oCG.AddRowToTable(dtList, dr)
         End If
      End If

      If Val(Me.cboPropane.Text) > 0 Then
         If Not CheckingOut Then
            Dim dr() As Object = {"Fuel", "Propane", Val(Me.cboPropane.Text), "Sale", UnFormat(Me.txtPropanePrice.Text), Now.ToString, 0, 0}
            oCG.AddRowToTable(dtList, dr)
         Else
            Dim dr() As Object = {"Fuel", "Propane", Val(Me.cboPropane.Text), "N/A", UnFormat(Me.txtPropanePrice.Text), Val(Me.cboPropane.Text) * UnFormat(Me.txtPropanePrice.Text), 0, "Sale", False, 0, UserName}
            oCG.AddRowToTable(dtList, dr)
         End If
      End If

      Dim serr As String = ""
      With Me
         If .colSaleItems.Count > 0 Then
            Dim o As New CToolsAndSupplies()
            For i = 1 To Me.colSaleItems.Count
               o = colSaleItems(i)
               If Not CheckingOut Then
                  Dim dr() As Object = {o.ItemId, Replace(o.ItemName, "'", "''"), o.ItemCount.ToString, "Sale", o.ItemPrice, Now.ToString, 0, 0}
                  oCG.AddRowToTable(dtList, dr)
               Else
                  SQL = "insert into tempitems "
                  SQL &= "(itemid,itemname,itemcount,itemprice,itemextendedprice,"
                  SQL &= "rentorsale,itemdeposit,itemperiod,meter_required,hour_meter,user_id) "
                  SQL &= "values('" & o.ItemId & "', "
                  SQL &= "'" & Replace(o.ItemName, "'", "''") & "', "
                  SQL &= o.ItemCount.ToString & ", "
                  SQL &= o.ItemPrice.ToString & ", "
                  SQL &= o.ItemExtendedPrice.ToString & ", "
                  SQL &= "'" & SALE & "', "
                  SQL &= o.Deposit.ToString & ",'"
                  SQL &= "'" & SALE & "', "
                  SQL &= False & ", "
                  SQL &= "0, "
                  SQL &= "'" & UserName & "' "
                  SQL = ")"
                  If oDA.SendActionSql(SQL, ConnectString, serr) < 1 Then
                     MsgBox("Can't add item to temp table, " & serr, MsgBoxStyle.Critical)
                     Exit Sub
                  End If
               End If
            Next i
         End If
      End With

      Me.Close()
      DoEvents()
   End Sub

   Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
      e.Handled = CkKeyPressNumeric(Asc(Val(e.KeyChar)), TextBox1) = 0
   End Sub
   Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub TextBox1_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.Enter
      TextBox1.Text = UnFmt_T_B(TextBox1)
      TextBox1.SelectionStart = 0
      TextBox1.SelectionLength = TextBox1.Text.Trim.Length
   End Sub
   Private Sub TextBox1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.Leave
      TextBox1.Text = Fmt_T_B(TextBox1)
   End Sub


   Private Sub frmMiscCheckInItems_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
      Dim dt As New DataTable()
      Dim i As Short

      SQL = "select Labor_Type,Price from labor_charges "
      SQL &= "order by labor_type"
      oCG.ClearDataTableForRebinding(dtLabor)
      If oDA.SendQuery(SQL, dtLabor, ConnectString, "dt") > 0 Then
         Dim formats() As String = _
             {"", "230", "T", "L", _
             "$#,##0.00", "60", "T", "R"}
         oCG.SetTablesStyle("Charge", dtLabor, Me.dgLaborTypes, formats)
         oCG.BindDataTableToGrid(dtLabor, dgLaborTypes)
      End If
      SQL = "select * from fuel_price order by fuel_type"
      If oDA.SendQuery(SQL, dt, ConnectString) > 0 Then
         For i = 0 To dt.Rows.Count - 1
            With dt.Rows(i)
               Select Case CType(.Item("fuel_type"), String).ToUpper
                  Case "PROPANE"
                     Me.txtPropanePrice.Text = FormatCurrency(.Item("Price"))
                  Case "GASOLINE"
                     Me.txtGasPrice.Text = FormatCurrency(.Item("Price"))
                  Case "DIESEL"
                     Me.TextBox1.Text = FormatCurrency(.Item("Price"))
               End Select
            End With
         Next
      End If
      LoadTheGrid()
      iHitRow = 0
      LoadTextBoxes()
      mbFormLoading = False

   End Sub

#End Region

#Region " Private Methods "
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


   Private Sub LoadTextBoxes()
      Try
         With Me
            Me.txtPricePerUnit.Text = FormatCurrency(dtTS.Rows(iHitRow).Item("priceperunit"))
            Me.txtExtendedPrice.Text = FormatCurrency(dtTS.Rows(iHitRow).Item("priceperunit") * Val(Me.cboQty.Text))
         End With
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

   Private Sub LoadTheGrid()
      Dim i As Integer

      Try
         SQL = "select ProductId,ProductName,ProductDescription,"
         SQL &= "PricePerUnit,UnitsInStock,ReorderLevel "
         SQL &= "from Products order by productname"
         oCG.ClearDataTableForRebinding(dtTS)

         If oDA.SendQuery(SQL, dtTS, ConnectString, "dt") > 0 Then
            Dim formats() As String = _
               {"", "60", "T", "L", _
                "", "150", "T", "L", _
                "", "150", "T", "L", _
                "$#,##0.00", "60", "T", "R", _
                "", "60", "T", "R", _
                "", "80", "T", "R"}
            oCG.SetTablesStyle(dtTS, Me.dbgToolnSupplies, formats)
            oCG.BindDataTableToGrid(dtTS, Me.dbgToolnSupplies)
            oCG.DisableAddNew(Me.dbgToolnSupplies, Me)
         End If
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

#End Region

End Class
