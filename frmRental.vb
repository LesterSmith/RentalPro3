Option Strict Off
Option Explicit On 
Imports System.Windows.Forms.Application
Imports DAO.DBEngineClass
Imports System.Drawing.Printing

Friend Class frmRental
    Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
    Public Sub New()
        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()
        oDA = New CDataAccess()
        oCG = New CGrid()
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
    Public WithEvents cboEquipType As System.Windows.Forms.ComboBox

    Friend WithEvents dbgEquipment As System.Windows.Forms.DataGrid
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents fraSearchCriteria As System.Windows.Forms.GroupBox
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents lblFieldLable As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddToCart As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCancelOrder As System.Windows.Forms.MenuItem
    Public WithEvents mnuCategoryMaintenance As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCheckin As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCheckOut As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCustMaint As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCustomerStatements As System.Windows.Forms.MenuItem
    Public WithEvents mnuEdit As System.Windows.Forms.MenuItem
    Public WithEvents mnuEquipmentMaintenance As System.Windows.Forms.MenuItem
    Public WithEvents mnuFile As System.Windows.Forms.MenuItem
    Public WithEvents mnuFileExit As System.Windows.Forms.MenuItem
    Public WithEvents mnuFileSep1 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuHelp As System.Windows.Forms.MenuItem
    Friend WithEvents mnuHelpAbout As System.Windows.Forms.MenuItem
    Public WithEvents mnuRatesMaintenance As System.Windows.Forms.MenuItem
    Friend WithEvents mnuReceivePayments As System.Windows.Forms.MenuItem
    Friend WithEvents mnuRepairCompact As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSelectItems As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTools As System.Windows.Forms.MenuItem
    Public WithEvents optShowAll As System.Windows.Forms.RadioButton
    Public WithEvents optShowAvailable As System.Windows.Forms.RadioButton
    Public MainMenu1 As System.Windows.Forms.MainMenu
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    Friend WithEvents mnuPreferences As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPreviewBeforePrint As System.Windows.Forms.MenuItem
    Friend WithEvents mnuShowOpenItems As System.Windows.Forms.MenuItem
    Friend WithEvents mnuShowAllInvoices As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPrintInvoices As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPrintStatements As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPreviewStatements As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPreviewARReport As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPrintARReport As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCreditDebitMemo As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuLaborMaint As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSalesTaxReport As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEquipUsageReport As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSetupConfig As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem9 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuUpdateEquipCost As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem10 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPrintAged As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPreviewAged As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMeterReadingMaint As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExport As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExportCustomer As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExportInvoices As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExportInvoiceDetails As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExportEquipList As System.Windows.Forms.MenuItem
    Friend WithEvents mnuReRentReport As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTCOReport As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTCOPreviewOnID As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTCOPrintonID As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem11 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem12 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTCOPreviewonName As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTCOPrinttonName As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSelectDatabase As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem13 As System.Windows.Forms.MenuItem
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents mnuCheckDueItems As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEquipReport As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEqupReportPreview As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEquipReportPrint As System.Windows.Forms.MenuItem
    Friend WithEvents mnuRefreshGrid As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAutoRunReminder As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem15 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem16 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDeliveryReport As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem8 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuLicensingAgreement As System.Windows.Forms.MenuItem
    Friend WithEvents mnuProductInventoryReport As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEquipmentListRates As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPreviewEqListRates As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPrintEqListRates As System.Windows.Forms.MenuItem
    Public WithEvents Button1 As System.Windows.Forms.Button
    Public WithEvents cmdRentThis As System.Windows.Forms.Button
    Public WithEvents cmdCheckIn As System.Windows.Forms.Button
    Public WithEvents cmdCancelOrder As System.Windows.Forms.Button
    Public WithEvents cmdWhoHasIt As System.Windows.Forms.Button
    Public WithEvents cmdReserve As System.Windows.Forms.Button
    Public WithEvents btnClose As System.Windows.Forms.Button
    Public WithEvents btnReRent As System.Windows.Forms.Button
    Friend WithEvents mnuFuelCostsMaintenance As System.Windows.Forms.MenuItem
    Friend WithEvents mnuConfigMaintenance As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem17 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuShowPrinters As System.Windows.Forms.MenuItem
    Friend WithEvents PrintDialog1 As System.Windows.Forms.PrintDialog
    Friend WithEvents PrintDocument1 As System.Drawing.Printing.PrintDocument
    Friend WithEvents mnuEmployeeMaintenance As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPrinters As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSendEmail As System.Windows.Forms.MenuItem
    Friend WithEvents mnuShowRentedAndDue As System.Windows.Forms.MenuItem
    Friend WithEvents mnuShowDBLocation As MenuItem
    Friend WithEvents StatusBar1 As System.Windows.Forms.StatusBar
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmRental))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Button1 = New System.Windows.Forms.Button()
        Me.cmdRentThis = New System.Windows.Forms.Button()
        Me.cmdCheckIn = New System.Windows.Forms.Button()
        Me.cmdCancelOrder = New System.Windows.Forms.Button()
        Me.cmdWhoHasIt = New System.Windows.Forms.Button()
        Me.cmdReserve = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnReRent = New System.Windows.Forms.Button()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.optShowAvailable = New System.Windows.Forms.RadioButton()
        Me.optShowAll = New System.Windows.Forms.RadioButton()
        Me.fraSearchCriteria = New System.Windows.Forms.GroupBox()
        Me.cboEquipType = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblFieldLable = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MainMenu(Me.components)
        Me.mnuFile = New System.Windows.Forms.MenuItem()
        Me.mnuRefreshGrid = New System.Windows.Forms.MenuItem()
        Me.mnuCheckDueItems = New System.Windows.Forms.MenuItem()
        Me.mnuSendEmail = New System.Windows.Forms.MenuItem()
        Me.mnuExport = New System.Windows.Forms.MenuItem()
        Me.mnuExportCustomer = New System.Windows.Forms.MenuItem()
        Me.mnuExportInvoices = New System.Windows.Forms.MenuItem()
        Me.mnuExportInvoiceDetails = New System.Windows.Forms.MenuItem()
        Me.mnuExportEquipList = New System.Windows.Forms.MenuItem()
        Me.mnuFileSep1 = New System.Windows.Forms.MenuItem()
        Me.mnuSelectDatabase = New System.Windows.Forms.MenuItem()
        Me.mnuShowDBLocation = New System.Windows.Forms.MenuItem()
        Me.MenuItem13 = New System.Windows.Forms.MenuItem()
        Me.mnuPrinters = New System.Windows.Forms.MenuItem()
        Me.mnuShowPrinters = New System.Windows.Forms.MenuItem()
        Me.mnuFileExit = New System.Windows.Forms.MenuItem()
        Me.MenuItem1 = New System.Windows.Forms.MenuItem()
        Me.mnuReceivePayments = New System.Windows.Forms.MenuItem()
        Me.mnuCreditDebitMemo = New System.Windows.Forms.MenuItem()
        Me.mnuShowAllInvoices = New System.Windows.Forms.MenuItem()
        Me.MenuItem7 = New System.Windows.Forms.MenuItem()
        Me.MenuItem5 = New System.Windows.Forms.MenuItem()
        Me.MenuItem9 = New System.Windows.Forms.MenuItem()
        Me.mnuUpdateEquipCost = New System.Windows.Forms.MenuItem()
        Me.mnuMeterReadingMaint = New System.Windows.Forms.MenuItem()
        Me.mnuShowRentedAndDue = New System.Windows.Forms.MenuItem()
        Me.MenuItem16 = New System.Windows.Forms.MenuItem()
        Me.MenuItem10 = New System.Windows.Forms.MenuItem()
        Me.mnuPrintAged = New System.Windows.Forms.MenuItem()
        Me.mnuPreviewAged = New System.Windows.Forms.MenuItem()
        Me.MenuItem4 = New System.Windows.Forms.MenuItem()
        Me.mnuPreviewARReport = New System.Windows.Forms.MenuItem()
        Me.mnuPrintARReport = New System.Windows.Forms.MenuItem()
        Me.mnuDeliveryReport = New System.Windows.Forms.MenuItem()
        Me.mnuEquipReport = New System.Windows.Forms.MenuItem()
        Me.mnuEqupReportPreview = New System.Windows.Forms.MenuItem()
        Me.mnuEquipReportPrint = New System.Windows.Forms.MenuItem()
        Me.mnuEquipmentListRates = New System.Windows.Forms.MenuItem()
        Me.mnuPreviewEqListRates = New System.Windows.Forms.MenuItem()
        Me.mnuPrintEqListRates = New System.Windows.Forms.MenuItem()
        Me.mnuEquipUsageReport = New System.Windows.Forms.MenuItem()
        Me.mnuReRentReport = New System.Windows.Forms.MenuItem()
        Me.MenuItem15 = New System.Windows.Forms.MenuItem()
        Me.mnuPrintInvoices = New System.Windows.Forms.MenuItem()
        Me.mnuCustomerStatements = New System.Windows.Forms.MenuItem()
        Me.mnuPrintStatements = New System.Windows.Forms.MenuItem()
        Me.mnuPreviewStatements = New System.Windows.Forms.MenuItem()
        Me.mnuShowOpenItems = New System.Windows.Forms.MenuItem()
        Me.mnuSalesTaxReport = New System.Windows.Forms.MenuItem()
        Me.mnuTCOReport = New System.Windows.Forms.MenuItem()
        Me.MenuItem11 = New System.Windows.Forms.MenuItem()
        Me.mnuTCOPreviewOnID = New System.Windows.Forms.MenuItem()
        Me.mnuTCOPrintonID = New System.Windows.Forms.MenuItem()
        Me.MenuItem12 = New System.Windows.Forms.MenuItem()
        Me.mnuTCOPreviewonName = New System.Windows.Forms.MenuItem()
        Me.mnuTCOPrinttonName = New System.Windows.Forms.MenuItem()
        Me.mnuProductInventoryReport = New System.Windows.Forms.MenuItem()
        Me.mnuPreferences = New System.Windows.Forms.MenuItem()
        Me.mnuPreviewBeforePrint = New System.Windows.Forms.MenuItem()
        Me.mnuAutoRunReminder = New System.Windows.Forms.MenuItem()
        Me.MenuItem6 = New System.Windows.Forms.MenuItem()
        Me.mnuSetupConfig = New System.Windows.Forms.MenuItem()
        Me.mnuEdit = New System.Windows.Forms.MenuItem()
        Me.mnuRatesMaintenance = New System.Windows.Forms.MenuItem()
        Me.mnuEquipmentMaintenance = New System.Windows.Forms.MenuItem()
        Me.mnuCategoryMaintenance = New System.Windows.Forms.MenuItem()
        Me.MenuItem2 = New System.Windows.Forms.MenuItem()
        Me.mnuCustMaint = New System.Windows.Forms.MenuItem()
        Me.mnuFuelCostsMaintenance = New System.Windows.Forms.MenuItem()
        Me.mnuLaborMaint = New System.Windows.Forms.MenuItem()
        Me.mnuEmployeeMaintenance = New System.Windows.Forms.MenuItem()
        Me.MenuItem3 = New System.Windows.Forms.MenuItem()
        Me.mnuConfigMaintenance = New System.Windows.Forms.MenuItem()
        Me.MenuItem17 = New System.Windows.Forms.MenuItem()
        Me.mnuRepairCompact = New System.Windows.Forms.MenuItem()
        Me.mnuHelp = New System.Windows.Forms.MenuItem()
        Me.mnuHelpAbout = New System.Windows.Forms.MenuItem()
        Me.mnuLicensingAgreement = New System.Windows.Forms.MenuItem()
        Me.MenuItem8 = New System.Windows.Forms.MenuItem()
        Me.mnuSelectItems = New System.Windows.Forms.MenuItem()
        Me.mnuAddToCart = New System.Windows.Forms.MenuItem()
        Me.mnuCheckOut = New System.Windows.Forms.MenuItem()
        Me.mnuCheckin = New System.Windows.Forms.MenuItem()
        Me.mnuTools = New System.Windows.Forms.MenuItem()
        Me.mnuCancelOrder = New System.Windows.Forms.MenuItem()
        Me.dbgEquipment = New System.Windows.Forms.DataGrid()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.StatusBar1 = New System.Windows.Forms.StatusBar()
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog()
        Me.PrintDocument1 = New System.Drawing.Printing.PrintDocument()
        Me.Frame2.SuspendLayout()
        Me.fraSearchCriteria.SuspendLayout()
        CType(Me.lblFieldLable, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dbgEquipment, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.SystemColors.Control
        Me.Button1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Button1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Button1.Location = New System.Drawing.Point(19, 98)
        Me.Button1.Name = "Button1"
        Me.Button1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Button1.Size = New System.Drawing.Size(139, 102)
        Me.Button1.TabIndex = 23
        Me.Button1.TabStop = False
        Me.Button1.Text = "&Supplies"
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.Button1, "Show tools and supplies for sale")
        Me.Button1.UseVisualStyleBackColor = False
        '
        'cmdRentThis
        '
        Me.cmdRentThis.BackColor = System.Drawing.SystemColors.Control
        Me.cmdRentThis.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdRentThis.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRentThis.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdRentThis.Image = CType(resources.GetObject("cmdRentThis.Image"), System.Drawing.Image)
        Me.cmdRentThis.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdRentThis.Location = New System.Drawing.Point(158, 98)
        Me.cmdRentThis.Name = "cmdRentThis"
        Me.cmdRentThis.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdRentThis.Size = New System.Drawing.Size(140, 102)
        Me.cmdRentThis.TabIndex = 24
        Me.cmdRentThis.TabStop = False
        Me.cmdRentThis.Text = "Check &Out"
        Me.cmdRentThis.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.cmdRentThis, "Ring me up, I've gotta go to work!!!!!!!!!")
        Me.cmdRentThis.UseVisualStyleBackColor = False
        '
        'cmdCheckIn
        '
        Me.cmdCheckIn.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCheckIn.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCheckIn.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCheckIn.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCheckIn.Image = CType(resources.GetObject("cmdCheckIn.Image"), System.Drawing.Image)
        Me.cmdCheckIn.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdCheckIn.Location = New System.Drawing.Point(298, 98)
        Me.cmdCheckIn.Name = "cmdCheckIn"
        Me.cmdCheckIn.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCheckIn.Size = New System.Drawing.Size(139, 102)
        Me.cmdCheckIn.TabIndex = 25
        Me.cmdCheckIn.TabStop = False
        Me.cmdCheckIn.Text = "Check &In"
        Me.cmdCheckIn.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.cmdCheckIn, "Check in returned equipment")
        Me.cmdCheckIn.UseVisualStyleBackColor = False
        '
        'cmdCancelOrder
        '
        Me.cmdCancelOrder.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancelOrder.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancelOrder.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancelOrder.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancelOrder.Image = CType(resources.GetObject("cmdCancelOrder.Image"), System.Drawing.Image)
        Me.cmdCancelOrder.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdCancelOrder.Location = New System.Drawing.Point(576, 98)
        Me.cmdCancelOrder.Name = "cmdCancelOrder"
        Me.cmdCancelOrder.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancelOrder.Size = New System.Drawing.Size(139, 102)
        Me.cmdCancelOrder.TabIndex = 26
        Me.cmdCancelOrder.TabStop = False
        Me.cmdCancelOrder.Text = "&Cancel"
        Me.cmdCancelOrder.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.cmdCancelOrder, "Cancel customer's list")
        Me.cmdCancelOrder.UseVisualStyleBackColor = False
        '
        'cmdWhoHasIt
        '
        Me.cmdWhoHasIt.BackColor = System.Drawing.SystemColors.Control
        Me.cmdWhoHasIt.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdWhoHasIt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdWhoHasIt.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdWhoHasIt.Image = CType(resources.GetObject("cmdWhoHasIt.Image"), System.Drawing.Image)
        Me.cmdWhoHasIt.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdWhoHasIt.Location = New System.Drawing.Point(437, 98)
        Me.cmdWhoHasIt.Name = "cmdWhoHasIt"
        Me.cmdWhoHasIt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdWhoHasIt.Size = New System.Drawing.Size(139, 102)
        Me.cmdWhoHasIt.TabIndex = 27
        Me.cmdWhoHasIt.TabStop = False
        Me.cmdWhoHasIt.Text = "&Who Has It?"
        Me.cmdWhoHasIt.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.cmdWhoHasIt, "Show who has equipment checked out")
        Me.cmdWhoHasIt.UseVisualStyleBackColor = False
        '
        'cmdReserve
        '
        Me.cmdReserve.BackColor = System.Drawing.SystemColors.Control
        Me.cmdReserve.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdReserve.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdReserve.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdReserve.Image = CType(resources.GetObject("cmdReserve.Image"), System.Drawing.Image)
        Me.cmdReserve.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdReserve.Location = New System.Drawing.Point(715, 98)
        Me.cmdReserve.Name = "cmdReserve"
        Me.cmdReserve.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdReserve.Size = New System.Drawing.Size(139, 102)
        Me.cmdReserve.TabIndex = 28
        Me.cmdReserve.TabStop = False
        Me.cmdReserve.Text = "&Reserve"
        Me.cmdReserve.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.cmdReserve, "Look at future reservations and usage")
        Me.cmdReserve.UseVisualStyleBackColor = False
        '
        'btnClose
        '
        Me.btnClose.BackColor = System.Drawing.SystemColors.Control
        Me.btnClose.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnClose.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnClose.Image = CType(resources.GetObject("btnClose.Image"), System.Drawing.Image)
        Me.btnClose.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnClose.Location = New System.Drawing.Point(994, 98)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnClose.Size = New System.Drawing.Size(139, 102)
        Me.btnClose.TabIndex = 29
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "C&lose"
        Me.btnClose.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.btnClose, "Close down and go HOME!")
        Me.btnClose.UseVisualStyleBackColor = False
        '
        'btnReRent
        '
        Me.btnReRent.BackColor = System.Drawing.SystemColors.Control
        Me.btnReRent.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnReRent.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnReRent.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnReRent.Image = CType(resources.GetObject("btnReRent.Image"), System.Drawing.Image)
        Me.btnReRent.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnReRent.Location = New System.Drawing.Point(854, 98)
        Me.btnReRent.Name = "btnReRent"
        Me.btnReRent.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnReRent.Size = New System.Drawing.Size(140, 102)
        Me.btnReRent.TabIndex = 30
        Me.btnReRent.TabStop = False
        Me.btnReRent.Text = "&ReRent"
        Me.btnReRent.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.btnReRent, "Show tools and supplies for sale")
        Me.btnReRent.UseVisualStyleBackColor = False
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.optShowAvailable)
        Me.Frame2.Controls.Add(Me.optShowAll)
        Me.Frame2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(331, 6)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(296, 79)
        Me.Frame2.TabIndex = 10
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Availability"
        '
        'optShowAvailable
        '
        Me.optShowAvailable.BackColor = System.Drawing.SystemColors.Control
        Me.optShowAvailable.Checked = True
        Me.optShowAvailable.Cursor = System.Windows.Forms.Cursors.Default
        Me.optShowAvailable.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optShowAvailable.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optShowAvailable.Location = New System.Drawing.Point(19, 23)
        Me.optShowAvailable.Name = "optShowAvailable"
        Me.optShowAvailable.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optShowAvailable.Size = New System.Drawing.Size(114, 37)
        Me.optShowAvailable.TabIndex = 0
        Me.optShowAvailable.TabStop = True
        Me.optShowAvailable.Text = "Show A&vailable"
        Me.optShowAvailable.UseVisualStyleBackColor = False
        '
        'optShowAll
        '
        Me.optShowAll.BackColor = System.Drawing.SystemColors.Control
        Me.optShowAll.Cursor = System.Windows.Forms.Cursors.Default
        Me.optShowAll.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optShowAll.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optShowAll.Location = New System.Drawing.Point(154, 23)
        Me.optShowAll.Name = "optShowAll"
        Me.optShowAll.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optShowAll.Size = New System.Drawing.Size(128, 28)
        Me.optShowAll.TabIndex = 1
        Me.optShowAll.Text = "Show &All"
        Me.optShowAll.UseVisualStyleBackColor = False
        '
        'fraSearchCriteria
        '
        Me.fraSearchCriteria.BackColor = System.Drawing.SystemColors.Control
        Me.fraSearchCriteria.Controls.Add(Me.cboEquipType)
        Me.fraSearchCriteria.Controls.Add(Me.Label1)
        Me.fraSearchCriteria.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraSearchCriteria.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraSearchCriteria.Location = New System.Drawing.Point(11, 6)
        Me.fraSearchCriteria.Name = "fraSearchCriteria"
        Me.fraSearchCriteria.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraSearchCriteria.Size = New System.Drawing.Size(309, 80)
        Me.fraSearchCriteria.TabIndex = 3
        Me.fraSearchCriteria.TabStop = False
        Me.fraSearchCriteria.Text = "Category Search"
        '
        'cboEquipType
        '
        Me.cboEquipType.BackColor = System.Drawing.SystemColors.Window
        Me.cboEquipType.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboEquipType.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboEquipType.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboEquipType.Location = New System.Drawing.Point(61, 26)
        Me.cboEquipType.Name = "cboEquipType"
        Me.cboEquipType.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboEquipType.Size = New System.Drawing.Size(235, 26)
        Me.cboEquipType.TabIndex = 0
        Me.cboEquipType.TabStop = False
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(10, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(80, 43)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Equip &Type"
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFile, Me.MenuItem1, Me.MenuItem9, Me.MenuItem16, Me.mnuPreferences, Me.mnuEdit, Me.mnuHelp})
        '
        'mnuFile
        '
        Me.mnuFile.Index = 0
        Me.mnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuRefreshGrid, Me.mnuCheckDueItems, Me.mnuSendEmail, Me.mnuExport, Me.mnuFileSep1, Me.mnuSelectDatabase, Me.mnuShowDBLocation, Me.MenuItem13, Me.mnuPrinters, Me.mnuShowPrinters, Me.mnuFileExit})
        Me.mnuFile.Text = "&File"
        '
        'mnuRefreshGrid
        '
        Me.mnuRefreshGrid.Index = 0
        Me.mnuRefreshGrid.Shortcut = System.Windows.Forms.Shortcut.F2
        Me.mnuRefreshGrid.Text = "&Refresh Grid"
        '
        'mnuCheckDueItems
        '
        Me.mnuCheckDueItems.Index = 1
        Me.mnuCheckDueItems.Shortcut = System.Windows.Forms.Shortcut.CtrlD
        Me.mnuCheckDueItems.Text = "Check Due Items"
        '
        'mnuSendEmail
        '
        Me.mnuSendEmail.Index = 2
        Me.mnuSendEmail.Text = "Send Email"
        '
        'mnuExport
        '
        Me.mnuExport.Index = 3
        Me.mnuExport.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuExportCustomer, Me.mnuExportInvoices, Me.mnuExportInvoiceDetails, Me.mnuExportEquipList})
        Me.mnuExport.Text = "Ex&port to Quik Books"
        Me.mnuExport.Visible = False
        '
        'mnuExportCustomer
        '
        Me.mnuExportCustomer.Index = 0
        Me.mnuExportCustomer.Text = "Customer List"
        '
        'mnuExportInvoices
        '
        Me.mnuExportInvoices.Index = 1
        Me.mnuExportInvoices.Text = "Invoice Header"
        '
        'mnuExportInvoiceDetails
        '
        Me.mnuExportInvoiceDetails.Index = 2
        Me.mnuExportInvoiceDetails.Text = "Invoice Details"
        '
        'mnuExportEquipList
        '
        Me.mnuExportEquipList.Index = 3
        Me.mnuExportEquipList.Text = "Equipment List"
        '
        'mnuFileSep1
        '
        Me.mnuFileSep1.Index = 4
        Me.mnuFileSep1.Text = "-"
        '
        'mnuSelectDatabase
        '
        Me.mnuSelectDatabase.Index = 5
        Me.mnuSelectDatabase.Text = "Select &Database"
        '
        'mnuShowDBLocation
        '
        Me.mnuShowDBLocation.Index = 6
        Me.mnuShowDBLocation.Text = "Show Database Location"
        '
        'MenuItem13
        '
        Me.MenuItem13.Index = 7
        Me.MenuItem13.Text = "-"
        '
        'mnuPrinters
        '
        Me.mnuPrinters.Enabled = False
        Me.mnuPrinters.Index = 8
        Me.mnuPrinters.Text = "Set Default Printer"
        '
        'mnuShowPrinters
        '
        Me.mnuShowPrinters.Index = 9
        Me.mnuShowPrinters.Text = "Page Setup"
        '
        'mnuFileExit
        '
        Me.mnuFileExit.Index = 10
        Me.mnuFileExit.Shortcut = System.Windows.Forms.Shortcut.CtrlE
        Me.mnuFileExit.Text = "E&xit"
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 1
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuReceivePayments, Me.mnuCreditDebitMemo, Me.mnuShowAllInvoices, Me.MenuItem7, Me.MenuItem5})
        Me.MenuItem1.Text = "&Customers"
        '
        'mnuReceivePayments
        '
        Me.mnuReceivePayments.Index = 0
        Me.mnuReceivePayments.Text = "&Receive Payments"
        '
        'mnuCreditDebitMemo
        '
        Me.mnuCreditDebitMemo.Index = 1
        Me.mnuCreditDebitMemo.Text = "Credit/&Debit Customer Account"
        '
        'mnuShowAllInvoices
        '
        Me.mnuShowAllInvoices.Index = 2
        Me.mnuShowAllInvoices.Text = "Show &All Invoices"
        Me.mnuShowAllInvoices.Visible = False
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 3
        Me.MenuItem7.Text = "-"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 4
        Me.MenuItem5.Text = "&View Customer Data"
        '
        'MenuItem9
        '
        Me.MenuItem9.Index = 2
        Me.MenuItem9.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuUpdateEquipCost, Me.mnuMeterReadingMaint, Me.mnuShowRentedAndDue})
        Me.MenuItem9.Text = "&Equipment"
        '
        'mnuUpdateEquipCost
        '
        Me.mnuUpdateEquipCost.Index = 0
        Me.mnuUpdateEquipCost.Text = "&Ownership Costs"
        '
        'mnuMeterReadingMaint
        '
        Me.mnuMeterReadingMaint.Index = 1
        Me.mnuMeterReadingMaint.Text = "&Meter Reading Maintenance"
        '
        'mnuShowRentedAndDue
        '
        Me.mnuShowRentedAndDue.Index = 2
        Me.mnuShowRentedAndDue.Text = "Show Rented && Due"
        '
        'MenuItem16
        '
        Me.MenuItem16.Index = 3
        Me.MenuItem16.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem10, Me.MenuItem4, Me.mnuDeliveryReport, Me.mnuEquipReport, Me.mnuEquipmentListRates, Me.mnuEquipUsageReport, Me.mnuReRentReport, Me.MenuItem15, Me.mnuPrintInvoices, Me.mnuCustomerStatements, Me.mnuShowOpenItems, Me.mnuSalesTaxReport, Me.mnuTCOReport, Me.mnuProductInventoryReport})
        Me.MenuItem16.Text = "&Reports"
        '
        'MenuItem10
        '
        Me.MenuItem10.Index = 0
        Me.MenuItem10.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuPrintAged, Me.mnuPreviewAged})
        Me.MenuItem10.Text = "Aged &Receivables Report"
        '
        'mnuPrintAged
        '
        Me.mnuPrintAged.Index = 0
        Me.mnuPrintAged.Text = "P&rint"
        '
        'mnuPreviewAged
        '
        Me.mnuPreviewAged.Index = 1
        Me.mnuPreviewAged.Text = "Pre&view"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 1
        Me.MenuItem4.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuPreviewARReport, Me.mnuPrintARReport})
        Me.MenuItem4.Text = "&Accounts Receivables Report"
        '
        'mnuPreviewARReport
        '
        Me.mnuPreviewARReport.Index = 0
        Me.mnuPreviewARReport.Text = "Pre&view"
        '
        'mnuPrintARReport
        '
        Me.mnuPrintARReport.Index = 1
        Me.mnuPrintARReport.Text = "&Print"
        '
        'mnuDeliveryReport
        '
        Me.mnuDeliveryReport.Index = 2
        Me.mnuDeliveryReport.Text = "Delivery Cost Report"
        '
        'mnuEquipReport
        '
        Me.mnuEquipReport.Index = 3
        Me.mnuEquipReport.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuEqupReportPreview, Me.mnuEquipReportPrint})
        Me.mnuEquipReport.Text = "Equipment List"
        '
        'mnuEqupReportPreview
        '
        Me.mnuEqupReportPreview.Index = 0
        Me.mnuEqupReportPreview.Text = "Pre&view"
        '
        'mnuEquipReportPrint
        '
        Me.mnuEquipReportPrint.Index = 1
        Me.mnuEquipReportPrint.Text = "&Print"
        '
        'mnuEquipmentListRates
        '
        Me.mnuEquipmentListRates.Index = 4
        Me.mnuEquipmentListRates.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuPreviewEqListRates, Me.mnuPrintEqListRates})
        Me.mnuEquipmentListRates.Text = "Equipment List/Rates"
        '
        'mnuPreviewEqListRates
        '
        Me.mnuPreviewEqListRates.Index = 0
        Me.mnuPreviewEqListRates.Text = "Preview"
        '
        'mnuPrintEqListRates
        '
        Me.mnuPrintEqListRates.Index = 1
        Me.mnuPrintEqListRates.Text = "Print"
        '
        'mnuEquipUsageReport
        '
        Me.mnuEquipUsageReport.Index = 5
        Me.mnuEquipUsageReport.Text = "&Equipment Usage Report"
        '
        'mnuReRentReport
        '
        Me.mnuReRentReport.Index = 6
        Me.mnuReRentReport.Text = "&ReRent Report"
        '
        'MenuItem15
        '
        Me.MenuItem15.Index = 7
        Me.MenuItem15.Text = "Meter Reading Report"
        '
        'mnuPrintInvoices
        '
        Me.mnuPrintInvoices.Index = 8
        Me.mnuPrintInvoices.Text = "&Print Invoices"
        '
        'mnuCustomerStatements
        '
        Me.mnuCustomerStatements.Index = 9
        Me.mnuCustomerStatements.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuPrintStatements, Me.mnuPreviewStatements})
        Me.mnuCustomerStatements.Text = "&Print All Customer Statements"
        '
        'mnuPrintStatements
        '
        Me.mnuPrintStatements.Index = 0
        Me.mnuPrintStatements.Text = "&Print"
        '
        'mnuPreviewStatements
        '
        Me.mnuPreviewStatements.Index = 1
        Me.mnuPreviewStatements.Text = "Preview"
        '
        'mnuShowOpenItems
        '
        Me.mnuShowOpenItems.Index = 10
        Me.mnuShowOpenItems.Text = "View/Print &Single Statement"
        '
        'mnuSalesTaxReport
        '
        Me.mnuSalesTaxReport.Index = 11
        Me.mnuSalesTaxReport.Text = "&Sales && Tax Report"
        '
        'mnuTCOReport
        '
        Me.mnuTCOReport.Index = 12
        Me.mnuTCOReport.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem11, Me.MenuItem12})
        Me.mnuTCOReport.Text = "&Total Cost Ownership Report"
        '
        'MenuItem11
        '
        Me.MenuItem11.Index = 0
        Me.MenuItem11.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuTCOPreviewOnID, Me.mnuTCOPrintonID})
        Me.MenuItem11.Text = "Sort By Equip ID"
        '
        'mnuTCOPreviewOnID
        '
        Me.mnuTCOPreviewOnID.Index = 0
        Me.mnuTCOPreviewOnID.Text = "Preview"
        '
        'mnuTCOPrintonID
        '
        Me.mnuTCOPrintonID.Index = 1
        Me.mnuTCOPrintonID.Text = "Print"
        '
        'MenuItem12
        '
        Me.MenuItem12.Index = 1
        Me.MenuItem12.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuTCOPreviewonName, Me.mnuTCOPrinttonName})
        Me.MenuItem12.Text = "Sort By Equip Name"
        '
        'mnuTCOPreviewonName
        '
        Me.mnuTCOPreviewonName.Index = 0
        Me.mnuTCOPreviewonName.Text = "Preview"
        '
        'mnuTCOPrinttonName
        '
        Me.mnuTCOPrinttonName.Index = 1
        Me.mnuTCOPrinttonName.Text = "Print"
        '
        'mnuProductInventoryReport
        '
        Me.mnuProductInventoryReport.Index = 13
        Me.mnuProductInventoryReport.Text = "Product Inventory Report"
        '
        'mnuPreferences
        '
        Me.mnuPreferences.Index = 4
        Me.mnuPreferences.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuPreviewBeforePrint, Me.mnuAutoRunReminder, Me.MenuItem6, Me.mnuSetupConfig})
        Me.mnuPreferences.Text = "&Preferences"
        '
        'mnuPreviewBeforePrint
        '
        Me.mnuPreviewBeforePrint.Checked = True
        Me.mnuPreviewBeforePrint.Index = 0
        Me.mnuPreviewBeforePrint.Text = "&Preview Before Print"
        '
        'mnuAutoRunReminder
        '
        Me.mnuAutoRunReminder.Index = 1
        Me.mnuAutoRunReminder.Text = "Auto Run Daily Reminder"
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 2
        Me.MenuItem6.Text = "-"
        '
        'mnuSetupConfig
        '
        Me.mnuSetupConfig.Index = 3
        Me.mnuSetupConfig.Text = "Setup Configuration"
        '
        'mnuEdit
        '
        Me.mnuEdit.Index = 5
        Me.mnuEdit.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuRatesMaintenance, Me.mnuEquipmentMaintenance, Me.mnuCategoryMaintenance, Me.MenuItem2, Me.mnuCustMaint, Me.mnuFuelCostsMaintenance, Me.mnuLaborMaint, Me.mnuEmployeeMaintenance, Me.MenuItem3, Me.mnuConfigMaintenance, Me.MenuItem17, Me.mnuRepairCompact})
        Me.mnuEdit.Text = "&Database Maintenance"
        '
        'mnuRatesMaintenance
        '
        Me.mnuRatesMaintenance.Index = 0
        Me.mnuRatesMaintenance.Text = "&Rates Maintenance"
        '
        'mnuEquipmentMaintenance
        '
        Me.mnuEquipmentMaintenance.Index = 1
        Me.mnuEquipmentMaintenance.Text = "&Equipment Maintenance"
        '
        'mnuCategoryMaintenance
        '
        Me.mnuCategoryMaintenance.Index = 2
        Me.mnuCategoryMaintenance.Text = "&Category Maintenance"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 3
        Me.MenuItem2.Text = "Edit &Sale Items"
        '
        'mnuCustMaint
        '
        Me.mnuCustMaint.Index = 4
        Me.mnuCustMaint.Text = "C&ustomer Maintenance"
        '
        'mnuFuelCostsMaintenance
        '
        Me.mnuFuelCostsMaintenance.Index = 5
        Me.mnuFuelCostsMaintenance.Text = "Fuel Costs Maintenance"
        '
        'mnuLaborMaint
        '
        Me.mnuLaborMaint.Index = 6
        Me.mnuLaborMaint.Text = "&Labor Charges Maintenance"
        '
        'mnuEmployeeMaintenance
        '
        Me.mnuEmployeeMaintenance.Index = 7
        Me.mnuEmployeeMaintenance.Text = "E&mployee Maintenance"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 8
        Me.MenuItem3.Text = "-"
        '
        'mnuConfigMaintenance
        '
        Me.mnuConfigMaintenance.Index = 9
        Me.mnuConfigMaintenance.Text = "System Configuration Maintenance"
        '
        'MenuItem17
        '
        Me.MenuItem17.Index = 10
        Me.MenuItem17.Text = "-"
        '
        'mnuRepairCompact
        '
        Me.mnuRepairCompact.Index = 11
        Me.mnuRepairCompact.Text = "Re&pair and Compact"
        '
        'mnuHelp
        '
        Me.mnuHelp.Index = 6
        Me.mnuHelp.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuHelpAbout, Me.mnuLicensingAgreement, Me.MenuItem8, Me.mnuSelectItems, Me.mnuAddToCart, Me.mnuCheckOut, Me.mnuCheckin, Me.mnuTools, Me.mnuCancelOrder})
        Me.mnuHelp.Text = "&Help"
        '
        'mnuHelpAbout
        '
        Me.mnuHelpAbout.Index = 0
        Me.mnuHelpAbout.Text = "Help About"
        '
        'mnuLicensingAgreement
        '
        Me.mnuLicensingAgreement.Index = 1
        Me.mnuLicensingAgreement.Text = "Licenseing Agreement"
        '
        'MenuItem8
        '
        Me.MenuItem8.Index = 2
        Me.MenuItem8.Text = "-"
        '
        'mnuSelectItems
        '
        Me.mnuSelectItems.Index = 3
        Me.mnuSelectItems.Text = "Selecting Items to Rent"
        '
        'mnuAddToCart
        '
        Me.mnuAddToCart.Index = 4
        Me.mnuAddToCart.Text = "Add To Cart"
        '
        'mnuCheckOut
        '
        Me.mnuCheckOut.Index = 5
        Me.mnuCheckOut.Text = "Check Out"
        '
        'mnuCheckin
        '
        Me.mnuCheckin.Index = 6
        Me.mnuCheckin.Text = "Check In"
        '
        'mnuTools
        '
        Me.mnuTools.Index = 7
        Me.mnuTools.Text = "Tools & Supplies"
        '
        'mnuCancelOrder
        '
        Me.mnuCancelOrder.Index = 8
        Me.mnuCancelOrder.Text = "Cancel Order"
        '
        'dbgEquipment
        '
        Me.dbgEquipment.AllowSorting = False
        Me.dbgEquipment.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dbgEquipment.CaptionText = "Select Items to Rent"
        Me.dbgEquipment.DataMember = ""
        Me.dbgEquipment.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dbgEquipment.Location = New System.Drawing.Point(3, 206)
        Me.dbgEquipment.Name = "dbgEquipment"
        Me.dbgEquipment.RowHeadersVisible = False
        Me.dbgEquipment.Size = New System.Drawing.Size(722, 157)
        Me.dbgEquipment.TabIndex = 0
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.Filter = "Database Files (*.MDB)|*.MDB)"
        '
        'StatusBar1
        '
        Me.StatusBar1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StatusBar1.Location = New System.Drawing.Point(0, 372)
        Me.StatusBar1.Name = "StatusBar1"
        Me.StatusBar1.Size = New System.Drawing.Size(744, 32)
        Me.StatusBar1.TabIndex = 22
        Me.StatusBar1.Text = "StatusBar1"
        '
        'PrintDialog1
        '
        Me.PrintDialog1.UseEXDialog = True
        '
        'frmRental
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(8, 19)
        Me.ClientSize = New System.Drawing.Size(744, 404)
        Me.Controls.Add(Me.btnReRent)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.cmdReserve)
        Me.Controls.Add(Me.cmdWhoHasIt)
        Me.Controls.Add(Me.cmdCancelOrder)
        Me.Controls.Add(Me.cmdCheckIn)
        Me.Controls.Add(Me.cmdRentThis)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.StatusBar1)
        Me.Controls.Add(Me.dbgEquipment)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.fraSearchCriteria)
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(125, 138)
        Me.Menu = Me.MainMenu1
        Me.Name = "frmRental"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "RentalPro - Select Rental Equipment Version 4"
        Me.Frame2.ResumeLayout(False)
        Me.fraSearchCriteria.ResumeLayout(False)
        CType(Me.lblFieldLable, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dbgEquipment, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
#End Region

#Region " Module Variables "
    Dim SQL As String
    Dim mbDirty As Boolean
    Dim rsFC As New DataTable()
    Dim Noise As Boolean
    Dim LastSearchType As String
    Dim oDA As CDataAccess
    Public dtEquip As New DataTable()
    Private oCG As CGrid
    Public miHitRow As Integer
    Private msgMeter As String = "The meter reading for this equipment has never been entered into the database; please go to the Meter Reading maintenance screen and update the hours for this equipment."
    Private formLoading As Boolean


#End Region

#Region " Private Methods "
    ''' <summary>
    '''
    ''' </summary>
    ''' <param name = "eventSender"></param>
    Private Sub frmRental_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load


        Dim loRS As New DataTable()
        Dim i As Integer
        Dim oRES As CTransaction
        Dim AssyName As String = System.Reflection.Assembly.GetExecutingAssembly.GetName.Name.ToString
        formLoading = True
        Try
            Me.Left = VB6.TwipsToPixelsX(CSng(GetSetting(System.Reflection.Assembly.GetExecutingAssembly.GetName.Name, "Settings", "MainLeft", CStr(1000))))
            Me.Top = VB6.TwipsToPixelsY(CSng(GetSetting(System.Reflection.Assembly.GetExecutingAssembly.GetName.Name, "Settings", "MainTop", CStr(1000))))
            Me.Width = VB6.TwipsToPixelsX(CSng(GetSetting(System.Reflection.Assembly.GetExecutingAssembly.GetName.Name, "Settings", "MainWidth", CStr(6500))))
            Me.Height = VB6.TwipsToPixelsY(CSng(GetSetting(System.Reflection.Assembly.GetExecutingAssembly.GetName.Name, "Settings", "MainHeight", CStr(6500))))
            Me.mnuPreviewBeforePrint.Checked =
               GetSetting(System.Reflection.Assembly.GetExecutingAssembly.GetName.Name, SETTINGS, "Preview", True)
            Me.mnuAutoRunReminder.Checked = GetSetting(RENTALPRO, SETTINGS, "AUTOREMIND", True)

            ' load the equipment type combo
            SQL = "select * from equipment_type order by equip_type_id"
            oDA.SendQuery(SQL, loRS, modMain.ConnectString) ' return value ignored here
            oRES = New CTransaction()
            Call oRES.RemoveTempReservation("")

            Me.cboEquipType.Items.Clear()
            Me.cboEquipType.Items.Add(("All Equipment"))
            For i = 0 To loRS.Rows.Count - 1
                With loRS.Rows(i)
                    Me.cboEquipType.Items.Add(CType(.Item("equip_type_id"), String).PadLeft(2) &
                       " - " & .Item("equip_type"))
                End With
            Next i

            If Me.cboEquipType.Items.Count <> 0 Then
                Noise = True
                Me.cboEquipType.SelectedIndex = 0
                Me.cboEquipType.SelectionLength = 0
                Noise = False
            End If
            Me.LoadEquipGridFromType(0)
            Me.StatusBar1.Text = DatabaseName

            DoEvents()
            Me.dbgEquipment.Focus()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

    Private Sub frmRental_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyUp

        Try
            Dim KeyCode As Short = eventArgs.KeyCode
            Dim Shift As Short = eventArgs.KeyData \ &H10000
            Select Case Chr(KeyCode)
                'Case "A"
                '   AddItemsToCollection()
                Case "W"
                    Dim oFrm As New frmWhoHasIt()
                    oFrm.ShowDialog()
                    SendKeys.Send("{TAB}")
                Case "O"
                    Dim oFrm As New frmCustomers()
                    oFrm.ShowDialog()
                    SendKeys.Send("{TAB}")
                Case "I"
                    cmdCheckIn_Click(New Object(), New Object())
                    SendKeys.Send("{TAB}")
                Case "R"
                    oCG.SelectCkBoxRow(Me.dtEquip, Me.dbgEquipment, "RentMe")
                    Me.AddItemsToCollection(Me.dbgEquipment.CurrentRowIndex)
                    SendKeys.Send("{TAB}")
                Case "C"
                    Me.cmdCancelOrder_Click(New Object(), New Object())
                    'Case "L"
                    '   Me.cmdShowList_Click(New Object(), New Object())
                Case "S"
                    'me.dbgEquipment.SelBookmarks.Add(
                Case "T"
                    Button1_Click(New Object(), New Object())
                    SendKeys.Send("{TAB}")
            End Select
            Me.dbgEquipment.Focus()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Returns true if meter has ever been read, otherwise false.
    ''' </summary>
    ''' <param name = "HitRow"></param>
    ''' <returns>Single</returns>
    Private Function GetMeterReading(ByVal HitRow As Integer) As Single
        Dim dt As New DataTable()


        Try
            SQL = "select meter_reading,date_entered from meter_reading "
            SQL &= "where equip_id = '" & dbgEquipment.DataSource.rows(HitRow).item("equip_id") & "' "
            SQL &= "and meter_reading is not null "
            SQL &= "order by date_entered desc"
            If oDA.SendQuery(SQL, dt, ConnectString) > 0 Then
                Return dt.Rows(0).Item("meter_reading")
            Else
                Return 0
            End If
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Function

    ''' <summary>
    '''
    ''' </summary>
    ''' <param name = "HitRow"></param>
    Private Overloads Sub AddItemsToCollection(ByVal HitRow As Integer)
        Dim nTotal As Integer
        Dim nTotalSelRows As Short
        Dim i As Integer = HitRow
        Dim oItem As CItems
        Dim oFrm As frmAddToList
        Dim meterReading As Single

        Try
            If dbgEquipment.DataSource.rows(i).item("meter") = "Yes" Then 'meter_required
                meterReading = GetMeterReading(HitRow)
                If meterReading = 0 Then
                    dtEquip.Rows(dbgEquipment.CurrentCell.RowNumber).Item("RentMe") = "false"
                    MsgBox(msgMeter, MsgBoxStyle.Exclamation)
                    Exit Sub
                End If
            End If

            DoEvents()
            oFrm = New frmAddToList()
            DoEvents()
            With oFrm
                Dim dr As DataRow = Me.dbgEquipment.DataSource.rows(i)

                .ItemID = dr("equip_id") ' Me.dbgEquipment.DataSource.rows(i).item("equip_id")
                .ItemName = dr("equip_name") 'Me.dbgEquipment.DataSource.rows(i).item("equip_name")
                .PriceID = dr("price_id") 'Me.dbgEquipment.DataSource.rows(i).item("price_id")
                .MeterRequired = (dr("meter") = "Yes") 'dbgEquipment.DataSource.rows(i).item("meter_required")
                .MeterHours = meterReading
                .dailyRate = MND(dr("daily"))
                .halfDayRate = MND(dr("halfday"))
                .hourRate = MND(dr("hourrate"))
                .weeklyRate = MND(dr("Weekly"))
                .monthlyRate = MND(dr("Monthly"))
                .weekendRate = MND(dr("weekend"))
                .depositRate = MND(dr("deposit"))
                .ShowDialog()
            End With
            RefreshGridFromLastSQL()

        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Add multiple items to collection.
    ''' </summary>
    ''' <param name = "al"></param>
    Private Sub AddMultipleItemsToCollection(ByRef al As ArrayList)
        Dim i As Integer
        Dim oFrm As frmAddToList
        Dim j As Integer
        Try
            ' pass the arraylist of selected grid items to the addtolist form

            oFrm = New frmAddToList()
            With oFrm
                i = al(0) ' index of first selected grid row
                .ItemID = Me.dbgEquipment.DataSource.rows(i).item("equip_id")
                .ItemName = Me.dbgEquipment.DataSource.rows(i).item("equip_name")
                .PriceID = Me.dbgEquipment.DataSource.rows(i).item("price_id")
                .MeterRequired = dbgEquipment.DataSource.rows(i).item("meter_required")
                .MeterHours = 0
                oFrm.NbrRowsSelected = al.Count
                oFrm.alMultiItems = al
                oFrm.ShowDialog()
                DoEvents()
            End With
            'DoEvents()

            RefreshGridFromLastSQL()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Add multiple items to collection.
    ''' </summary>
    ''' <param name = "al"></param>
    Private Sub oldAddMultipleItemsToCollection(ByRef al As ArrayList)
        Dim i As Integer
        Dim oFrm As frmAddToList
        Dim j As Integer
        Try
            For j = 0 To al.Count - 1
                i = al(j)
                oFrm = New frmAddToList()
                DoEvents()
                With oFrm
                    .ItemID = Me.dbgEquipment.DataSource.rows(i).item("equip_id")
                    .ItemName = Me.dbgEquipment.DataSource.rows(i).item("equip_name")
                    .PriceID = Me.dbgEquipment.DataSource.rows(i).item("price_id")
                    .MeterRequired = dbgEquipment.DataSource.rows(i).item("meter_required")
                    .MeterHours = 0
                    .ShowDialog()
                    DoEvents()
                End With
            Next
            RefreshGridFromLastSQL()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub



    Private Sub RefreshGridFromLastSQL()


        Try
            If dtEquip.Rows.Count = 0 Then Exit Sub

            If Me.cboEquipType.Text = "All Equipment" Then
                Call LoadEquipGridFromType(0)
            Else
                Call LoadEquipGridFromType(Val(GetToken((Me.cboEquipType.Text), "")))
            End If
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


#End Region

#Region " Public Methods "
    Public Sub frmRental_Closing(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Dim Cancel As Short = eventArgs.Cancel
        Dim oRES As CTransaction

        oRES = New CTransaction()
        Call oRES.RemoveTempReservation("")
        oRES = Nothing

        If Me.WindowState <> System.Windows.Forms.FormWindowState.Minimized And Me.WindowState <> System.Windows.Forms.FormWindowState.Maximized Then
            SaveSetting(System.Reflection.Assembly.GetExecutingAssembly.GetName.Name, SETTINGS, "MainLeft", CStr(VB6.PixelsToTwipsX(Me.Left)))
            SaveSetting(System.Reflection.Assembly.GetExecutingAssembly.GetName.Name, SETTINGS, "MainTop", CStr(VB6.PixelsToTwipsY(Me.Top)))
            SaveSetting(System.Reflection.Assembly.GetExecutingAssembly.GetName.Name, SETTINGS, "MainWidth", CStr(VB6.PixelsToTwipsX(Me.Width)))
            SaveSetting(System.Reflection.Assembly.GetExecutingAssembly.GetName.Name, SETTINGS, "MainHeight", CStr(VB6.PixelsToTwipsY(Me.Height)))
            'SaveSetting App.Title, SETTINGS, "ShowRental", Me.mnuShowRentalFormOnLoad.Checked
            'SaveSetting App.Title, SETTINGS, "MaxOpen", Me.mnuMaximizeonOpen.Checked
            'SaveSetting App.Title, SETTINGS, "HourlyRates", Me.mnuHourlyRates.Checked
        End If

        eventArgs.Cancel = Cancel
    End Sub

    Public Sub cmdCancelOrder_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancelOrder.Click
        Dim oRES As CTransaction

        Try
            oRES = New CTransaction()
            If MsgBox("Are you sure you want to clear the shopping list?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                oRES = New CTransaction()
                Call oRES.RemoveTempReservation("")
                Me.LoadEquipGridFromType(0)
            End If
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Public Sub cboEquipType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboEquipType.SelectedIndexChanged

        Try
            If Noise Then Exit Sub
            If Me.cboEquipType.Text = "All Equipment" Then
                Call LoadEquipGridFromType(0)
            Else
                Call LoadEquipGridFromType(Val(GetToken((Me.cboEquipType.Text), "")))
            End If
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Public Sub LoadEquipGridFromType(ByRef riType As Short)
        Dim bWhere As Boolean


        Try
            SQL = ""
            SQL &= "SELECT Equipment.Equip_ID, "
            SQL &= "Equipment.Equip_Name, Equipment.Available, iif(isnull(Equipment.Damage),'',equipment.damage) as Damaged, "
            SQL &= "rental_rates.hourrate, "
            SQL &= "Rental_Rates.HalfDay, "
            SQL &= "Rental_Rates.Daily, Rental_Rates.Weekly, "
            SQL &= "Rental_Rates.Monthly, rental_rates.weekend, Rental_Rates.Minimum, "
            SQL &= "Rental_Rates.Deposit, equipment.Price_ID, purchase_date, "
            SQL &= "iif(meter_required=True, 'Yes', 'No') as Meter, Damage_Desc "
            SQL &= "FROM Rental_Rates RIGHT JOIN Equipment ON "
            SQL &= "Rental_Rates.Price_ID = Equipment.Price_Id "

            If riType > 0 Then
                SQL = SQL & "where equip_type_id = " & riType & " "
                bWhere = True
            End If

            If Me.optShowAvailable.Checked Then
                If bWhere Then
                    SQL = SQL & "and ucase(available) = 'YES' and (isnull(damage) or damage<>'H') "
                Else
                    SQL = SQL & "where ucase(available) = 'YES' and (isnull(damage) or damage<>'H') "
                End If
            End If

            SQL &= "order by Equipment.equip_name,reserving_company_id"
            LastSearchType = "Type"
            dtEquip = New DataTable("dt")
            Try
                Me.dbgEquipment.SetDataBinding(dtEquip, "")
            Catch ex As System.Exception
                StructuredErrorHandler(ex, False)
                MsgBox("Error encountered, please refresh grid (press F2); call developer if problem persists.", MsgBoxStyle.Information)
            End Try

            Dim Formats() As String = {"", "60", "T", "L",
                                       "", "150", "T", "L",
                                       "", "50", "T", "L",
                                       "", "50", "F", "L",
                                       "$#,##0.00", "60", "T", "R",
                                       "$#,##0.00", "60", "T", "R",
                                       "$#,##0.00", "60", "T", "R",
                                       "$#,##0.00", "60", "T", "R",
                                       "$#,##0.00", "60", "T", "R",
                                       "$#,##0.00", "60", "T", "R",
                                       "", "60", "T", "R",
                                       "$#,##0.00", "60", "T", "R",
                                       "0", "60", "T", "R",
                                       "", "60", "T", "R",
                                       "MM/dd/yyyy", "100", "F", "L",
                                       "", "70", "F", "L",
                                       "", "80", "F", "L"}

            If oDA.SendQuery(SQL, dtEquip, ConnectString, "dt") > 0 Then
                oCG.SetTablesStyle("RentMe", dtEquip, Me.dbgEquipment, Formats)
                oCG.BindDataTableToGrid(dtEquip, Me.dbgEquipment)
                oCG.DisableAddNew(Me.dbgEquipment, Me)
            End If

        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


#End Region

#Region " Form & Control Events "
    Private Sub frmRental_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated

        Try
            If formLoading Then
                formLoading = False
                Text = "RentalPro - Select Rental Equipment  (Version " &
                     System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileMajorPart &
                     "." & System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileMinorPart & ")"
                If Me.mnuAutoRunReminder.Checked Then
                    Dim o As New frmReminder()
                    o.ShowDialog()
                End If
            End If
            If modMain.ReloadRentalGrid Then
                Me.LoadEquipGridFromType(0)
                modMain.ReloadRentalGrid = False
            End If
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub frmRental_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Enter

        Try
            Me.Button1.Focus()
            If modMain.ReloadRentalGrid Then
                Me.LoadEquipGridFromType(0)
                modMain.ReloadRentalGrid = False
            End If
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub cmdCheckIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCheckIn.Click
        Dim oFrm As New frmSelectCheckInInvoice()

        Try
            Dim inv As Integer = oFrm.Display()
            If inv > 0 Then
                '#If Reliable Then
                Dim oFC As New frmCheckinNew()
                '#Else
                '         Dim oFC As New frmCheckIn()
                '#End If
                oFC.CurrentInvoice = inv
                oFC.ShowDialog()
            End If
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuFileExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileExit.Click
        Dim oRes As CTransaction


        Try
            oRes = New CTransaction()
            Call oRes.RemoveTempReservation("")
            Me.Close()
            DoEvents()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuHelpAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuHelpAbout.Click
        Dim oFRM As New frmAbout()

        Try
            oFRM.ShowDialog()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Dim oFrm As New frmSupplies()
            oFrm.ShowDialog()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click
        Dim oFrm As New frmSaleItemsMaintenance()

        Try
            oFrm.ShowDialog()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Dim oRES As CTransaction


        Try
            oRES = New CTransaction()
            Call oRES.RemoveTempReservation("")
            oRES = Nothing

            If Me.WindowState <> System.Windows.Forms.FormWindowState.Minimized And Me.WindowState <> System.Windows.Forms.FormWindowState.Maximized Then
                SaveSetting(System.Reflection.Assembly.GetExecutingAssembly.GetName.Name, SETTINGS, "MainLeft", CStr(VB6.PixelsToTwipsX(Me.Left)))
                SaveSetting(System.Reflection.Assembly.GetExecutingAssembly.GetName.Name, SETTINGS, "MainTop", CStr(VB6.PixelsToTwipsY(Me.Top)))
                SaveSetting(System.Reflection.Assembly.GetExecutingAssembly.GetName.Name, SETTINGS, "MainWidth", CStr(VB6.PixelsToTwipsX(Me.Width)))
                SaveSetting(System.Reflection.Assembly.GetExecutingAssembly.GetName.Name, SETTINGS, "MainHeight", CStr(VB6.PixelsToTwipsY(Me.Height)))
                'SaveSetting App.Title, SETTINGS, "ShowRental", Me.mnuShowRentalFormOnLoad.Checked
                'SaveSetting App.Title, SETTINGS, "MaxOpen", Me.mnuMaximizeonOpen.Checked
                'SaveSetting App.Title, SETTINGS, "HourlyRates", Me.mnuHourlyRates.Checked
            End If

            Me.Close()
            DoEvents()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub



    Private Sub mnuRepairCompact_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRepairCompact.Click
        Dim db As DAO.DBEngine
        Dim sUFN As String
        Dim sFN As String = System.IO.Path.GetFileNameWithoutExtension(DatabaseName)

        Try
            sUFN = AppPath & "\" & sFN & Format(Now, "MMddyyyyHHmmss") & ".mdb"
            Rename(DatabaseName, sUFN)
            db = New DAO.DBEngine()
            db.CompactDatabase(sUFN, AppPath & "\" & sFN)

        Catch ex As System.Exception
            Dim sb As New System.Text.StringBuilder()
            sb.Append("RentalPro is not able to backup the database at this" & Chr(10))
            sb.Append("time.  Another computer has the database locked." & Chr(10))
            sb.Append("Please ensure that all other copies of RentalPro" & Chr(10))
            sb.Append("are shut down.  " & Chr(10))
            sb.Append("" & Chr(10))
            sb.Append("If this error persists, you should try to backup the" & Chr(10))
            sb.Append("database after all machines are rebooted, possibly" & Chr(10))
            sb.Append("the next business day." & Chr(10))
            sb.Append("" & Chr(10))
            sb.Append("An alternate method is to bring up the Task Manager" & Chr(10))
            sb.Append("on all machines and terminate any copy of Rental.Exe" & Chr(10))
            sb.Append("that are found." & Chr(10))
            sb.Append("" & Chr(10))
            MsgBox(sb.ToString(), CType(48, Microsoft.VisualBasic.MsgBoxStyle), "Compact & Repair Error")

            StructuredErrorHandler(ex, False)
        End Try
    End Sub

    Private Sub mnuReceivePayments_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuReceivePayments.Click
        Dim oAC As New frmApplyCash()

        Try
            oAC.ShowDialog()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuSelectItems_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSelectItems.Click

        Try
            Dim sTxt As String = ""
            sTxt &= "To select a single piece of equipment to rent, do the "
            sTxt &= "following steps:" & vbCrLf & vbCrLf
            sTxt &= "1. Scroll to the single item that you want to rent and "
            sTxt &= "click in the check box for that item. " & Chr(13) & Chr(10)
            sTxt &= "2. When the Add To List form displays, select the period "
            sTxt &= "that you want to rent for.  If you want to rent for more "
            sTxt &= "than one of the selected periods, 2 days for example, "
            sTxt &= "select 2 from the drop down number list." & vbCrLf & vbCrLf
            sTxt &= "In special cases, like 'Braces & Bucks', you can select "
            sTxt &= "multiple items of this type at one time.  You can only do "
            sTxt &= "this for items whose equipment identifier is like "
            sTxt &= "'nn-nnn-01', 'nn-nnn-02' (where all of the digits of the ID "
            sTxt &= "are the same except the last 2 digits).  To select more than "
            sTxt &= "one of these special cases, do the following steps:" & Chr(13) & Chr(10)
            sTxt &= "1. Click in the left margin of the grid (list) with the "
            sTxt &= "Left Mouse Button on the first item and drag the mouse down "
            sTxt &= "to the last item you want to select. " & Chr(13) & Chr(10)
            sTxt &= "2. Release the Left mouse button." & Chr(13) & Chr(10)
            sTxt &= "3. Place the mouse cursor in the left margin of the grid "
            sTxt &= "and click the Right Mouse Button." & Chr(13) & Chr(10)
            sTxt &= "4. The Add To List form will be displayed just as if you "
            sTxt &= "had only selected one item.  It will display the first "
            sTxt &= "selected equipment item.  Follow the steps outlined above "
            sTxt &= "for one piece of equipment.  " & Chr(13) & Chr(10)
            sTxt &= "5. When you click the Add Button, the Add to List form will "
            sTxt &= "automatically add each of the remaining "
            sTxt &= "selected grid items to the customer's list.  So, if you select 6 items of this "
            sTxt &= "special type, the Add To List form will handle all of the items with one display of the form. "
            sTxt &= Chr(13) & Chr(10)

            Dim oHelp As New frmHelp()
            oHelp.CannedMessage = sTxt
            oHelp.ShowDialog()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuAddToCart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddToCart.Click

        Try
            Dim sTxt As String = ""
            sTxt &= "To add all of the selected items to the shopping cart, "
            sTxt &= "click the Add To Cart tool button." & Chr(13) & Chr(10)
            Dim oHelp As New frmHelp()
            oHelp.CannedMessage = sTxt
            oHelp.ShowDialog()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuCheckOut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCheckOut.Click

        Try
            Dim sTxt As String = ""
            sTxt &= "Once you have selected all of the items to be rented, plus "
            sTxt &= "any sale items (tools and supplies), and added them to the "
            sTxt &= "shopping cart, click the Check Out Tool button to print the "
            sTxt &= "invoice and place the equipment on rent." & Chr(13) & Chr(10)
            Dim oHelp As New frmHelp()
            oHelp.CannedMessage = sTxt
            oHelp.ShowDialog()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuCheckin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCheckin.Click

        Try
            Dim sTxt As String = ""
            sTxt &= "When a customer returns equipment, click the Check In Tool "
            sTxt &= "button." & Chr(13) & Chr(10)
            sTxt &= "1. When the customer/invoice selection dialog appears, you "
            sTxt &= "should select the respective customer.  Clicking on a "
            sTxt &= "customer will show the outstanding invoices for that "
            sTxt &= "customer." & Chr(13) & Chr(10)
            sTxt &= "2. Click on the desired invoice and click the Ok button.  "
            sTxt &= "The Check In Form will be displayed." & Chr(13) & Chr(10)
            Dim oHelp As New frmHelp()
            oHelp.CannedMessage = sTxt
            oHelp.ShowDialog()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuTools_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTools.Click

        Try
            Dim sTxt As String = ""
            sTxt &= "Clicking the Tools and Supplies tool button will cause the "
            sTxt &= "sale items dialog to be displayed.   From that form, you "
            sTxt &= "may select sale items, which are not listed in the "
            sTxt &= "equipment grid.  For help with that form, click the blue "
            sTxt &= "help link button on the form." & Chr(13) & Chr(10)
            Dim oHelp As New frmHelp()
            oHelp.CannedMessage = sTxt
            oHelp.ShowDialog()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuCancelOrder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCancelOrder.Click

        Try
            Dim sTxt As String = ""
            sTxt &= "At any time, you can cancel all of the items that the "
            sTxt &= "customer has added to their shopping cart by clicking the "
            sTxt &= "Cancel Order button." & Chr(13) & Chr(10)
            Dim oHelp As New frmHelp()
            oHelp.CannedMessage = sTxt
            oHelp.ShowDialog()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuCustMaint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCustMaint.Click
        Dim oFrm As New frmCustomerMaintenance()

        Try
            DoEvents()
            oFrm.ShowDialog()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuPreviewBeforePrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPreviewBeforePrint.Click

        Try
            Me.mnuPreviewBeforePrint.Checked = Not Me.mnuPreviewBeforePrint.Checked
            SaveSetting(System.Reflection.Assembly.GetExecutingAssembly.GetName.Name, SETTINGS, "Preview", Me.mnuPreviewBeforePrint.Checked)
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuShowOpenItems_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuShowOpenItems.Click
        Dim oFrm As New frmViewCustomerAccount()

        Try
            oFrm.ShowAll = False
            oFrm.ShowDialog()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuShowAllInvoices_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuShowAllInvoices.Click
        Dim oFrm As New frmViewCustomerAccount()

        Try
            oFrm.ShowAll = True
            oFrm.ShowDialog()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub



    Private Sub mnuPrintStatements_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPrintStatements.Click
        Dim oCS As New CPrintStatements()

        Try
            oCS.Preview = False
            oCS.PrintStatements()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuPreviewStatements_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPreviewStatements.Click
        Dim oCS As New CPrintStatements()

        Try
            oCS.Preview = True
            oCS.PrintStatements()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuPrintInvoices_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPrintInvoices.Click
        Dim oFrm As New frmSelectInvoicesToPrint()

        Try
            oFrm.ShowDialog()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuPreviewARReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPreviewARReport.Click
        Dim oAR As New CAcctsRecReport()

        Try
            oAR.Preview = True
            oAR.PrintARReport()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuPrintARReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPrintARReport.Click
        Dim oAR As New CAcctsRecReport()

        Try
            oAR.Preview = False
            oAR.PrintARReport()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuCreditDebitMemo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCreditDebitMemo.Click
        'MsgBox("Unimplemented")
        'Exit Sub
        Dim oFrm As New frmSelectCheckInInvoice()

        Try
            Dim inv As Integer = oFrm.Display(ShowAllInvoices:=True)
            If inv > 0 Then
                Dim oCD As New frmMemo()
                oCD.Text = "Select Invoice to Credit or Debit"
                oCD.CurrentInvoice = inv
                oCD.ShowDialog()
            End If
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub MenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem5.Click
        Dim oFrm As New frmViewCustomer()
        'If InputBox("Please enter the Password for use of this restricted feature.", "Restricted Feature", "") <> "pray" Then
        '   MsgBox("You are not authorized to use this feature.", MsgBoxStyle.Exclamation)
        '   Exit Sub
        'End If
        'Dim ofrm As New frmCustInvMaint()

        Try
            oFrm.ShowDialog()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuLaborMaint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuLaborMaint.Click
        Dim oFrm As New frmLaborItemsMaint()

        Try
            oFrm.ShowDialog()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuSalesTaxReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSalesTaxReport.Click
        Dim ofrm As New frmSalesTaxReport()

        Try
            ofrm.AcctBasis = AccountingBasis
            ofrm.ShowDialog()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub




    Private Sub mnuEquipUsageReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEquipUsageReport.Click
        Dim oFrm As New frmPrintEquipUsage()

        Try
            oFrm.ShowDialog()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub




    ''' <summary>
    ''' Load the reservations form.
    ''' </summary>
    ''' <param name = "sender"></param>
    ''' <param name = "e"></param>
    Private Sub cmdReserve_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReserve.Click
        Dim oFrm As New frmReservations()

        Try
            oFrm.ShowDialog()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuSetupConfig_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSetupConfig.Click
        Dim oFrm As New frmSetup()

        Try
            oFrm.ShowDialog()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuUpdateEquipCost_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuUpdateEquipCost.Click
        Dim oFrm As New frmEquipCost()

        Try
            oFrm.ShowDialog()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuPrintAged_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPrintAged.Click
        Dim o As New CAcctsRecReport()

        Try
            o.Preview = False
            o.PrintAgedARReport()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuPreviewAged_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPreviewAged.Click
        Dim o As New CAcctsRecReport()

        Try
            o.Preview = True
            o.PrintAgedARReport()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuMeterReadingMaint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMeterReadingMaint.Click
        Dim f As New frmMeterReadingMaint()

        Try
            f.ShowDialog()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

    Private Sub Unimplemented()
        MsgBox("Implemented according to customer specifications after Contract Signing.", MsgBoxStyle.Information)
    End Sub
    Private Sub mnuExportCustomer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExportCustomer.Click
        Unimplemented()
    End Sub

    Private Sub mnuExportInvoices_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExportInvoices.Click
        Unimplemented()
    End Sub

    Private Sub mnuExportInvoiceDetails_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExportInvoiceDetails.Click
        Unimplemented()
    End Sub

    Private Sub mnuExportEquipList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExportEquipList.Click
        Unimplemented()
    End Sub

    Private Sub mnuReRentReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuReRentReport.Click
        Dim oFrm As New frmReRentReport()

        Try
            oFrm.Show()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuTCOPreviewOnID_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTCOPreviewOnID.Click
        Dim o As New CTotalCostOwnershipReport(True, "ID")

        Try
            o.PrintTCOReport()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuTCOPrintonID_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTCOPrintonID.Click
        Dim o As New CTotalCostOwnershipReport(False, "ID")

        Try
            o.PrintTCOReport()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuTCOPreviewonName_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTCOPreviewonName.Click
        Dim o As New CTotalCostOwnershipReport(True, "NAME")

        Try
            o.PrintTCOReport()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuTCOPrinttonName_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTCOPrinttonName.Click
        Dim o As New CTotalCostOwnershipReport(False, "NAME")

        Try
            o.PrintTCOReport()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuSelectDatabase_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSelectDatabase.Click

        Try
            Dim s As String = SelectDatabase(Me.OpenFileDialog1, True)

            If s.Length = 0 Then
                Exit Sub
            End If
            DatabaseName = s
            Me.StatusBar1.Text = DatabaseName

            ConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabaseName
            SaveSetting(modMain.RENTALPRO, modMain.SETTINGS, "DBNAME", DatabaseName)

            Dim oCF As New CConfig()
            oCF.GetConfig()
            Me.LoadEquipGridFromType(0)
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuCheckDueItems_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCheckDueItems.Click
        Dim oFrm As New frmReminder()

        Try
            oFrm.Show()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub



    Private Sub mnuEqupReportPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEqupReportPreview.Click
        Dim o As New CEquipmentList()

        Try
            o.Preview = True
            o.PrintEquipmentList()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuEquipReportPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEquipReportPrint.Click
        Dim o As New CEquipmentList()

        Try
            o.Preview = False
            o.PrintEquipmentList()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuRefreshGrid_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRefreshGrid.Click

        Try
            Me.LoadEquipGridFromType(0)
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuAutoRunReminder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAutoRunReminder.Click

        Try
            With Me.mnuAutoRunReminder
                If .Checked Then
                    .Checked = False
                    SaveSetting(RENTALPRO, SETTINGS, "AUTOREMIND", False)
                Else
                    .Checked = True
                    SaveSetting(RENTALPRO, SETTINGS, "AUTOREMIND", True)
                End If
            End With
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub



    Private Sub btnReRent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReRent.Click
        Dim f As New frmRerent()

        Try
            f.ShowDialog()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub MenuItem15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem15.Click
        Dim f As New frmMeterReadingReport()

        Try
            f.ShowDialog()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub



    Private Sub mnuDeliveryReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDeliveryReport.Click
        Dim oFrm As New frmDeliveryReport()

        Try
            oFrm.AcctBasis = AccountingBasis
            oFrm.ShowDialog()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuLicensingAgreement_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuLicensingAgreement.Click
        Dim f As New frmHelp()

        Try
            f.Text = "Licensing Agreement"
            Dim sTxt As String = ""
            sTxt &= "TERMS OF USE OF THIS SOFTWARE:" & vbCrLf & vbCrLf
            sTxt &= "THIS SOFTWARE IS LICENSED TO Pioneer Rental  BY HHI "
            sTxt &= "SOFTWARE, FOR USE ON ONE TO FIVE COMPUTERS AT ONE RENTAL "
            sTxt &= "STORE ONLY.  " & vbCrLf & vbCrLf
            sTxt &= "THIS LICENSE CANNOT BE TRANSFERRED AND CANNOT BE USED AT "
            sTxt &= "ADDITIONAL STORES WITHOUT THE PURCHASE OF ADDITIONAL "
            sTxt &= "LICENSES.  THIS LICENSE CANNOT BE TRANSFERRED TO ANOTHER "
            sTxt &= "COMPANY IN THE CASE WHERE THE OWNERSHIP OF THE COMPANY "
            sTxt &= "CHANGES, WITHOUT THE EXPRESS PERMISSION OF HHI SOFTWARE." & Chr(13) & Chr(10)
            sTxt &= "THIS SOFTWARE IS PROVIDED BY HHI SOFTWARE, 'AS IS', AND ANY "
            sTxt &= "EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED "
            sTxt &= "TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS "
            sTxt &= "FOR A PARTICULAR PURPOSE ARE DISCLAIMED.  IN NO EVENT SHALL "
            sTxt &= "HHI SOFTWARE BE LIABLE FOR ANY DIRECT, INDIRECT, "
            sTxt &= "INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES "
            sTxt &= "(INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE "
            sTxt &= "GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR "
            sTxt &= "BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF "
            sTxt &= "LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT "
            sTxt &= "(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT "
            sTxt &= "OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF SUCH "
            sTxt &= "DAMAGE." & vbCrLf & vbCrLf
            f.CannedMessage = sTxt
            f.ShowDialog()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuProductInventoryReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuProductInventoryReport.Click
        Dim f As New frmSuppliesReport()

        Try
            f.ShowDialog()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuPreviewEqListRates_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPreviewEqListRates.Click
        Dim o As New CEquipAndRates()

        Try
            o.Preview = True
            o.PrintEquipmentList()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub mnuPrintEqListRates_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPrintEqListRates.Click
        Dim o As New CEquipAndRates()

        Try
            o.Preview = False
            o.PrintEquipmentList()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub



    Private Sub dbgEquipment_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dbgEquipment.MouseUp
        Dim bChecked As Boolean

        Try
            If Noise Then Exit Sub
            Noise = True

            ' see if multiple rows selected
            If e.Button = MouseButtons.Right Then
                With Me.dbgEquipment
                    Dim lastName As String = String.Empty
                    Dim i As Integer
                    Dim selCount As Integer
                    Dim al As New ArrayList()

                    For i = 0 To Me.dtEquip.Rows.Count - 1
                        If .IsSelected(i) Then
                            If lastName.Length = 0 Then
                                lastName = Me.dtEquip.Rows(i).Item("equip_name")
                            Else
                                If lastName <> Me.dtEquip.Rows(i).Item("equip_name") Or
                                   Me.dtEquip.Rows(i).Item("meter_required") Then
                                    MsgBox("If you select multiple items, they must all be the same equipment name, and they can't have meters.", MsgBoxStyle.Exclamation)
                                    Noise = False
                                    Exit Sub
                                End If
                            End If
                            al.Add(i)
                            selCount += 1
                        End If
                    Next

                    ' now, if more than one row selected
                    ' we need to tell the add list form to handle
                    ' all of them with one click
                    If selCount < 2 Then
                        GoTo [Continue]
                    End If

                    Me.AddMultipleItemsToCollection(al)
                    Noise = False
                    Exit Sub
                End With
            End If

[Continue]:
            If e.Button = MouseButtons.Right Then
                Static busy As Boolean
                If busy Then Exit Sub
                busy = True
                miHitRow = oCG.SelectCkBoxRow(Me.dbgEquipment, e)
                'dbgEquipment.Select(miHitRow)
                If e.Button = MouseButtons.Right Then
                    Dim oFrm As New frmMarkDamage()
                    oFrm.frm = Me
                    oFrm.ShowDialog()
                    Me.LoadEquipGridFromType(0)
                End If
                busy = False
            Else
                miHitRow = oCG.SelectCkBoxRow(dtEquip, Me.dbgEquipment, e, "RentMe", bChecked, 0)
                Dim oCO As New CCheckOut()
                If Not oCO.IsEquipAvailable(dtEquip.Rows(miHitRow).Item("equip_id"), True) Then
                    Me.LoadEquipGridFromType(0)
                    Noise = False
                    miHitRow = oCG.SelectCkBoxRow(dtEquip, Me.dbgEquipment, e, "RentMe", bChecked, 0)
                    Exit Sub
                End If
                If bChecked Then
                    Me.AddItemsToCollection(miHitRow)
                End If
            End If
            Noise = False
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
            Noise = False
        End Try
    End Sub


    Private Sub optShowAll_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optShowAll.CheckedChanged

        Try
            If eventSender.Checked Then
                RefreshGridFromLastSQL()
            End If
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub optShowAvailable_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optShowAvailable.CheckedChanged

        Try
            If eventSender.Checked Then
                RefreshGridFromLastSQL()
            End If
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub




    Private Sub txtSearchKey_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs)
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            KeyAscii = 0
        End If
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub


    Public Sub mnuCategoryMaintenance_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCategoryMaintenance.Popup

        Try
            mnuCategoryMaintenance_Click(eventSender, eventArgs)
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

    Public Sub mnuCategoryMaintenance_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCategoryMaintenance.Click
        Dim oFRM As New frmCategoryMaint()

        Try
            oFRM.ShowDialog()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Public Sub mnuEquipmentMaintenance_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuEquipmentMaintenance.Popup

        Try
            mnuEquipmentMaintenance_Click(eventSender, eventArgs)
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

    Public Sub mnuEquipmentMaintenance_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuEquipmentMaintenance.Click
        Dim oFrm As New frmEquipMaint()

        Try
            oFrm.ShowDialog()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Public Sub mnuRatesMaintenance_Popup(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuRatesMaintenance.Popup

        Try
            mnuRatesMaintenance_Click(eventSender, eventArgs)
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

    Public Sub mnuRatesMaintenance_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuRatesMaintenance.Click
        Dim oFrm As New frmRentalRates()

        Try
            oFrm.ShowDialog()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub



    Private Sub cmdRentThis_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdRentThis.Click
        Dim oFrm As New frmCustomers()

        Try
            oFrm.ShowDialog()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub cmdWhoHasIt_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdWhoHasIt.Click
        Dim oFrm As New frmWhoHasIt()

        Try
            oFrm.ShowDialog()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub



    Private Sub cboEquipType_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles cboEquipType.KeyPress

        Try
            Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
            ComboAutoSearch((Me.cboEquipType), KeyAscii)
            If KeyAscii = 0 Then
                eventArgs.Handled = True
            End If
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub





    Private Sub dbgEquipment_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dbgEquipment.KeyUp
        ' MsgBox("changed")
    End Sub

    Private Sub mnuFuelCostsMaintenance_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFuelCostsMaintenance.Click
        Dim frm As New frmFuelCostsMaintenance()
        frm.ShowDialog()
    End Sub

    Private Sub mnuConfigMaintenance_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuConfigMaintenance.Click
        Dim frm As New frmSetup()
        frm.ShowDialog()
    End Sub
    Private Sub mnuShowPrinters_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuShowPrinters.Click
        'Dim frm As New frmPrinters
        'frm.ShowDialog()
        PrintDialog1.Document = PrintDocument1
        PrintDialog1.PrinterSettings = PrintDocument1.PrinterSettings
        PrintDialog1.AllowSomePages = True
        If PrintDialog1.ShowDialog = DialogResult.OK Then
            'PrintDocument1.PrinterSettings = PrintDialog1.PrinterSettings
            'PrintDocument1.Print()
        End If
    End Sub

    Private Sub mnuEmployeeMaintenance_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEmployeeMaintenance.Click
        Dim oFrm As New frmEmployees()

        Try
            oFrm.ShowDialog()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub
    Private Sub mnuPrinters_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPrinters.Click
        Dim frm As New frmPrinters
        frm.ShowDialog()
    End Sub
    Private Sub mnuSendEmail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSendEmail.Click
        Dim oFrm As New frmEmail
        oFrm.ShowDialog()
    End Sub
#End Region

    Private Sub mnuShowRentedAndDue_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuShowRentedAndDue.Click
        Dim frm As New frmReminder
        frm.Show()
    End Sub

    Private Sub mnuShowDBLocation_Click(sender As Object, e As EventArgs) Handles mnuShowDBLocation.Click
        Dim dbLoc As String = GetSetting(RENTALPRO, SETTINGS, "DBNAME", "")
        If Not String.IsNullOrEmpty(dbLoc) Then
            MsgBox("RentalPro is expecting to use or using the database located at" + vbCrLf + dbLoc + ".", MsgBoxStyle.OkOnly, "Database Location")
        Else
            MsgBox("RentalPro does not know where the database should be located.  Please use the Select Database Menu Option to select a database.", MsgBoxStyle.OkOnly, "DB Location Unknown")
        End If
    End Sub
End Class