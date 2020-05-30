Option Strict Off
Option Explicit On
Friend Class frmEquipMaint
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
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents _lblLabels_0 As System.Windows.Forms.Label
    Public WithEvents _lblLabels_1 As System.Windows.Forms.Label
    Public WithEvents _lblLabels_2 As System.Windows.Forms.Label
    Public WithEvents _lblLabels_3 As System.Windows.Forms.Label
    Public WithEvents _lblLabels_4 As System.Windows.Forms.Label
    Public WithEvents _lblLabels_5 As System.Windows.Forms.Label
    Public WithEvents _lblLabels_6 As System.Windows.Forms.Label
    Public WithEvents _lblLabels_7 As System.Windows.Forms.Label
    Public WithEvents _lblLabels_8 As System.Windows.Forms.Label
    Public WithEvents _lblLabels_9 As System.Windows.Forms.Label
    Public WithEvents cboCategory As System.Windows.Forms.ComboBox
    Public WithEvents cmdAdd As System.Windows.Forms.Button
    Public WithEvents cmdClose As System.Windows.Forms.Button
    Public WithEvents cmdDelete As System.Windows.Forms.Button
    Public WithEvents cmdUpdate As System.Windows.Forms.Button
    Friend WithEvents dbcPriceID As System.Windows.Forms.ComboBox
    Friend WithEvents dbgEquipment As System.Windows.Forms.DataGrid
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents lblLabels As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents txtAvailable As System.Windows.Forms.ComboBox
    Public WithEvents txtAvailableDateTime As System.Windows.Forms.TextBox
    Public WithEvents txtEquip_Name As System.Windows.Forms.TextBox
    Public WithEvents txtEquipDesc As System.Windows.Forms.TextBox
    Public WithEvents txtEquipID As System.Windows.Forms.TextBox
    Public WithEvents txtModelNumber As System.Windows.Forms.TextBox
    Public WithEvents txtSerialNumber As System.Windows.Forms.TextBox
    Friend WithEvents chkMeterRequired As System.Windows.Forms.CheckBox
    Friend WithEvents txtMeterReading As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents chkSavePurchasePrice As System.Windows.Forms.CheckBox
    Friend WithEvents dtpPurchaseDate As System.Windows.Forms.DateTimePicker
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmEquipMaint))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdUpdate = New System.Windows.Forms.Button()
        Me.cmdClose = New System.Windows.Forms.Button()
        Me.cmdAdd = New System.Windows.Forms.Button()
        Me.cmdDelete = New System.Windows.Forms.Button()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.dtpPurchaseDate = New System.Windows.Forms.DateTimePicker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtMeterReading = New System.Windows.Forms.TextBox()
        Me.chkMeterRequired = New System.Windows.Forms.CheckBox()
        Me.dbcPriceID = New System.Windows.Forms.ComboBox()
        Me.txtAvailable = New System.Windows.Forms.ComboBox()
        Me.cboCategory = New System.Windows.Forms.ComboBox()
        Me.txtAvailableDateTime = New System.Windows.Forms.TextBox()
        Me.txtModelNumber = New System.Windows.Forms.TextBox()
        Me.txtSerialNumber = New System.Windows.Forms.TextBox()
        Me.txtEquipDesc = New System.Windows.Forms.TextBox()
        Me.txtEquipID = New System.Windows.Forms.TextBox()
        Me.txtEquip_Name = New System.Windows.Forms.TextBox()
        Me._lblLabels_0 = New System.Windows.Forms.Label()
        Me._lblLabels_1 = New System.Windows.Forms.Label()
        Me._lblLabels_2 = New System.Windows.Forms.Label()
        Me._lblLabels_3 = New System.Windows.Forms.Label()
        Me._lblLabels_4 = New System.Windows.Forms.Label()
        Me._lblLabels_5 = New System.Windows.Forms.Label()
        Me._lblLabels_6 = New System.Windows.Forms.Label()
        Me._lblLabels_7 = New System.Windows.Forms.Label()
        Me._lblLabels_8 = New System.Windows.Forms.Label()
        Me._lblLabels_9 = New System.Windows.Forms.Label()
        Me.lblLabels = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.dbgEquipment = New System.Windows.Forms.DataGrid()
        Me.chkSavePurchasePrice = New System.Windows.Forms.CheckBox()
        Me.Frame1.SuspendLayout()
        CType(Me.lblLabels, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dbgEquipment, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdUpdate
        '
        Me.cmdUpdate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdUpdate.BackColor = System.Drawing.SystemColors.Control
        Me.cmdUpdate.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdUpdate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdUpdate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdUpdate.Location = New System.Drawing.Point(710, 234)
        Me.cmdUpdate.Name = "cmdUpdate"
        Me.cmdUpdate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdUpdate.Size = New System.Drawing.Size(96, 26)
        Me.cmdUpdate.TabIndex = 10
        Me.cmdUpdate.Text = "&Save Updates"
        Me.ToolTip1.SetToolTip(Me.cmdUpdate, "Save changes or new record")
        Me.cmdUpdate.UseVisualStyleBackColor = False
        '
        'cmdClose
        '
        Me.cmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
        Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdClose.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdClose.Location = New System.Drawing.Point(710, 328)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClose.Size = New System.Drawing.Size(96, 26)
        Me.cmdClose.TabIndex = 13
        Me.cmdClose.Text = "&Close"
        Me.ToolTip1.SetToolTip(Me.cmdClose, "Close without saving current changes")
        Me.cmdClose.UseVisualStyleBackColor = False
        '
        'cmdAdd
        '
        Me.cmdAdd.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdAdd.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAdd.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAdd.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAdd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAdd.Location = New System.Drawing.Point(710, 264)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAdd.Size = New System.Drawing.Size(96, 26)
        Me.cmdAdd.TabIndex = 11
        Me.cmdAdd.Text = "&Clear To Add"
        Me.cmdAdd.UseVisualStyleBackColor = False
        '
        'cmdDelete
        '
        Me.cmdDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdDelete.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDelete.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdDelete.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDelete.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDelete.Location = New System.Drawing.Point(710, 296)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDelete.Size = New System.Drawing.Size(96, 26)
        Me.cmdDelete.TabIndex = 12
        Me.cmdDelete.Text = "&Delete Selected"
        Me.cmdDelete.UseVisualStyleBackColor = False
        '
        'Frame1
        '
        Me.Frame1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.chkSavePurchasePrice)
        Me.Frame1.Controls.Add(Me.dtpPurchaseDate)
        Me.Frame1.Controls.Add(Me.Label1)
        Me.Frame1.Controls.Add(Me.txtMeterReading)
        Me.Frame1.Controls.Add(Me.chkMeterRequired)
        Me.Frame1.Controls.Add(Me.dbcPriceID)
        Me.Frame1.Controls.Add(Me.txtAvailable)
        Me.Frame1.Controls.Add(Me.cboCategory)
        Me.Frame1.Controls.Add(Me.txtAvailableDateTime)
        Me.Frame1.Controls.Add(Me.txtModelNumber)
        Me.Frame1.Controls.Add(Me.txtSerialNumber)
        Me.Frame1.Controls.Add(Me.txtEquipDesc)
        Me.Frame1.Controls.Add(Me.txtEquipID)
        Me.Frame1.Controls.Add(Me.txtEquip_Name)
        Me.Frame1.Controls.Add(Me._lblLabels_0)
        Me.Frame1.Controls.Add(Me._lblLabels_1)
        Me.Frame1.Controls.Add(Me._lblLabels_2)
        Me.Frame1.Controls.Add(Me._lblLabels_3)
        Me.Frame1.Controls.Add(Me._lblLabels_4)
        Me.Frame1.Controls.Add(Me._lblLabels_5)
        Me.Frame1.Controls.Add(Me._lblLabels_6)
        Me.Frame1.Controls.Add(Me._lblLabels_7)
        Me.Frame1.Controls.Add(Me._lblLabels_8)
        Me.Frame1.Controls.Add(Me._lblLabels_9)
        Me.Frame1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(10, 226)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(694, 158)
        Me.Frame1.TabIndex = 14
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Edit Equipment"
        '
        'dtpPurchaseDate
        '
        Me.dtpPurchaseDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpPurchaseDate.Location = New System.Drawing.Point(451, 12)
        Me.dtpPurchaseDate.Name = "dtpPurchaseDate"
        Me.dtpPurchaseDate.Size = New System.Drawing.Size(93, 20)
        Me.dtpPurchaseDate.TabIndex = 30
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(360, 126)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(76, 14)
        Me.Label1.TabIndex = 29
        Me.Label1.Text = "Meter Reading"
        Me.Label1.Visible = False
        '
        'txtMeterReading
        '
        Me.txtMeterReading.Location = New System.Drawing.Point(448, 123)
        Me.txtMeterReading.Name = "txtMeterReading"
        Me.txtMeterReading.Size = New System.Drawing.Size(86, 20)
        Me.txtMeterReading.TabIndex = 28
        Me.txtMeterReading.Text = "0"
        Me.txtMeterReading.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtMeterReading.Visible = False
        '
        'chkMeterRequired
        '
        Me.chkMeterRequired.Location = New System.Drawing.Point(448, 106)
        Me.chkMeterRequired.Name = "chkMeterRequired"
        Me.chkMeterRequired.Size = New System.Drawing.Size(160, 16)
        Me.chkMeterRequired.TabIndex = 27
        Me.chkMeterRequired.Text = "Meter Required"
        '
        'dbcPriceID
        '
        Me.dbcPriceID.Location = New System.Drawing.Point(449, 82)
        Me.dbcPriceID.Name = "dbcPriceID"
        Me.dbcPriceID.Size = New System.Drawing.Size(179, 22)
        Me.dbcPriceID.TabIndex = 26
        '
        'txtAvailable
        '
        Me.txtAvailable.BackColor = System.Drawing.SystemColors.Window
        Me.txtAvailable.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtAvailable.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAvailable.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAvailable.Items.AddRange(New Object() {"YES", "NO", "RES"})
        Me.txtAvailable.Location = New System.Drawing.Point(450, 58)
        Me.txtAvailable.Name = "txtAvailable"
        Me.txtAvailable.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtAvailable.Size = New System.Drawing.Size(91, 22)
        Me.txtAvailable.TabIndex = 25
        '
        'cboCategory
        '
        Me.cboCategory.BackColor = System.Drawing.SystemColors.Window
        Me.cboCategory.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboCategory.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboCategory.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboCategory.Location = New System.Drawing.Point(92, 58)
        Me.cboCategory.Name = "cboCategory"
        Me.cboCategory.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboCategory.Size = New System.Drawing.Size(243, 22)
        Me.cboCategory.TabIndex = 3
        '
        'txtAvailableDateTime
        '
        Me.txtAvailableDateTime.AcceptsReturn = True
        Me.txtAvailableDateTime.BackColor = System.Drawing.SystemColors.Window
        Me.txtAvailableDateTime.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtAvailableDateTime.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAvailableDateTime.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtAvailableDateTime.Location = New System.Drawing.Point(450, 36)
        Me.txtAvailableDateTime.MaxLength = 0
        Me.txtAvailableDateTime.Name = "txtAvailableDateTime"
        Me.txtAvailableDateTime.ReadOnly = True
        Me.txtAvailableDateTime.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtAvailableDateTime.Size = New System.Drawing.Size(175, 19)
        Me.txtAvailableDateTime.TabIndex = 8
        Me.txtAvailableDateTime.Tag = "$#,##0.00;($#,##0.00)"
        Me.txtAvailableDateTime.Visible = False
        '
        'txtModelNumber
        '
        Me.txtModelNumber.AcceptsReturn = True
        Me.txtModelNumber.BackColor = System.Drawing.SystemColors.Window
        Me.txtModelNumber.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtModelNumber.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtModelNumber.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtModelNumber.Location = New System.Drawing.Point(92, 124)
        Me.txtModelNumber.MaxLength = 0
        Me.txtModelNumber.Name = "txtModelNumber"
        Me.txtModelNumber.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtModelNumber.Size = New System.Drawing.Size(175, 19)
        Me.txtModelNumber.TabIndex = 6
        Me.txtModelNumber.Tag = "$#,##0.00;($#,##0.00)"
        '
        'txtSerialNumber
        '
        Me.txtSerialNumber.AcceptsReturn = True
        Me.txtSerialNumber.BackColor = System.Drawing.SystemColors.Window
        Me.txtSerialNumber.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSerialNumber.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSerialNumber.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSerialNumber.Location = New System.Drawing.Point(92, 102)
        Me.txtSerialNumber.MaxLength = 0
        Me.txtSerialNumber.Name = "txtSerialNumber"
        Me.txtSerialNumber.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSerialNumber.Size = New System.Drawing.Size(243, 19)
        Me.txtSerialNumber.TabIndex = 5
        Me.txtSerialNumber.Tag = "$#,##0.00;($#,##0.00)"
        '
        'txtEquipDesc
        '
        Me.txtEquipDesc.AcceptsReturn = True
        Me.txtEquipDesc.BackColor = System.Drawing.SystemColors.Window
        Me.txtEquipDesc.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtEquipDesc.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEquipDesc.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtEquipDesc.Location = New System.Drawing.Point(92, 80)
        Me.txtEquipDesc.MaxLength = 0
        Me.txtEquipDesc.Name = "txtEquipDesc"
        Me.txtEquipDesc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtEquipDesc.Size = New System.Drawing.Size(243, 19)
        Me.txtEquipDesc.TabIndex = 4
        Me.txtEquipDesc.Tag = "$#,##0.00;($#,##0.00)"
        '
        'txtEquipID
        '
        Me.txtEquipID.AcceptsReturn = True
        Me.txtEquipID.BackColor = System.Drawing.Color.White
        Me.txtEquipID.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtEquipID.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEquipID.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtEquipID.Location = New System.Drawing.Point(92, 36)
        Me.txtEquipID.MaxLength = 0
        Me.txtEquipID.Name = "txtEquipID"
        Me.txtEquipID.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtEquipID.Size = New System.Drawing.Size(87, 19)
        Me.txtEquipID.TabIndex = 2
        '
        'txtEquip_Name
        '
        Me.txtEquip_Name.AcceptsReturn = True
        Me.txtEquip_Name.BackColor = System.Drawing.SystemColors.Window
        Me.txtEquip_Name.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtEquip_Name.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEquip_Name.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtEquip_Name.Location = New System.Drawing.Point(92, 14)
        Me.txtEquip_Name.MaxLength = 0
        Me.txtEquip_Name.Name = "txtEquip_Name"
        Me.txtEquip_Name.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtEquip_Name.Size = New System.Drawing.Size(245, 19)
        Me.txtEquip_Name.TabIndex = 1
        '
        '_lblLabels_0
        '
        Me._lblLabels_0.AutoSize = True
        Me._lblLabels_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblLabels_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_0, CType(0, Short))
        Me._lblLabels_0.Location = New System.Drawing.Point(12, 17)
        Me._lblLabels_0.Name = "_lblLabels_0"
        Me._lblLabels_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_0.Size = New System.Drawing.Size(69, 14)
        Me._lblLabels_0.TabIndex = 24
        Me._lblLabels_0.Text = "Equip_Name:"
        '
        '_lblLabels_1
        '
        Me._lblLabels_1.AutoSize = True
        Me._lblLabels_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblLabels_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_1, CType(1, Short))
        Me._lblLabels_1.Location = New System.Drawing.Point(12, 40)
        Me._lblLabels_1.Name = "_lblLabels_1"
        Me._lblLabels_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_1.Size = New System.Drawing.Size(67, 14)
        Me._lblLabels_1.TabIndex = 23
        Me._lblLabels_1.Text = "Company ID:"
        '
        '_lblLabels_2
        '
        Me._lblLabels_2.AutoSize = True
        Me._lblLabels_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblLabels_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_2, CType(2, Short))
        Me._lblLabels_2.Location = New System.Drawing.Point(12, 61)
        Me._lblLabels_2.Name = "_lblLabels_2"
        Me._lblLabels_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_2.Size = New System.Drawing.Size(83, 14)
        Me._lblLabels_2.TabIndex = 22
        Me._lblLabels_2.Text = "Equip Category:"
        '
        '_lblLabels_3
        '
        Me._lblLabels_3.AutoSize = True
        Me._lblLabels_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblLabels_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_3, CType(3, Short))
        Me._lblLabels_3.Location = New System.Drawing.Point(12, 84)
        Me._lblLabels_3.Name = "_lblLabels_3"
        Me._lblLabels_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_3.Size = New System.Drawing.Size(64, 14)
        Me._lblLabels_3.TabIndex = 21
        Me._lblLabels_3.Text = "Equip Desc:"
        '
        '_lblLabels_4
        '
        Me._lblLabels_4.AutoSize = True
        Me._lblLabels_4.BackColor = System.Drawing.SystemColors.Control
        Me._lblLabels_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_4, CType(4, Short))
        Me._lblLabels_4.Location = New System.Drawing.Point(12, 106)
        Me._lblLabels_4.Name = "_lblLabels_4"
        Me._lblLabels_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_4.Size = New System.Drawing.Size(77, 14)
        Me._lblLabels_4.TabIndex = 20
        Me._lblLabels_4.Text = "Serial Number:"
        '
        '_lblLabels_5
        '
        Me._lblLabels_5.AutoSize = True
        Me._lblLabels_5.BackColor = System.Drawing.SystemColors.Control
        Me._lblLabels_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_5, CType(5, Short))
        Me._lblLabels_5.Location = New System.Drawing.Point(28, 124)
        Me._lblLabels_5.Name = "_lblLabels_5"
        Me._lblLabels_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_5.Size = New System.Drawing.Size(58, 14)
        Me._lblLabels_5.TabIndex = 19
        Me._lblLabels_5.Text = "Model Nbr:"
        '
        '_lblLabels_6
        '
        Me._lblLabels_6.AutoSize = True
        Me._lblLabels_6.BackColor = System.Drawing.SystemColors.Control
        Me._lblLabels_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_6, CType(6, Short))
        Me._lblLabels_6.Location = New System.Drawing.Point(358, 17)
        Me._lblLabels_6.Name = "_lblLabels_6"
        Me._lblLabels_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_6.Size = New System.Drawing.Size(78, 14)
        Me._lblLabels_6.TabIndex = 18
        Me._lblLabels_6.Text = "Purchase Date"
        '
        '_lblLabels_7
        '
        Me._lblLabels_7.AutoSize = True
        Me._lblLabels_7.BackColor = System.Drawing.SystemColors.Control
        Me._lblLabels_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_7.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_7, CType(7, Short))
        Me._lblLabels_7.Location = New System.Drawing.Point(358, 38)
        Me._lblLabels_7.Name = "_lblLabels_7"
        Me._lblLabels_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_7.Size = New System.Drawing.Size(76, 14)
        Me._lblLabels_7.TabIndex = 17
        Me._lblLabels_7.Text = "Available Date"
        Me._lblLabels_7.Visible = False
        '
        '_lblLabels_8
        '
        Me._lblLabels_8.AutoSize = True
        Me._lblLabels_8.BackColor = System.Drawing.SystemColors.Control
        Me._lblLabels_8.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_8.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_8, CType(8, Short))
        Me._lblLabels_8.Location = New System.Drawing.Point(358, 60)
        Me._lblLabels_8.Name = "_lblLabels_8"
        Me._lblLabels_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_8.Size = New System.Drawing.Size(55, 14)
        Me._lblLabels_8.TabIndex = 16
        Me._lblLabels_8.Text = "Avaibable"
        '
        '_lblLabels_9
        '
        Me._lblLabels_9.AutoSize = True
        Me._lblLabels_9.BackColor = System.Drawing.SystemColors.Control
        Me._lblLabels_9.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_9.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_9, CType(9, Short))
        Me._lblLabels_9.Location = New System.Drawing.Point(358, 82)
        Me._lblLabels_9.Name = "_lblLabels_9"
        Me._lblLabels_9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_9.Size = New System.Drawing.Size(46, 14)
        Me._lblLabels_9.TabIndex = 15
        Me._lblLabels_9.Text = "Price ID:"
        '
        'dbgEquipment
        '
        Me.dbgEquipment.AllowSorting = False
        Me.dbgEquipment.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dbgEquipment.DataMember = ""
        Me.dbgEquipment.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dbgEquipment.Location = New System.Drawing.Point(0, 0)
        Me.dbgEquipment.Name = "dbgEquipment"
        Me.dbgEquipment.Size = New System.Drawing.Size(811, 224)
        Me.dbgEquipment.TabIndex = 15
        '
        'chkSavePurchasePrice
        '
        Me.chkSavePurchasePrice.AutoSize = True
        Me.chkSavePurchasePrice.Location = New System.Drawing.Point(551, 14)
        Me.chkSavePurchasePrice.Name = "chkSavePurchasePrice"
        Me.chkSavePurchasePrice.Size = New System.Drawing.Size(136, 18)
        Me.chkSavePurchasePrice.TabIndex = 31
        Me.chkSavePurchasePrice.Text = "Update Purchase Price"
        Me.chkSavePurchasePrice.UseVisualStyleBackColor = True
        '
        'frmEquipMaint
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(812, 390)
        Me.Controls.Add(Me.dbgEquipment)
        Me.Controls.Add(Me.cmdAdd)
        Me.Controls.Add(Me.cmdUpdate)
        Me.Controls.Add(Me.cmdDelete)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.Frame1)
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(9, 260)
        Me.MinimizeBox = False
        Me.Name = "frmEquipMaint"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Equipment Maintenance"
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        CType(Me.lblLabels, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dbgEquipment, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
#End Region
#Region " Module Variables "
    Dim msAddorEdit As String
    Dim mbDirty As Boolean
    Dim mbFormLoading As Boolean
    Dim oDA As CDataAccess
    Private dtEM As DataTable
    Private iHitRow As Integer
    Dim oCG As New CGrid()


#End Region

#Region " Public Methods "
    Public Function ConvertTextDateTimeToDouble(ByRef rsDateTime As Object) As Double
        Dim lsTemp As String
        Dim lsDate As String
        Dim lsTime As String

        lsTemp = rsDateTime
        lsDate = GetToken(lsTemp, "")
        lsTime = lsTemp


    End Function

#End Region

#Region " Private Methods "
    Private Sub LoadRates(ByVal i As Integer)
        Dim j As Integer
        Dim s As String

        Try
            With dtEM.Rows(i)
                Me.txtEquip_Name.Text = MNS(.Item("equip_name"))
                Me.txtEquipID.Text = MNS(.Item("equip_id"))
                Me.cboCategory.Text = Val(MNI(.Item("equip_type_id"))) ' - 1
                Me.txtEquipDesc.Text = MNS(.Item("equip_desc"))
                Me.txtSerialNumber.Text = MNS(.Item("serial_number"))
                Me.txtModelNumber.Text = MNS(.Item("model_number"))
                Me.dtpPurchaseDate.Value = IIf(IsDBNull(.Item("purchase_date")), Today, .Item("purchase_date"))
                Me.txtAvailableDateTime.Text = IIf(IsDBNull(.Item("available_date")), Now.ToString, .Item("available_date"))
                Me.txtAvailable.Text = MNS((.Item("available")))
                Dim desc As String = GetPriceDescription(MNI(.Item("price_id")))
                Me.dbcPriceID.Text = CType(MNI(.Item("price_id")), String) & " - " & desc
                'Me.txtMeterReading.Text = MNSng(.Item("hour_meter"))
                Me.chkMeterRequired.Checked = .Item("meter_required") = "Yes" 'CType(.Item("meter_required"), Boolean)
                If Not IsDBNull(.Item("purchase_date")) Then
                    Me.dtpPurchaseDate.Value = DateValue(.Item("purchase_date"))
                End If
            End With
            Me.mbDirty = False
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

    Private Sub ClearTextBoxes()
        With Me
            On Error Resume Next
            .txtEquip_Name.Text = String.Empty
            .txtEquipID.Text = String.Empty
            '.cboCategory.ListIndex = 0
            .txtEquipDesc.Text = String.Empty
            .txtSerialNumber.Text = String.Empty
            .txtModelNumber.Text = String.Empty
            .txtAvailableDateTime.Text = String.Empty
            .txtAvailable.Text = "Yes"
            Me.dbcPriceID.Text = String.Empty
            .txtMeterReading.Text = String.Empty
            .chkMeterRequired.Checked = False
        End With
    End Sub


    Private Function GetPriceDescription(ByVal PriceId As Integer) As String
        Dim sql As String
        Dim dt As New DataTable()

        sql = "select equip_name from rental_rates "
        sql &= "where price_id = " & PriceId
        If oDA.SendQuery(sql, dt, ConnectString) > 0 Then
            If IsDBNull(dt.Rows(0).Item(0)) Then
                Return ""
            Else
                Return CType(dt.Rows(0).Item(0), String)
            End If
        Else
            Return ""
        End If
    End Function

    Private Sub LoadTheGrid()
        Dim SQL As String
        Dim oDA As New CDataAccess()

        Try
            dtEM = New DataTable("dt")
            'SQL &= "iif(meter_required=True, 'Yes', 'No') as Meter,
            SQL = "select equip_name, "
            SQL &= "equip_id, equip_type_id, equip_desc,"
            SQL &= "serial_number, model_number, available_date, "
            SQL &= "available, price_id, discontinued,purchase_date, "
            SQL &= "iif(meter_required = True, 'Yes', 'No') as Meter_Required " ',hour_meter "
            SQL &= "from equipment "
            SQL &= " order by Equip_name"
            Dim Formats() As String =
               {"", "150", "T", "L",
               "", "60", "T", "L",
               "", "60", "T", "L",
               "", "100", "T", "L",
               "", "60", "T", "L",
               "", "60", "T", "L",
               "MM/dd/yyyy HH:mm", "100", "T", "L",
               "", "60", "T", "L",
               "", "60", "T", "R",
               "", "60", "T", "L",
               "MM/dd/yyyy HH:mm", "100", "T", "L",
               "", "60", "T", "L"}

            If oDA.SendQuery(SQL, dtEM, modMain.ConnectString, "dt") > 0 Then
                oCG.SetTablesStyle(dtEM, Me.dbgEquipment, Formats)
                Me.dbgEquipment.SetDataBinding(dtEM, "")
                oCG.DisableAddNew(dbgEquipment, Me)
            End If

        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

    Private Sub LoadPriceIDCombo()
        Dim SQL As String
        Dim dt As New DataTable()
        Dim i As Integer


        Try
            Me.dbcPriceID.Items.Clear()

            SQL = "Select Equip_name,price_id from rental_rates order by equip_name"
            If oDA.SendQuery(SQL, dt, ConnectString) > 0 Then
                For i = 0 To dt.Rows.Count - 1
                    With dt.Rows(i) 'Me.dbcPriceID
                        Me.dbcPriceID.Items.Add(CType(.Item("price_id"), String).PadLeft(3) & "-" & CType(.Item("equip_name"), String))
                    End With
                Next i
            End If
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

#End Region

#Region " Form & Control Events "
    Private Sub txtEquipID_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEquipID.Leave
        Me.txtEquipID.Text = UCase(Me.txtEquipID.Text)
    End Sub

    Private Sub txtSerialNumber_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSerialNumber.TextChanged
        mbDirty = True
    End Sub

    Private Sub txtAvailableDateTime_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAvailableDateTime.TextChanged
        mbDirty = True
    End Sub

    Private Sub txtAvailable_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAvailable.TextChanged
        mbDirty = True
    End Sub

    Private Sub txtEquip_Name_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEquip_Name.TextChanged
        mbDirty = True
    End Sub

    Private Sub txtEquipDesc_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEquipDesc.TextChanged
        mbDirty = True
    End Sub


    Private Sub txtAvailable_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtAvailable.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then KeyAscii = 0
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtAvailable_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtAvailable.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error Resume Next
        If KeyCode = 13 Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
        If KeyCode = System.Windows.Forms.Keys.Up Then System.Windows.Forms.SendKeys.SendWait("+{TAB}")
        If KeyCode = System.Windows.Forms.Keys.Down Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
    End Sub
    Private Sub txtAvailable_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAvailable.Enter
        txtAvailable.SelectionStart = 0
        txtAvailable.SelectionLength = Len(Trim(txtAvailable.Text))
    End Sub
    Private Sub txtAvailableDateTime_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtAvailableDateTime.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then KeyAscii = 0
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtAvailableDateTime_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtAvailableDateTime.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error Resume Next
        If KeyCode = 13 Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
        If KeyCode = System.Windows.Forms.Keys.Up Then System.Windows.Forms.SendKeys.SendWait("+{TAB}")
        If KeyCode = System.Windows.Forms.Keys.Down Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
    End Sub
    Private Sub txtAvailableDateTime_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAvailableDateTime.Enter
        txtAvailableDateTime.SelectionStart = 0
        txtAvailableDateTime.SelectionLength = Len(Trim(txtAvailableDateTime.Text))
    End Sub

    Private Sub txtDiscontinued_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        mbDirty = True
    End Sub


    Private Sub txtModelNumber_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtModelNumber.TextChanged
        mbDirty = True
    End Sub

    Private Sub txtModelNumber_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtModelNumber.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then KeyAscii = 0
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtModelNumber_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtModelNumber.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error Resume Next
        If KeyCode = 13 Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
        If KeyCode = System.Windows.Forms.Keys.Up Then System.Windows.Forms.SendKeys.SendWait("+{TAB}")
        If KeyCode = System.Windows.Forms.Keys.Down Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
    End Sub
    Private Sub txtModelNumber_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtModelNumber.Enter
        txtModelNumber.SelectionStart = 0
        txtModelNumber.SelectionLength = Len(Trim(txtModelNumber.Text))
    End Sub

    Private Sub txtSerialNumber_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSerialNumber.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then KeyAscii = 0
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtSerialNumber_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtSerialNumber.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error Resume Next
        If KeyCode = 13 Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
        If KeyCode = System.Windows.Forms.Keys.Up Then System.Windows.Forms.SendKeys.SendWait("+{TAB}")
        If KeyCode = System.Windows.Forms.Keys.Down Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
    End Sub
    Private Sub txtSerialNumber_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSerialNumber.Enter
        txtSerialNumber.SelectionStart = 0
        txtSerialNumber.SelectionLength = Len(Trim(txtSerialNumber.Text))
    End Sub

    Private Sub txtEquipDesc_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtEquipDesc.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then KeyAscii = 0
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtEquipDesc_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtEquipDesc.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error Resume Next
        If KeyCode = 13 Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
        If KeyCode = System.Windows.Forms.Keys.Up Then System.Windows.Forms.SendKeys.SendWait("+{TAB}")
        If KeyCode = System.Windows.Forms.Keys.Down Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
    End Sub
    Private Sub txtEquipDesc_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEquipDesc.Enter
        txtEquipDesc.SelectionStart = 0
        txtEquipDesc.SelectionLength = Len(Trim(txtEquipDesc.Text))
    End Sub



    Private Sub txtEquipID_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtEquipID.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then KeyAscii = 0
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtEquipID_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtEquipID.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error Resume Next
        If KeyCode = 13 Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
        If KeyCode = System.Windows.Forms.Keys.Up Then System.Windows.Forms.SendKeys.SendWait("+{TAB}")
        If KeyCode = System.Windows.Forms.Keys.Down Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
    End Sub
    Private Sub txtEquipID_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEquipID.Enter
        txtEquipID.SelectionStart = 0
        txtEquipID.SelectionLength = Len(Trim(txtEquipID.Text))
    End Sub


    Private Sub txtEquip_Name_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtEquip_Name.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then KeyAscii = 0
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtEquip_Name_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtEquip_Name.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error Resume Next
        If KeyCode = 13 Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
        If KeyCode = System.Windows.Forms.Keys.Up Then System.Windows.Forms.SendKeys.SendWait("+{TAB}")
        If KeyCode = System.Windows.Forms.Keys.Down Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
    End Sub
    Private Sub txtEquip_Name_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEquip_Name.Enter
        txtEquip_Name.SelectionStart = 0
        txtEquip_Name.SelectionLength = Len(Trim(txtEquip_Name.Text))
    End Sub

    Private Sub dbgEquipment_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dbgEquipment.MouseUp

        Try
            Dim pt = New Point(e.X, e.Y)
            Dim hti As DataGrid.HitTestInfo = dbgEquipment.HitTest(pt)
            dbgEquipment.Select(hti.Row)
            iHitRow = hti.Row
            Me.LoadRates(hti.Row)
            mbDirty = False
            msAddorEdit = "E"
        Catch ex As System.Exception
            'StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub txtMeterReading_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMeterReading.TextChanged
        mbDirty = True
    End Sub

    Private Sub chkMeterRequired_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkMeterRequired.CheckedChanged
        mbDirty = True
    End Sub


    Private Sub cboCategory_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboCategory.Enter
        Me.cboCategory.SelectionStart = 0
        Me.cboCategory.SelectionLength = Len(Trim(Me.cboCategory.Text))
    End Sub


    Private Sub cboCategory_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles cboCategory.KeyPress
        'Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        'ComboAutoSearch((Me.cboCategory), KeyAscii)
        'If KeyAscii = 0 Then
        '   eventArgs.Handled = True
        'End If
    End Sub


    Private Sub cmdAdd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAdd.Click
        Dim SQL As String
        Dim dt As New DataTable()


        Try
            If mbDirty Then
                If MsgBox("You have unsaved changes; do you want to add without saving your changes?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                    Exit Sub
                End If
            End If


            ClearTextBoxes()
            msAddorEdit = "A"

            Me.txtEquip_Name.Focus()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
        If mbDirty Then
            If MsgBox("You have unsaved changes; do you want to close without saving your changes?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Exit Sub
            End If
        End If

        Me.Close()
        System.Windows.Forms.Application.DoEvents()
    End Sub

    Private Sub cmdDelete_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDelete.Click
        Dim SQL As String
        Dim dt As New DataTable()
        Dim oDA As New CDataAccess()

        Try
            Dim sErr As String = ""
            If Me.txtAvailable.Text = "ON RENT" Then
                Dim sMsg As String
                sMsg = "The equipment selected for deletion is marked as " & Chr(10)
                sMsg &= "ON RENT.  That means that you must check in the " & Chr(10)
                sMsg &= "invoice that has this equipemnt on it before you" & Chr(10)
                sMsg &= "can delete it." & Chr(10)
                sMsg &= "" & Chr(10)
                sMsg &= "If you cannot find an invoice with this equipment rented, " & Chr(10)
                sMsg &= "you can change the Available Field to Available and then" & Chr(10)
                sMsg &= "delete the equipment.  You can use the WHO HAS IT button " & Chr(10)
                sMsg &= "to determine who has the equipment.  " & Chr(10)
                sMsg &= "" & Chr(10)
                sMsg &= "Rather than delete the equipment, mark it as not" & Chr(10)
                sMsg &= "available and then it will not appear in the grid as" & Chr(10)
                sMsg &= "rentable." & Chr(10)
                sMsg &= "" & Chr(10)
                MsgBox(sMsg, CType(48, Microsoft.VisualBasic.MsgBoxStyle), "Delete Denied")
                Exit Sub
            End If
            SQL = "select count(*) from equipment "
            SQL = SQL & "where price_id = " & Me.dtEM.Rows(Me.iHitRow).Item("price_id")
            Dim iRows = oDA.SendQuery(SQL, dt, ConnectString)

            If MsgBox("Are you sure you want to delete the selected row?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Exit Sub
            End If
            SQL = "delete from equipment "
            SQL &= "where equip_id = '" & Me.dtEM.Rows(Me.iHitRow).Item("equip_id") & "'"
            iRows = oDA.SendActionSql(SQL, ConnectString, sErr)
            If iRows = 0 Then
                MsgBox("Delete of equipment item failed.  " & Chr(10) & sErr, MsgBoxStyle.Critical)
                Exit Sub
            End If
            Me.LoadTheGrid()
            Me.LoadRates(0)
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub cmdUpdate_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdUpdate.Click
        Dim SQL As String
        Dim sErr As String
        Dim outil As New CUtilities()
        Dim dt As New DataTable()
        Dim sPID As String

        Try
            If Me.txtEquipID.Text.Trim.Length = 0 Or Me.txtEquip_Name.Text.Trim.Length = 0 Then
                MsgBox("Equipment ID and Name cannot be blank.", MsgBoxStyle.Information)
                Exit Sub
            End If
            Dim trash As String = MNS(Me.dbcPriceID.Text)
            If trash.Trim.Length = 0 Then
                MsgBox("You must select a Price ID.", MsgBoxStyle.Information)
                Exit Sub
            End If
            sPID = outil.GetToken(trash, "")
            If sPID.IndexOf("-") > -1 Then
                sPID = sPID.Substring(0, sPID.IndexOf("-")).Trim
            End If

            With Me
                If msAddorEdit = "A" Then
                    SQL = "select equip_id from equipment where equip_id = '" & Me.txtEquipID.Text & "' "
                    If oDA.SendQuery(SQL, dt, ConnectString) > 0 Then
                        MsgBox("Equip Id must be unique; the one you entered is already in use.", MsgBoxStyle.Exclamation)
                        Exit Sub
                    End If
                    SQL = "insert into equipment "
                    SQL &= "(equip_name, equip_id, equip_type_id, Equip_desc, serial_number, "
                    SQL &= "Model_number, purchase_date, available_date,price_id, available, "
                    SQL &= "meter_required ) "
                    SQL &= "values('"
                    SQL &= Replace(.txtEquip_Name.Text, "'", "''") & "', "
                    SQL &= "'" & .txtEquipID.Text & "', "
                    SQL &= GetToken((.cboCategory.Text), "") & ", "
                    SQL &= "'" & .txtEquipDesc.Text & "', "
                    SQL &= "'" & .txtSerialNumber.Text & "', "
                    SQL &= "'" & .txtModelNumber.Text & "', "
                    SQL &= "#" & .dtpPurchaseDate.Value & "#, "
                    SQL &= "#" & Now.ToString & "#, "
                    SQL &= sPID & ", "
                    SQL &= "'" & UCase(.txtAvailable.Text) & "', "
                    SQL &= IIf(Me.chkMeterRequired.Checked, True, False) & ") "
                    'SQL &= Val(Me.txtMeterReading.Text) & ") "
                Else
                    SQL = "update equipment "
                    SQL &= "set equip_name = '" & Replace(.txtEquip_Name.Text, "'", "''") & "', "
                    'SQL = SQL & "equip_id = " & .txtEquipID & ", "
                    SQL &= "equip_type_id = " & GetToken((.cboCategory.Text), "") & ", "
                    SQL &= "equip_desc = '" & .txtEquipDesc.Text & "', "
                    SQL &= "serial_number = '" & .txtSerialNumber.Text & "', "
                    SQL &= "model_number = '" & .txtModelNumber.Text & "', "
                    If chkSavePurchasePrice.Checked Then
                        SQL &= "purchase_date = #" & Me.dtpPurchaseDate.Value & "#, "
                    End If
                    SQL &= "available = '" & UCase(.txtAvailable.Text) & "', "
                    SQL &= "price_id = " & sPID & ", "
                    'SQL &= "hour_meter = " & Val(Me.txtMeterReading.Text) & ", "
                    SQL &= "meter_required = " & IIf(Me.chkMeterRequired.Checked, True, False) & " "
                    SQL &= "where equip_id = '" & .txtEquipID.Text & "'"
                    End If
            End With
            If oDA.SendActionSql(SQL, ConnectString, sErr) <= 0 Then
                MsgBox("Update of Equipment failed." & Chr(10) & sErr & Chr(10), MsgBoxStyle.Exclamation)
                Exit Sub
            End If
            msAddorEdit = "E"
            LoadTheGrid()
            mbDirty = False
            chkSavePurchasePrice.Checked = False
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub



    Private Sub frmEquipMaint_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'IncrementChildCount (Me)
        CenterForm(Me)
        LoadCategoryCombo()
        LoadTheGrid()
        LoadRates(0)
        msAddorEdit = "E"
        mbFormLoading = True
        LoadPriceIDCombo()
    End Sub
    Private Sub LoadCategoryCombo()
        Dim SQL As String
        Dim dt As New DataTable()
        Dim i As Integer


        Try
            SQL = "select * from Equipment_type order by equip_type_id"
            oDA.SendQuery(SQL, dt, ConnectString)

            Me.cboCategory.Items.Clear()

            For i = 0 To dt.Rows.Count - 1
                With dt.Rows(i)
                    Me.cboCategory.Items.Add(CType(.Item("equip_type_id"), String).PadLeft(2) & " - " & .Item("equip_type"))
                End With
            Next i

        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


#End Region

End Class