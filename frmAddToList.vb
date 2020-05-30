Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports System.Threading
Public Class frmAddToList
    Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
    Public Sub New()
        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()
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
    Public WithEvents _Label1_3 As System.Windows.Forms.Label
    Public WithEvents _Label1_4 As System.Windows.Forms.Label
    Public WithEvents _Label1_7 As System.Windows.Forms.Label
    Public WithEvents chkDeposit As System.Windows.Forms.CheckBox
    Public WithEvents fraItemDesc As System.Windows.Forms.GroupBox
    Public WithEvents fraRentalPeriod As System.Windows.Forms.GroupBox
    Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents Line1 As System.Windows.Forms.Label
    Public WithEvents optDaily As System.Windows.Forms.RadioButton
    Public WithEvents optHalfDay As System.Windows.Forms.RadioButton
    Public WithEvents optHour As System.Windows.Forms.RadioButton
    Public WithEvents optMonthly As System.Windows.Forms.RadioButton
    Public WithEvents optSaleItem As System.Windows.Forms.RadioButton
    Public WithEvents optWeekEnd As System.Windows.Forms.RadioButton
    Public WithEvents optWeekly As System.Windows.Forms.RadioButton
    Public WithEvents txtDaily As System.Windows.Forms.TextBox
    Public WithEvents txtDeposit As System.Windows.Forms.TextBox
    Public WithEvents txtHalfDay As System.Windows.Forms.TextBox
    Public WithEvents txtHourly As System.Windows.Forms.TextBox
    Public WithEvents txtItemDesc As System.Windows.Forms.TextBox
    Public WithEvents txtItemID As System.Windows.Forms.TextBox
    Public WithEvents txtItemName As System.Windows.Forms.TextBox
    Public WithEvents txtItemTotal As System.Windows.Forms.TextBox
    Public WithEvents txtMonthly As System.Windows.Forms.TextBox
    Public WithEvents txtPeriods As System.Windows.Forms.ComboBox
    Public WithEvents txtPrice As System.Windows.Forms.TextBox
    Public WithEvents txtSaleItem As System.Windows.Forms.TextBox
    Public WithEvents txtWeekEnd As System.Windows.Forms.TextBox
    Public WithEvents txtWeekly As System.Windows.Forms.TextBox
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Public WithEvents cmdReserve As System.Windows.Forms.Button
    Public WithEvents btnAddToList As System.Windows.Forms.Button
    Public WithEvents cmdCancelOrder As System.Windows.Forms.Button
    Friend WithEvents txtMultiCount As System.Windows.Forms.TextBox
    Friend WithEvents lblNbrItems As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAddToList))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.chkDeposit = New System.Windows.Forms.CheckBox()
        Me.cmdReserve = New System.Windows.Forms.Button()
        Me.btnAddToList = New System.Windows.Forms.Button()
        Me.cmdCancelOrder = New System.Windows.Forms.Button()
        Me.txtWeekEnd = New System.Windows.Forms.TextBox()
        Me.txtMonthly = New System.Windows.Forms.TextBox()
        Me.txtWeekly = New System.Windows.Forms.TextBox()
        Me.txtDaily = New System.Windows.Forms.TextBox()
        Me.txtHalfDay = New System.Windows.Forms.TextBox()
        Me.txtHourly = New System.Windows.Forms.TextBox()
        Me.fraItemDesc = New System.Windows.Forms.GroupBox()
        Me.lblNbrItems = New System.Windows.Forms.Label()
        Me.txtMultiCount = New System.Windows.Forms.TextBox()
        Me.txtItemDesc = New System.Windows.Forms.TextBox()
        Me.txtItemID = New System.Windows.Forms.TextBox()
        Me.txtItemName = New System.Windows.Forms.TextBox()
        Me._Label1_3 = New System.Windows.Forms.Label()
        Me._Label1_2 = New System.Windows.Forms.Label()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me.fraRentalPeriod = New System.Windows.Forms.GroupBox()
        Me.txtPeriods = New System.Windows.Forms.ComboBox()
        Me.optWeekEnd = New System.Windows.Forms.RadioButton()
        Me.txtDeposit = New System.Windows.Forms.TextBox()
        Me.txtPrice = New System.Windows.Forms.TextBox()
        Me.txtItemTotal = New System.Windows.Forms.TextBox()
        Me.txtSaleItem = New System.Windows.Forms.TextBox()
        Me.optSaleItem = New System.Windows.Forms.RadioButton()
        Me.optMonthly = New System.Windows.Forms.RadioButton()
        Me.optWeekly = New System.Windows.Forms.RadioButton()
        Me.optDaily = New System.Windows.Forms.RadioButton()
        Me.optHalfDay = New System.Windows.Forms.RadioButton()
        Me.optHour = New System.Windows.Forms.RadioButton()
        Me.Line1 = New System.Windows.Forms.Label()
        Me._Label1_7 = New System.Windows.Forms.Label()
        Me._Label1_4 = New System.Windows.Forms.Label()
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel()
        Me.fraItemDesc.SuspendLayout()
        Me.fraRentalPeriod.SuspendLayout()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'chkDeposit
        '
        Me.chkDeposit.BackColor = System.Drawing.SystemColors.Control
        Me.chkDeposit.Checked = True
        Me.chkDeposit.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkDeposit.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkDeposit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkDeposit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkDeposit.Location = New System.Drawing.Point(256, 112)
        Me.chkDeposit.Name = "chkDeposit"
        Me.chkDeposit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkDeposit.Size = New System.Drawing.Size(62, 17)
        Me.chkDeposit.TabIndex = 31
        Me.chkDeposit.Text = "De&posit"
        Me.ToolTip1.SetToolTip(Me.chkDeposit, "Click to remove this deposit")
        Me.chkDeposit.UseVisualStyleBackColor = False
        '
        'cmdReserve
        '
        Me.cmdReserve.BackColor = System.Drawing.SystemColors.Control
        Me.cmdReserve.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdReserve.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdReserve.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdReserve.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdReserve.Location = New System.Drawing.Point(8, 304)
        Me.cmdReserve.Name = "cmdReserve"
        Me.cmdReserve.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdReserve.Size = New System.Drawing.Size(76, 41)
        Me.cmdReserve.TabIndex = 24
        Me.cmdReserve.TabStop = False
        Me.cmdReserve.Text = "&Reserve Equipment"
        Me.ToolTip1.SetToolTip(Me.cmdReserve, "Look at future reservations and usage")
        Me.cmdReserve.UseVisualStyleBackColor = False
        '
        'btnAddToList
        '
        Me.btnAddToList.BackColor = System.Drawing.SystemColors.Control
        Me.btnAddToList.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnAddToList.Enabled = False
        Me.btnAddToList.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAddToList.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnAddToList.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnAddToList.Location = New System.Drawing.Point(262, 304)
        Me.btnAddToList.Name = "btnAddToList"
        Me.btnAddToList.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnAddToList.Size = New System.Drawing.Size(70, 41)
        Me.btnAddToList.TabIndex = 25
        Me.btnAddToList.TabStop = False
        Me.btnAddToList.Text = "&Add to List"
        Me.ToolTip1.SetToolTip(Me.btnAddToList, "Add item to list")
        Me.btnAddToList.UseVisualStyleBackColor = False
        '
        'cmdCancelOrder
        '
        Me.cmdCancelOrder.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancelOrder.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancelOrder.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancelOrder.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancelOrder.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdCancelOrder.Location = New System.Drawing.Point(341, 304)
        Me.cmdCancelOrder.Name = "cmdCancelOrder"
        Me.cmdCancelOrder.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancelOrder.Size = New System.Drawing.Size(70, 41)
        Me.cmdCancelOrder.TabIndex = 26
        Me.cmdCancelOrder.TabStop = False
        Me.cmdCancelOrder.Text = "&Cancel"
        Me.ToolTip1.SetToolTip(Me.cmdCancelOrder, "Cancel item")
        Me.cmdCancelOrder.UseVisualStyleBackColor = False
        '
        'txtWeekEnd
        '
        Me.txtWeekEnd.AcceptsReturn = True
        Me.txtWeekEnd.BackColor = System.Drawing.SystemColors.Window
        Me.txtWeekEnd.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtWeekEnd.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtWeekEnd.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtWeekEnd.Location = New System.Drawing.Point(280, 60)
        Me.txtWeekEnd.MaxLength = 0
        Me.txtWeekEnd.Name = "txtWeekEnd"
        Me.txtWeekEnd.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtWeekEnd.Size = New System.Drawing.Size(75, 19)
        Me.txtWeekEnd.TabIndex = 18
        Me.txtWeekEnd.TabStop = False
        Me.txtWeekEnd.Tag = "(No Auto Formatting)"
        Me.txtWeekEnd.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtWeekEnd, "Change this price manually to give a price break")
        '
        'txtMonthly
        '
        Me.txtMonthly.AcceptsReturn = True
        Me.txtMonthly.BackColor = System.Drawing.SystemColors.Window
        Me.txtMonthly.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMonthly.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMonthly.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMonthly.Location = New System.Drawing.Point(280, 38)
        Me.txtMonthly.MaxLength = 0
        Me.txtMonthly.Name = "txtMonthly"
        Me.txtMonthly.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMonthly.Size = New System.Drawing.Size(75, 19)
        Me.txtMonthly.TabIndex = 17
        Me.txtMonthly.TabStop = False
        Me.txtMonthly.Tag = "(No Auto Formatting)"
        Me.txtMonthly.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtMonthly, "Change this price manually to give a price break")
        '
        'txtWeekly
        '
        Me.txtWeekly.AcceptsReturn = True
        Me.txtWeekly.BackColor = System.Drawing.SystemColors.Window
        Me.txtWeekly.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtWeekly.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtWeekly.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtWeekly.Location = New System.Drawing.Point(280, 16)
        Me.txtWeekly.MaxLength = 0
        Me.txtWeekly.Name = "txtWeekly"
        Me.txtWeekly.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtWeekly.Size = New System.Drawing.Size(75, 19)
        Me.txtWeekly.TabIndex = 16
        Me.txtWeekly.TabStop = False
        Me.txtWeekly.Tag = "(No Auto Formatting)"
        Me.txtWeekly.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtWeekly, "Change this price manually to give a price break")
        '
        'txtDaily
        '
        Me.txtDaily.AcceptsReturn = True
        Me.txtDaily.BackColor = System.Drawing.SystemColors.Window
        Me.txtDaily.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDaily.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDaily.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDaily.Location = New System.Drawing.Point(92, 59)
        Me.txtDaily.MaxLength = 0
        Me.txtDaily.Name = "txtDaily"
        Me.txtDaily.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDaily.Size = New System.Drawing.Size(75, 19)
        Me.txtDaily.TabIndex = 15
        Me.txtDaily.TabStop = False
        Me.txtDaily.Tag = "(No Auto Formatting)"
        Me.txtDaily.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtDaily, "Change this price manually to give a price break")
        '
        'txtHalfDay
        '
        Me.txtHalfDay.AcceptsReturn = True
        Me.txtHalfDay.BackColor = System.Drawing.SystemColors.Window
        Me.txtHalfDay.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtHalfDay.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtHalfDay.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtHalfDay.Location = New System.Drawing.Point(92, 37)
        Me.txtHalfDay.MaxLength = 0
        Me.txtHalfDay.Name = "txtHalfDay"
        Me.txtHalfDay.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtHalfDay.Size = New System.Drawing.Size(75, 19)
        Me.txtHalfDay.TabIndex = 14
        Me.txtHalfDay.TabStop = False
        Me.txtHalfDay.Tag = "(No Auto Formatting)"
        Me.txtHalfDay.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtHalfDay, "Change this price manually to give a price break")
        '
        'txtHourly
        '
        Me.txtHourly.AcceptsReturn = True
        Me.txtHourly.BackColor = System.Drawing.SystemColors.Window
        Me.txtHourly.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtHourly.Enabled = False
        Me.txtHourly.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtHourly.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtHourly.Location = New System.Drawing.Point(92, 16)
        Me.txtHourly.MaxLength = 0
        Me.txtHourly.Name = "txtHourly"
        Me.txtHourly.ReadOnly = True
        Me.txtHourly.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtHourly.Size = New System.Drawing.Size(75, 19)
        Me.txtHourly.TabIndex = 13
        Me.txtHourly.TabStop = False
        Me.txtHourly.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtHourly, "Change this price manually to give a price break")
        '
        'fraItemDesc
        '
        Me.fraItemDesc.BackColor = System.Drawing.SystemColors.Control
        Me.fraItemDesc.Controls.Add(Me.lblNbrItems)
        Me.fraItemDesc.Controls.Add(Me.txtMultiCount)
        Me.fraItemDesc.Controls.Add(Me.txtItemDesc)
        Me.fraItemDesc.Controls.Add(Me.txtItemID)
        Me.fraItemDesc.Controls.Add(Me.txtItemName)
        Me.fraItemDesc.Controls.Add(Me._Label1_3)
        Me.fraItemDesc.Controls.Add(Me._Label1_2)
        Me.fraItemDesc.Controls.Add(Me._Label1_0)
        Me.fraItemDesc.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraItemDesc.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraItemDesc.Location = New System.Drawing.Point(6, 6)
        Me.fraItemDesc.Name = "fraItemDesc"
        Me.fraItemDesc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraItemDesc.Size = New System.Drawing.Size(417, 101)
        Me.fraItemDesc.TabIndex = 22
        Me.fraItemDesc.TabStop = False
        Me.fraItemDesc.Text = "Rental Item"
        '
        'lblNbrItems
        '
        Me.lblNbrItems.AutoSize = True
        Me.lblNbrItems.Location = New System.Drawing.Point(240, 16)
        Me.lblNbrItems.Name = "lblNbrItems"
        Me.lblNbrItems.Size = New System.Drawing.Size(72, 14)
        Me.lblNbrItems.TabIndex = 30
        Me.lblNbrItems.Text = "Number Items"
        Me.lblNbrItems.Visible = False
        '
        'txtMultiCount
        '
        Me.txtMultiCount.Location = New System.Drawing.Point(314, 14)
        Me.txtMultiCount.MaxLength = 3
        Me.txtMultiCount.Name = "txtMultiCount"
        Me.txtMultiCount.ReadOnly = True
        Me.txtMultiCount.Size = New System.Drawing.Size(32, 20)
        Me.txtMultiCount.TabIndex = 29
        Me.txtMultiCount.Visible = False
        '
        'txtItemDesc
        '
        Me.txtItemDesc.AcceptsReturn = True
        Me.txtItemDesc.BackColor = System.Drawing.SystemColors.Window
        Me.txtItemDesc.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtItemDesc.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtItemDesc.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtItemDesc.Location = New System.Drawing.Point(48, 66)
        Me.txtItemDesc.MaxLength = 0
        Me.txtItemDesc.Name = "txtItemDesc"
        Me.txtItemDesc.ReadOnly = True
        Me.txtItemDesc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtItemDesc.Size = New System.Drawing.Size(357, 19)
        Me.txtItemDesc.TabIndex = 28
        Me.txtItemDesc.TabStop = False
        '
        'txtItemID
        '
        Me.txtItemID.AcceptsReturn = True
        Me.txtItemID.BackColor = System.Drawing.SystemColors.Window
        Me.txtItemID.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtItemID.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtItemID.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtItemID.Location = New System.Drawing.Point(48, 16)
        Me.txtItemID.MaxLength = 0
        Me.txtItemID.Name = "txtItemID"
        Me.txtItemID.ReadOnly = True
        Me.txtItemID.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtItemID.Size = New System.Drawing.Size(75, 19)
        Me.txtItemID.TabIndex = 26
        Me.txtItemID.TabStop = False
        '
        'txtItemName
        '
        Me.txtItemName.AcceptsReturn = True
        Me.txtItemName.BackColor = System.Drawing.SystemColors.Window
        Me.txtItemName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtItemName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtItemName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtItemName.Location = New System.Drawing.Point(48, 40)
        Me.txtItemName.MaxLength = 0
        Me.txtItemName.Name = "txtItemName"
        Me.txtItemName.ReadOnly = True
        Me.txtItemName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtItemName.Size = New System.Drawing.Size(229, 19)
        Me.txtItemName.TabIndex = 24
        Me.txtItemName.TabStop = False
        '
        '_Label1_3
        '
        Me._Label1_3.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_3, CType(3, Short))
        Me._Label1_3.Location = New System.Drawing.Point(10, 62)
        Me._Label1_3.Name = "_Label1_3"
        Me._Label1_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_3.Size = New System.Drawing.Size(47, 27)
        Me._Label1_3.TabIndex = 27
        Me._Label1_3.Text = "Item Desc"
        '
        '_Label1_2
        '
        Me._Label1_2.AutoSize = True
        Me._Label1_2.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_2, CType(2, Short))
        Me._Label1_2.Location = New System.Drawing.Point(10, 18)
        Me._Label1_2.Name = "_Label1_2"
        Me._Label1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_2.Size = New System.Drawing.Size(38, 14)
        Me._Label1_2.TabIndex = 25
        Me._Label1_2.Text = "Item ID"
        '
        '_Label1_0
        '
        Me._Label1_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_0, CType(0, Short))
        Me._Label1_0.Location = New System.Drawing.Point(10, 34)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_0.Size = New System.Drawing.Size(47, 27)
        Me._Label1_0.TabIndex = 23
        Me._Label1_0.Text = "Item Name"
        '
        'fraRentalPeriod
        '
        Me.fraRentalPeriod.BackColor = System.Drawing.SystemColors.Control
        Me.fraRentalPeriod.Controls.Add(Me.txtPeriods)
        Me.fraRentalPeriod.Controls.Add(Me.txtWeekEnd)
        Me.fraRentalPeriod.Controls.Add(Me.optWeekEnd)
        Me.fraRentalPeriod.Controls.Add(Me.chkDeposit)
        Me.fraRentalPeriod.Controls.Add(Me.txtDeposit)
        Me.fraRentalPeriod.Controls.Add(Me.txtPrice)
        Me.fraRentalPeriod.Controls.Add(Me.txtItemTotal)
        Me.fraRentalPeriod.Controls.Add(Me.txtSaleItem)
        Me.fraRentalPeriod.Controls.Add(Me.txtMonthly)
        Me.fraRentalPeriod.Controls.Add(Me.txtWeekly)
        Me.fraRentalPeriod.Controls.Add(Me.txtDaily)
        Me.fraRentalPeriod.Controls.Add(Me.txtHalfDay)
        Me.fraRentalPeriod.Controls.Add(Me.txtHourly)
        Me.fraRentalPeriod.Controls.Add(Me.optSaleItem)
        Me.fraRentalPeriod.Controls.Add(Me.optMonthly)
        Me.fraRentalPeriod.Controls.Add(Me.optWeekly)
        Me.fraRentalPeriod.Controls.Add(Me.optDaily)
        Me.fraRentalPeriod.Controls.Add(Me.optHalfDay)
        Me.fraRentalPeriod.Controls.Add(Me.optHour)
        Me.fraRentalPeriod.Controls.Add(Me.Line1)
        Me.fraRentalPeriod.Controls.Add(Me._Label1_7)
        Me.fraRentalPeriod.Controls.Add(Me._Label1_4)
        Me.fraRentalPeriod.Controls.Add(Me._Label1_1)
        Me.fraRentalPeriod.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraRentalPeriod.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraRentalPeriod.Location = New System.Drawing.Point(8, 116)
        Me.fraRentalPeriod.Name = "fraRentalPeriod"
        Me.fraRentalPeriod.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraRentalPeriod.Size = New System.Drawing.Size(413, 176)
        Me.fraRentalPeriod.TabIndex = 20
        Me.fraRentalPeriod.TabStop = False
        Me.fraRentalPeriod.Text = "Rental Period"
        '
        'txtPeriods
        '
        Me.txtPeriods.BackColor = System.Drawing.SystemColors.Window
        Me.txtPeriods.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtPeriods.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPeriods.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPeriods.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "45", "60", "90"})
        Me.txtPeriods.Location = New System.Drawing.Point(12, 118)
        Me.txtPeriods.Name = "txtPeriods"
        Me.txtPeriods.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPeriods.Size = New System.Drawing.Size(55, 22)
        Me.txtPeriods.TabIndex = 0
        Me.txtPeriods.Text = "1"
        '
        'optWeekEnd
        '
        Me.optWeekEnd.BackColor = System.Drawing.SystemColors.Control
        Me.optWeekEnd.Cursor = System.Windows.Forms.Cursors.Default
        Me.optWeekEnd.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optWeekEnd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optWeekEnd.Location = New System.Drawing.Point(186, 61)
        Me.optWeekEnd.Name = "optWeekEnd"
        Me.optWeekEnd.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optWeekEnd.Size = New System.Drawing.Size(89, 21)
        Me.optWeekEnd.TabIndex = 11
        Me.optWeekEnd.Text = "Week &End"
        Me.optWeekEnd.UseVisualStyleBackColor = False
        '
        'txtDeposit
        '
        Me.txtDeposit.AcceptsReturn = True
        Me.txtDeposit.BackColor = System.Drawing.SystemColors.Window
        Me.txtDeposit.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDeposit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDeposit.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDeposit.Location = New System.Drawing.Point(322, 112)
        Me.txtDeposit.MaxLength = 0
        Me.txtDeposit.Name = "txtDeposit"
        Me.txtDeposit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDeposit.Size = New System.Drawing.Size(79, 19)
        Me.txtDeposit.TabIndex = 2
        Me.txtDeposit.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtPrice
        '
        Me.txtPrice.AcceptsReturn = True
        Me.txtPrice.BackColor = System.Drawing.SystemColors.Window
        Me.txtPrice.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPrice.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPrice.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPrice.Location = New System.Drawing.Point(322, 90)
        Me.txtPrice.MaxLength = 0
        Me.txtPrice.Name = "txtPrice"
        Me.txtPrice.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPrice.Size = New System.Drawing.Size(79, 19)
        Me.txtPrice.TabIndex = 1
        Me.txtPrice.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtItemTotal
        '
        Me.txtItemTotal.AcceptsReturn = True
        Me.txtItemTotal.BackColor = System.Drawing.SystemColors.Window
        Me.txtItemTotal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtItemTotal.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtItemTotal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtItemTotal.Location = New System.Drawing.Point(322, 142)
        Me.txtItemTotal.MaxLength = 0
        Me.txtItemTotal.Name = "txtItemTotal"
        Me.txtItemTotal.ReadOnly = True
        Me.txtItemTotal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtItemTotal.Size = New System.Drawing.Size(79, 19)
        Me.txtItemTotal.TabIndex = 3
        Me.txtItemTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtSaleItem
        '
        Me.txtSaleItem.AcceptsReturn = True
        Me.txtSaleItem.BackColor = System.Drawing.SystemColors.Window
        Me.txtSaleItem.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSaleItem.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSaleItem.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSaleItem.Location = New System.Drawing.Point(96, 149)
        Me.txtSaleItem.MaxLength = 0
        Me.txtSaleItem.Name = "txtSaleItem"
        Me.txtSaleItem.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSaleItem.Size = New System.Drawing.Size(75, 19)
        Me.txtSaleItem.TabIndex = 19
        Me.txtSaleItem.TabStop = False
        Me.txtSaleItem.Tag = "(No Auto Formatting)"
        Me.txtSaleItem.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'optSaleItem
        '
        Me.optSaleItem.BackColor = System.Drawing.SystemColors.Control
        Me.optSaleItem.Cursor = System.Windows.Forms.Cursors.Default
        Me.optSaleItem.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optSaleItem.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optSaleItem.Location = New System.Drawing.Point(8, 149)
        Me.optSaleItem.Name = "optSaleItem"
        Me.optSaleItem.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optSaleItem.Size = New System.Drawing.Size(78, 19)
        Me.optSaleItem.TabIndex = 12
        Me.optSaleItem.Text = "&Sell Item"
        Me.optSaleItem.UseVisualStyleBackColor = False
        '
        'optMonthly
        '
        Me.optMonthly.BackColor = System.Drawing.SystemColors.Control
        Me.optMonthly.Cursor = System.Windows.Forms.Cursors.Default
        Me.optMonthly.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optMonthly.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optMonthly.Location = New System.Drawing.Point(186, 39)
        Me.optMonthly.Name = "optMonthly"
        Me.optMonthly.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optMonthly.Size = New System.Drawing.Size(71, 17)
        Me.optMonthly.TabIndex = 10
        Me.optMonthly.Text = "&Monthly"
        Me.optMonthly.UseVisualStyleBackColor = False
        '
        'optWeekly
        '
        Me.optWeekly.BackColor = System.Drawing.SystemColors.Control
        Me.optWeekly.Cursor = System.Windows.Forms.Cursors.Default
        Me.optWeekly.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optWeekly.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optWeekly.Location = New System.Drawing.Point(186, 16)
        Me.optWeekly.Name = "optWeekly"
        Me.optWeekly.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optWeekly.Size = New System.Drawing.Size(65, 17)
        Me.optWeekly.TabIndex = 9
        Me.optWeekly.Text = "&Weekly"
        Me.optWeekly.UseVisualStyleBackColor = False
        '
        'optDaily
        '
        Me.optDaily.BackColor = System.Drawing.SystemColors.Control
        Me.optDaily.Cursor = System.Windows.Forms.Cursors.Default
        Me.optDaily.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optDaily.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optDaily.Location = New System.Drawing.Point(10, 59)
        Me.optDaily.Name = "optDaily"
        Me.optDaily.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optDaily.Size = New System.Drawing.Size(77, 21)
        Me.optDaily.TabIndex = 8
        Me.optDaily.Text = "&Daily"
        Me.optDaily.UseVisualStyleBackColor = False
        '
        'optHalfDay
        '
        Me.optHalfDay.BackColor = System.Drawing.SystemColors.Control
        Me.optHalfDay.Cursor = System.Windows.Forms.Cursors.Default
        Me.optHalfDay.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optHalfDay.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optHalfDay.Location = New System.Drawing.Point(10, 34)
        Me.optHalfDay.Name = "optHalfDay"
        Me.optHalfDay.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optHalfDay.Size = New System.Drawing.Size(70, 25)
        Me.optHalfDay.TabIndex = 7
        Me.optHalfDay.Text = "&Half Day/ 4 Hrs"
        Me.optHalfDay.UseVisualStyleBackColor = False
        '
        'optHour
        '
        Me.optHour.BackColor = System.Drawing.SystemColors.Control
        Me.optHour.Cursor = System.Windows.Forms.Cursors.Default
        Me.optHour.Enabled = False
        Me.optHour.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optHour.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optHour.Location = New System.Drawing.Point(12, 16)
        Me.optHour.Name = "optHour"
        Me.optHour.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optHour.Size = New System.Drawing.Size(53, 15)
        Me.optHour.TabIndex = 6
        Me.optHour.TabStop = True
        Me.optHour.Text = "Hou&rly"
        Me.optHour.UseVisualStyleBackColor = False
        '
        'Line1
        '
        Me.Line1.BackColor = System.Drawing.SystemColors.WindowText
        Me.Line1.Location = New System.Drawing.Point(316, 136)
        Me.Line1.Name = "Line1"
        Me.Line1.Size = New System.Drawing.Size(85, 1)
        Me.Line1.TabIndex = 32
        '
        '_Label1_7
        '
        Me._Label1_7.AutoSize = True
        Me._Label1_7.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_7.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_7, CType(7, Short))
        Me._Label1_7.Location = New System.Drawing.Point(264, 144)
        Me._Label1_7.Name = "_Label1_7"
        Me._Label1_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_7.Size = New System.Drawing.Size(51, 14)
        Me._Label1_7.TabIndex = 30
        Me._Label1_7.Text = "Item Total"
        '
        '_Label1_4
        '
        Me._Label1_4.AutoSize = True
        Me._Label1_4.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_4, CType(4, Short))
        Me._Label1_4.Location = New System.Drawing.Point(242, 92)
        Me._Label1_4.Name = "_Label1_4"
        Me._Label1_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_4.Size = New System.Drawing.Size(79, 14)
        Me._Label1_4.TabIndex = 29
        Me._Label1_4.Text = "Extended Price"
        '
        '_Label1_1
        '
        Me._Label1_1.AutoSize = True
        Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_1, CType(1, Short))
        Me._Label1_1.Location = New System.Drawing.Point(10, 99)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.Size = New System.Drawing.Size(101, 14)
        Me._Label1_1.TabIndex = 21
        Me._Label1_1.Text = "# Periods / Quantity"
        '
        'LinkLabel1
        '
        Me.LinkLabel1.Location = New System.Drawing.Point(96, 308)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(88, 32)
        Me.LinkLabel1.TabIndex = 23
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "How Do I rent for a day and a half?"
        Me.LinkLabel1.Visible = False
        '
        'frmAddToList
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(424, 356)
        Me.ControlBox = False
        Me.Controls.Add(Me.cmdCancelOrder)
        Me.Controls.Add(Me.btnAddToList)
        Me.Controls.Add(Me.cmdReserve)
        Me.Controls.Add(Me.LinkLabel1)
        Me.Controls.Add(Me.fraItemDesc)
        Me.Controls.Add(Me.fraRentalPeriod)
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(41, 102)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmAddToList"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Add Selected Item To Cart"
        Me.TopMost = True
        Me.fraItemDesc.ResumeLayout(False)
        Me.fraItemDesc.PerformLayout()
        Me.fraRentalPeriod.ResumeLayout(False)
        Me.fraRentalPeriod.PerformLayout()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
#End Region

#Region "Private and Public Variables"
    Dim oItem As CItems
    Dim mbCancel As Boolean
    Dim mcurSelectedPrice As Decimal
    Dim miNbrPeriods As Short
    Dim msPeriod As String
    Dim mbWait As Boolean
    Dim mbFormLoading As Boolean = True
    Dim bNoise As Boolean = True
    Public ItemID As String
    Public ItemName As String
    Public PriceID As Integer
    Public MeterRequired As Boolean
    Public MeterHours As Single
    Public dailyRate As Decimal
    Public halfDayRate As Decimal
    Public hourRate As Decimal
    Public weeklyRate As Decimal
    Public monthlyRate As Decimal
    Public weekendRate As Decimal
    Public depositRate As Decimal
    Private SQL As String
    Dim oDA As New CDataAccess()
    Public NbrRowsSelected As Integer = 0
    Public alMultiItems As ArrayList
#End Region


#Region "Processing Methods"
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


    Private Sub TotalTheItem()
        Me.txtPrice.Text = FormatCurrency(Val(Me.txtPeriods.Text) * mcurSelectedPrice)
        TotalItUp()
    End Sub
    Private Function MarkItemSold(ByVal ItemId As String) As Boolean
        Dim Sql As String
        Dim oDA As New CDataAccess()
        Dim sErr As String

        Sql = "update equipment "
        Sql &= "set available = 'SOLD HOLD', "
        Sql &= "rented_date = #" & Now.ToString & "# "
        Sql &= "where equip_id = '" & ItemId & "'"
        If oDA.SendActionSql(Sql, ConnectString, sErr) <> 1 Then
            MsgBox("Database error, can't mark equipment as sold.", MsgBoxStyle.Critical)
            Return False
        Else
            Return True
        End If
    End Function
#End Region

#Region "Form & Control Events"
    Private Sub frmAddToList_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Dim oRES As CTransaction
        'Dim dt As New DataTable()
        'Dim oDA As New CDataAccess()
        Dim decDeposit As Decimal
        'Dim SQL As String = ""
        'SQL &= "select * from rental_rates "
        'SQL &= "where price_id = " & PriceID & " "
        'Dim i As Integer = oDA.SendQuery(SQL, dt, ConnectString)
        Dim ItemTotal As Decimal

        Try
            With Me
                bNoise = False
                .txtItemID.Text = ItemID
                .txtItemName.Text = ItemName
                .txtItemDesc.Text = ""
                'If i > 0 Then
                .txtDaily.Text = FormatCurrency(Me.dailyRate)
                '.txtDaily.Text = FormatCurrency(MND(dt.Rows(0).Item("daily")))
                .txtDaily.Text = FormatCurrency(Me.dailyRate)
                .txtHalfDay.Text = FormatCurrency(Me.halfDayRate)
                If Val(UnFormat(.txtHalfDay.Text)) = 0 Then
                    .optHalfDay.Enabled = False
                    .txtHalfDay.Enabled = False
                End If
                .txtHourly.Text = FormatCurrency(Me.hourRate)
                '.txtHourly.Text = FormatCurrency(MND(dt.Rows(0).Item("HourRate")))
                .txtWeekly.Text = FormatCurrency(Me.weeklyRate)
                '.txtWeekly.Text = FormatCurrency(MND(dt.Rows(0).Item("Weekly")))
                .txtMonthly.Text = FormatCurrency(Me.monthlyRate)
                '.txtMonthly.Text = FormatCurrency(MND(dt.Rows(0).Item("Monthly")))
                .txtWeekEnd.Text = FormatCurrency(Me.weekendRate)
                '.txtWeekEnd.Text = FormatCurrency(MND(dt.Rows(0).Item("WeekEnd")))
                'Else
                'End If

                '.txtDelivery = FormatCurrency(Me.Delivery)
                If UseDeposits Then
                    .txtDeposit.Text = FormatCurrency(Me.depositRate)
                    '.txtDeposit.Text = FormatCurrency(MND(dt.Rows(0).Item("Deposit")))
                Else
                    .txtDeposit.Text = FormatCurrency(0)
                    .txtDeposit.Enabled = False
                    bNoise = True
                    .chkDeposit.Checked = False
                    bNoise = False
                    .chkDeposit.Enabled = False
                End If
                If modMain.UseHourlyRates Then
                    Me.optHour.Enabled = True
                    Me.txtHourly.Enabled = True
                End If
                .txtItemTotal.Text = ""
            End With
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

    Private Sub frmAddToList_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) _
        Handles MyBase.Activated, MyBase.Enter
        If mbFormLoading Then
            mbFormLoading = False
            If NbrRowsSelected > 0 Then
                Me.lblNbrItems.Visible = True
                Me.txtMultiCount.Visible = True
                Me.txtMultiCount.Text = NbrRowsSelected
                Me.txtSaleItem.Visible = False
                Me.optSaleItem.Visible = False
            End If

            With Me
                If Not UseHourlyRates Then
                    .optHour.Enabled = UseHourlyRates
                    '.txtHourly.Text = ""
                    .optHour.Checked = False
                    .txtHourly.Enabled = False
                    .optHalfDay.Checked = False
                End If
                .txtPeriods.Focus()
            End With
        End If
    End Sub


    Private Sub frmAddToList_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        With Me
            If KeyAscii = 27 Then
                mbCancel = True
                mbWait = False
                Me.Close()
                System.Windows.Forms.Application.DoEvents()
            End If

            Select Case UCase(Chr(KeyAscii))
                Case "A"
                    mbCancel = False
                    mbWait = False
                Case "C"
                    mbCancel = True
                    mbWait = False
                Case "H"
                    .optHalfDay.Checked = True
                Case "D"
                    .optDaily.Checked = True
                Case "W"
                    .optWeekly.Checked = True
                Case "M"
                    .optMonthly.Checked = True
                Case "E"
                    .optWeekEnd.Checked = True
                Case "S"
                    .optSaleItem.Checked = True
                Case "P"
                    .chkDeposit.CheckState = IIf(.chkDeposit.CheckState = 0, 1, 0)
                    'Case 1 - 9
                    '   .txtPeriods.Text = Chr(KeyCode)
                Case Else
                    GoTo EventExitSub
            End Select
            KeyAscii = 0
        End With
EventExitSub:
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub


    Private Sub txtPeriods_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPeriods.SelectedIndexChanged

        Try
            Me.txtPrice.Text = FormatCurrency(Val(Me.txtPeriods.Text) * mcurSelectedPrice)
            TotalItUp()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub txtweekend_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtWeekEnd.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then KeyAscii = 0
        KeyAscii = CkKeyPressNumeric(KeyAscii, txtWeekEnd)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub



    Private Sub optWeekEnd_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optWeekEnd.CheckedChanged
        If eventSender.Checked Then
            With Me
                .txtPeriods.Focus()
                mcurSelectedPrice = UnFormat((Me.txtWeekEnd.Text))
                msPeriod = "WeekEnd"
                Me.btnAddToList.Enabled = True
                TotalTheItem()
            End With
        End If
    End Sub

    ''' <summary>
    ''' Add the selected equipment to the TempItems table, and also
    ''' place the equipment on hold.  If the equipment has a meter,
    ''' build an entry in the meter table.
    ''' </summary>
    ''' <param name = "eventSender"></param>
    ''' <param name = "eventArgs"></param>
    Private Sub btnAddToList_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnAddToList.Click

        Dim oRES As CTransaction
        Dim dt As New DataTable()
        Dim decDeposit As Decimal
        Dim i As Short
        Dim j As Integer
        Dim k As Integer
        Dim sLastTP As String = ""
        Dim sCurrTP
        Dim oCO As New CCheckOut()

        Try
            If Me.NbrRowsSelected = 0 Then
                Me.NbrRowsSelected = 1
            End If
            Dim SQL As String = ""

            ' set up loop for multiple items
            For k = 0 To Me.NbrRowsSelected - 1
                ' if this is multi item selection
                ' get the pointer to the selected datatable item
                If NbrRowsSelected > 1 Then
                    j = alMultiItems(k)
                    With modMain.fMainForm
                        Me.ItemID = .dbgEquipment.DataSource.rows(j).item("equip_id")
                        Me.ItemName = .dbgEquipment.DataSource.rows(j).item("equip_name")
                        Me.PriceID = .dbgEquipment.DataSource.rows(j).item("price_id")
                        Me.MeterRequired = .dbgEquipment.DataSource.rows(j).item("meter_required")
                        Me.MeterHours = 0
                    End With
                End If

                ' make sure another user hasn't held equipment while
                ' we were up in this form
                If Not oCO.IsEquipAvailable(Me.txtItemID.Text, True) Then
                    mbCancel = True
                    mbWait = False
                    Me.Close()
                    System.Windows.Forms.Application.DoEvents()
                    Exit Sub
                End If

                With Me
                    If .optSaleItem.Checked Then
                        If UnFormat(.txtSaleItem.Text) > 0 Then
                            Dim sMsg As String
                            Dim iRV As Integer
                            sMsg = "Are you sure that the sale price is correct?  Click Yes" & Chr(10)
                            sMsg &= "if the price is correct, otherwise No." & Chr(10)
                            sMsg &= "" & Chr(10)
                            iRV = MsgBox(sMsg, CType(36, Microsoft.VisualBasic.MsgBoxStyle), "Confirm Sales Price")

                            If iRV = 6 Then
                                ' Yes Code goes here
                                MarkItemSold(Me.txtItemID.Text.Trim)
                            Else
                                ' No code goes here
                                Exit Sub
                            End If
                        Else
                            MsgBox("You must enter the price of the equipment you are selling.", MsgBoxStyle.Information)
                            Exit Sub
                        End If
                    Else
                    End If
                End With

                Dim oCR As New CCheckReservations()
                Dim rentPeriod As String
                Select Case True
                    Case Me.optDaily.Checked : rentPeriod = DAILY
                    Case Me.optHalfDay.Checked : rentPeriod = HALF_DAY
                    Case Me.optWeekEnd.Checked : rentPeriod = WEEK_END
                    Case Me.optWeekly.Checked : rentPeriod = WEEKLY
                    Case Me.optMonthly.Checked : rentPeriod = MONTHLY
                    Case Me.optHour.Checked : rentPeriod = HOURLY
                End Select

                If Not oCR.IsRentable(PriceID, Now, rentPeriod, Me.txtPeriods.Text) Then
                    Exit Sub
                End If

                mbCancel = False
                mbWait = False
                ' write a temporary dataset to remember the
                ' state of the order
                If Me.txtItemName.Text.Length > 29 Then
                    Me.txtItemName.Text = Me.txtItemName.Text.Substring(0, 29)
                End If
                SQL = "Insert into TempItems (ItemID,ItemName,ItemCount, "
                SQL &= "ItemPeriod,ItemPrice,ItemExtendedPrice,ItemDeposit,"
                SQL &= "rentorsale,meter_required,hour_meter,user_id, "
                SQL &= "hourrate,halfday,daily,weekly,monthly,weekend,newprices) " ',hourly,halfday,daily,weekly,monthly ) "
                SQL &= "Values("
                SQL &= "'" & Me.ItemID & "', "
                SQL &= "'" & Replace(Me.ItemName, "'", "''") & "', "
                SQL &= Val(Me.txtPeriods.Text) & ", "
                Dim iPeriods As Integer = Val(Me.txtPeriods.Text)

                Dim decPrice As Decimal
                If Me.optDaily.Checked Then
                    SQL &= "'" & DAILY & "', "
                    decPrice = UnFormat(Me.txtDaily.Text)
                    SQL &= decPrice & ", "
                    sCurrTP = DAILY
                ElseIf Me.optHalfDay.Checked Then
                    SQL &= "'" & HALF_DAY & "', "
                    decPrice = UnFormat(Me.txtHalfDay.Text)
                    SQL &= decPrice & ", "
                    sCurrTP = HALF_DAY
                ElseIf Me.optMonthly.Checked Then
                    SQL &= "'" & MONTHLY & "', "
                    decPrice = UnFormat(Me.txtMonthly.Text)
                    SQL &= decPrice & ", "
                    sCurrTP = MONTHLY
                ElseIf Me.optWeekEnd.Checked Then
                    SQL &= "'" & WEEK_END & "', "
                    decPrice = UnFormat(Me.txtWeekEnd.Text)
                    SQL &= decPrice & ", "
                    sCurrTP = WEEK_END
                ElseIf Me.optWeekly.Checked Then
                    SQL &= "'" & WEEKLY & "', "
                    decPrice = UnFormat(Me.txtWeekly.Text)
                    SQL &= decPrice & ", "
                    sCurrTP = WEEKLY
                ElseIf Me.optSaleItem.Checked Then
                    SQL &= "'" & SALE & "', "
                    decPrice = UnFormat(Me.txtSaleItem.Text)
                    SQL &= decPrice & ", "
                    sCurrTP = SALE
                End If
                decPrice *= iPeriods
                SQL &= (decPrice) & ", "
                SQL &= UnFormat(Me.txtDeposit.Text) & ", "
                If Me.optSaleItem.Checked Then
                    SQL &= "'" & "SOLD" & "', "
                Else
                    SQL &= "'" & RENT & "', "
                End If
                SQL &= Me.MeterRequired & ", "
                If Me.MeterRequired Then
                    SQL &= Me.MeterHours & " "
                Else
                    SQL &= "0"
                End If
                SQL &= ",'" & UserName & "', "
                SQL &= UnFormat(Me.txtHourly.Text) & ", "
                SQL &= UnFormat(Me.txtHalfDay.Text) & ", "
                SQL &= UnFormat(Me.txtDaily.Text) & ", "
                SQL &= UnFormat(Me.txtWeekly.Text) & ", "
                SQL &= UnFormat(Me.txtMonthly.Text) & ", "
                SQL &= UnFormat(Me.txtWeekEnd.Text) & ", "
                SQL &= True & " "
                SQL &= ")"
                ' ensure that this item is being rented  for the same 
                ' time period as any other items
                Dim sql2 As String
                sql2 = "select * from tempitems where (rentorsale = '" & RENT & "' "
                sql2 &= "or itemid = '" & RERENT & "') "
                sql2 &= "and user_id = '" & UserName & "'"
                dt.Reset()

                If oDA.SendQuery(sql2, dt, ConnectString) > 0 Then
                    For i = 0 To dt.Rows.Count - 1
                        With dt.Rows(i)
                            sLastTP = IIf(IsDBNull(.Item("itemperiod")), "", .Item("itemperiod"))
                            If sLastTP.Trim <> sCurrTP Then
                                MsgBox("The time period for all rented items must be the same for all items on an invoice.", MsgBoxStyle.Exclamation)
                                Exit Sub
                            End If
                        End With
                    Next
                End If

                ' place temp hold on equip
                If Not Me.optSaleItem.Checked Then
                    ' ensure that it is available and mark as unavailable
                    ' in equipment list
                    oRES = New CTransaction()
                    With oRES
                        If Not .ReserveTemp(ItemID) Then
                            GoTo DisplayExit
                        End If
                    End With
                End If


                Dim sErr As String
                If oDA.SendActionSql(SQL, ConnectString, sErr) <> 1 Then
                    Throw New System.Exception("Database error, unable to create temp invoice item.")
                End If
DisplayExit:
                oRES = Nothing
                System.Windows.Forms.Application.DoEvents()
            Next k
            Me.Close()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

    Private Sub optDaily_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optDaily.CheckedChanged
        If bNoise Then Exit Sub
        If eventSender.Checked Then
            With Me
                .txtPeriods.Enabled = True
                .txtPeriods.Focus()
                mcurSelectedPrice = UnFormat((Me.txtDaily.Text))
                msPeriod = "Day"
                Me.btnAddToList.Enabled = True
                TotalTheItem()
            End With
        End If
    End Sub

    Private Sub optHalfDay_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optHalfDay.CheckedChanged
        If bNoise Then Exit Sub
        If eventSender.Checked Then
            Me.txtPeriods.Enabled = False
            Me.txtPeriods.Text = "1"
            mcurSelectedPrice = UnFormat((Me.txtHalfDay.Text))
            msPeriod = HALF_DAY
            Me.btnAddToList.Enabled = True
            TotalTheItem()
        End If
    End Sub


    Private Sub optHour_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optHour.CheckedChanged
        If bNoise Then Exit Sub
        If eventSender.Checked Then
            On Error Resume Next
            With Me
                .txtPeriods.Enabled = True
                .txtPeriods.Focus()
                mcurSelectedPrice = UnFormat((Me.txtHourly.Text))
                msPeriod = "Hours"
                Me.btnAddToList.Enabled = True
                TotalTheItem()
            End With
        End If
    End Sub

    Private Sub optMonthly_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optMonthly.CheckedChanged
        If bNoise Then Exit Sub
        If eventSender.Checked Then
            With Me
                .txtPeriods.Enabled = True
                .txtPeriods.Focus()
                mcurSelectedPrice = UnFormat((Me.txtMonthly.Text))
                msPeriod = "Month"
                Me.btnAddToList.Enabled = True
                TotalTheItem()
            End With
        End If
    End Sub

    Private Sub optSaleItem_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSaleItem.CheckedChanged
        If bNoise Then Exit Sub
        Static Busy As Boolean
        If Busy Then Exit Sub
        Busy = True
        If eventSender.Checked Then
            Dim sMsg As String
            Dim iRV As Integer
            sMsg = "If you are absolutely sure you want to sell this piece" & Chr(10)
            sMsg &= "of equipment, click the Yes button and enter the price" & Chr(10)
            sMsg &= "of the equipment into the Sell Item Price box." & Chr(10)
            sMsg &= "" & Chr(10)
            sMsg &= "If you do not want to sell the item click No and click" & Chr(10)
            sMsg &= "the desired rental period." & Chr(10)
            sMsg &= "" & Chr(10)
            iRV = MsgBox(sMsg, CType(52, Microsoft.VisualBasic.MsgBoxStyle), "Confirm Equipment Sale")

            If iRV = 6 Then
                ' Yes Code goes here
            Else
                ' No code goes here
                Busy = False
                Exit Sub
            End If
            With Me
                .txtPeriods.Enabled = False
                .txtPeriods.Text = "1"
                .txtSaleItem.Enabled = True
                Me.btnAddToList.Enabled = True
                Me.optSaleItem.Checked = True
                Busy = False
            End With
        End If
        Busy = False
    End Sub

    Private Sub optWeekly_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optWeekly.CheckedChanged
        If bNoise Then Exit Sub
        If eventSender.Checked Then
            With Me
                .txtPeriods.Enabled = True
                .txtPeriods.Focus()
                mcurSelectedPrice = UnFormat((Me.txtWeekly.Text))
                msPeriod = WEEKLY
                Me.btnAddToList.Enabled = True
                TotalTheItem()
            End With
        End If
    End Sub

    Private Sub txtDelivery_LostFocus()
        TotalItUp()
    End Sub


    Private Sub txtDeposit_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDeposit.Leave
        TotalItUp()
    End Sub

    Private Sub txtPeriods_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPeriods.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then KeyAscii = 0
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtPeriods_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtPeriods.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error Resume Next
        If KeyCode = 13 Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
        If KeyCode = System.Windows.Forms.Keys.Up Then System.Windows.Forms.SendKeys.SendWait("+{TAB}")
        If KeyCode = System.Windows.Forms.Keys.Down Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
    End Sub
    Private Sub txtPeriods_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPeriods.Enter
        txtPeriods.SelectionStart = 0
        txtPeriods.SelectionLength = Len(Trim(txtPeriods.Text))
    End Sub


    Private Sub cmdNo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancelOrder.Click
        mbCancel = True
        mbWait = False
        Me.Close()
        System.Windows.Forms.Application.DoEvents()
    End Sub

    Private Sub chkDeposit_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkDeposit.CheckStateChanged
        If bNoise Then Exit Sub
        With Me
            If .chkDeposit.CheckState = System.Windows.Forms.CheckState.Checked Then
                .txtDeposit.Visible = True
            Else
                .txtDeposit.Visible = False
            End If
            TotalItUp()
        End With
    End Sub

    Private Sub txtItemTotal_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtItemTotal.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error Resume Next
        If KeyCode = 13 Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
    End Sub
    Private Sub txtItemTotal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtItemTotal.Enter
        txtItemTotal.SelectionStart = 0
        txtItemTotal.SelectionLength = Len(Trim(txtItemTotal.Text))
    End Sub
    Private Sub txtItemTotal_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtItemTotal.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then KeyAscii = 0
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtItemTotal_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtItemTotal.Leave
        txtItemTotal.Text = UCase(txtItemTotal.Text)
    End Sub

    Private Sub txtDeposit_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDeposit.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then KeyAscii = 0
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtDeposit_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtDeposit.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error Resume Next
        If KeyCode = 13 Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
        If KeyCode = System.Windows.Forms.Keys.Up Then System.Windows.Forms.SendKeys.SendWait("+{TAB}")
        If KeyCode = System.Windows.Forms.Keys.Down Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
    End Sub
    Private Sub txtDeposit_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDeposit.Enter
        txtDeposit.SelectionStart = 0
        txtDeposit.SelectionLength = Len(Trim(txtDeposit.Text))
    End Sub


    Private Sub txtPrice_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPrice.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then KeyAscii = 0
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtPrice_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtPrice.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error Resume Next
        If KeyCode = 13 Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
        If KeyCode = System.Windows.Forms.Keys.Up Then System.Windows.Forms.SendKeys.SendWait("+{TAB}")
        If KeyCode = System.Windows.Forms.Keys.Down Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
    End Sub
    Private Sub txtPrice_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPrice.Enter
        txtPrice.SelectionStart = 0
        txtPrice.SelectionLength = Len(Trim(txtPrice.Text))
    End Sub


    Private Sub txtPeriods_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPeriods.Leave
        Me.txtPrice.Text = FormatCurrency(Val(Me.txtPeriods.Text) * mcurSelectedPrice)
        TotalItUp()
    End Sub


    Private Sub txtPrice_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPrice.Leave
        TotalItUp()
    End Sub
    Private Sub TotalItUp()
        Me.txtItemTotal.Text = FormatCurrency(UnFormat((Me.txtPrice).Text) + _
           IIf(Me.chkDeposit.CheckState = _
           System.Windows.Forms.CheckState.Checked, _
           UnFormat((Me.txtDeposit).Text), 0))
    End Sub

    Private Sub txtHalfDay_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtHalfDay.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtHalfDay_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtHalfDay.KeyDown
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
        If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
        If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
    End Sub
    Private Sub txtHalfDay_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHalfDay.Enter
        txtHalfDay.SelectionStart = 0
        txtHalfDay.SelectionLength = txtHalfDay.Text.Trim.Length
    End Sub
    Private Sub txtWeekly_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtWeekly.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtWeekly_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtWeekly.KeyDown
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
        If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
        If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
    End Sub
    Private Sub txtWeekly_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtWeekly.Enter
        txtWeekly.SelectionStart = 0
        txtWeekly.SelectionLength = txtWeekly.Text.Trim.Length
    End Sub
    Private Sub txtWeekEnd_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtWeekEnd.KeyDown
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
        If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
        If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
    End Sub
    Private Sub txtWeekEnd_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtWeekEnd.Enter
        txtWeekEnd.SelectionStart = 0
        txtWeekEnd.SelectionLength = txtWeekEnd.Text.Trim.Length
    End Sub
    Private Sub txtSaleItem_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSaleItem.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtSaleItem_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSaleItem.KeyDown
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
        If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
        If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
    End Sub
    Private Sub txtSaleItem_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSaleItem.Enter
        txtSaleItem.SelectionStart = 0
        txtSaleItem.SelectionLength = txtSaleItem.Text.Trim.Length
    End Sub
    Private Sub txtDaily_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDaily.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtDaily_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDaily.KeyDown
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
        If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
        If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
    End Sub
    Private Sub txtDaily_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDaily.Enter
        txtDaily.SelectionStart = 0
        txtDaily.SelectionLength = txtDaily.Text.Trim.Length
    End Sub
    Private Sub txtMonthly_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMonthly.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtMonthly_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtMonthly.KeyDown
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
        If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
        If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
    End Sub
    Private Sub txtMonthly_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMonthly.Enter
        txtMonthly.SelectionStart = 0
        txtMonthly.SelectionLength = txtMonthly.Text.Trim.Length
    End Sub

    Private Sub txtDaily_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDaily.Leave
        With Me
            .txtDaily.Text = FormatCurrency(.txtDaily.Text)
            .mcurSelectedPrice = UnFormat(.txtDaily.Text)
            TotalTheItem()
        End With
    End Sub

    Private Sub txtHalfDay_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHalfDay.Leave
        With Me
            .txtHalfDay.Text = FormatCurrency(.txtHalfDay.Text)
            .mcurSelectedPrice = UnFormat(.txtHalfDay.Text)
            TotalTheItem()
        End With
    End Sub

    Private Sub txtHourly_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHourly.Leave
        With Me
            .txtHourly.Text = FormatCurrency(.txtHourly.Text)
            .mcurSelectedPrice = UnFormat(.txtHourly.Text)
            TotalTheItem()
        End With
    End Sub

    Private Sub txtMonthly_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMonthly.Leave
        With Me
            .txtMonthly.Text = FormatCurrency(.txtMonthly.Text)
            .mcurSelectedPrice = UnFormat(.txtMonthly.Text)
            TotalTheItem()
        End With
    End Sub

    Private Sub txtWeekEnd_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtWeekEnd.Leave
        With Me
            .txtWeekEnd.Text = FormatCurrency(.txtWeekEnd.Text)
            .mcurSelectedPrice = UnFormat(.txtWeekEnd.Text)
            TotalTheItem()
        End With
    End Sub

    Private Sub txtWeekly_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtWeekly.Leave
        With Me
            .txtWeekly.Text = FormatCurrency(.txtWeekly.Text)
            .mcurSelectedPrice = UnFormat(.txtWeekly.Text)
            TotalTheItem()
        End With
    End Sub

    Private Sub txtSaleItem_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSaleItem.Leave
        Me.txtSaleItem.Text = FormatCurrency(Me.txtSaleItem.Text)
        Me.mcurSelectedPrice = UnFormat(Me.txtSaleItem.Text)
        TotalTheItem()
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim sTxt As String = ""
        sTxt &= "To rent a piece of equipment for a day and a half day, you "
        sTxt &= "must create two invoices because you are trying to rent one "
        sTxt &= "piece of equipment at two different rates on the same "
        sTxt &= "invoice.  That is not possible.  So, do the following:" & Chr(13) & Chr(10)
        sTxt &= "1) Rent the equipment for one day and check it out." & Chr(13) & Chr(10)
        sTxt &= "2) Rent the equiment for a half day and check it out." & Chr(13) & Chr(10)

        Dim oFrm As New frmHelp()
        oFrm.CannedMessage = sTxt
        oFrm.ShowDialog()
    End Sub

    ''' <summary>
    ''' Load the reservations form.
    ''' </summary>
    ''' <param name = "sender"></param>
    ''' <param name = "e"></param>
    Private Sub cmdReserve_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReserve.Click
        Dim oFrm As New frmReservations()
        oFrm.ShowDialog()
    End Sub
#End Region
End Class