Option Strict Off
Option Explicit On 
Imports VB = Microsoft.VisualBasic
Friend Class frmRentalRates
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
   Public WithEvents _lblLabels_7 As System.Windows.Forms.Label  
   Public WithEvents _lblLabels_8 As System.Windows.Forms.Label  
   Public WithEvents _lblLabels_9 As System.Windows.Forms.Label  
   Public WithEvents cmdAdd As System.Windows.Forms.Button  
   Public WithEvents cmdClose As System.Windows.Forms.Button  
   Public WithEvents cmdDelete As System.Windows.Forms.Button  
   Public WithEvents cmdUpdate As System.Windows.Forms.Button  
   Public WithEvents dbgRentalRates As System.Windows.Forms.DataGrid 
   Public WithEvents Frame1 As System.Windows.Forms.GroupBox  
   Public WithEvents lblLabels As Microsoft.VisualBasic.Compatibility.VB6.LabelArray  
   Public WithEvents txtDaily As System.Windows.Forms.TextBox  
   Public WithEvents txtDeposit As System.Windows.Forms.TextBox  
   Public WithEvents txtEquip_Name As System.Windows.Forms.TextBox  
   Public WithEvents txtHalfDay As System.Windows.Forms.TextBox  
   Public WithEvents txtHour As System.Windows.Forms.TextBox  
   Public WithEvents txtMinimum As System.Windows.Forms.TextBox  
   Public WithEvents txtMonthly As System.Windows.Forms.TextBox  
   Public WithEvents txtPrice_ID As System.Windows.Forms.TextBox  
   Public WithEvents txtWeekEnd As System.Windows.Forms.TextBox  
   Public WithEvents txtWeekly As System.Windows.Forms.TextBox  
   'NOTE: The following procedure is required by the Windows Form Designer
   'It can be modified using the Windows Form Designer.
   'Do not modify it using the code editor.
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
      Me.components = New System.ComponentModel.Container()
      Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmRentalRates))
      Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
      Me.cmdUpdate = New System.Windows.Forms.Button()
      Me.cmdClose = New System.Windows.Forms.Button()
      Me.cmdAdd = New System.Windows.Forms.Button()
      Me.cmdDelete = New System.Windows.Forms.Button()
      Me.Frame1 = New System.Windows.Forms.GroupBox()
      Me.txtMinimum = New System.Windows.Forms.TextBox()
      Me.txtDeposit = New System.Windows.Forms.TextBox()
      Me.txtWeekEnd = New System.Windows.Forms.TextBox()
      Me.txtMonthly = New System.Windows.Forms.TextBox()
      Me.txtWeekly = New System.Windows.Forms.TextBox()
      Me.txtDaily = New System.Windows.Forms.TextBox()
      Me.txtHalfDay = New System.Windows.Forms.TextBox()
      Me.txtHour = New System.Windows.Forms.TextBox()
      Me.txtPrice_ID = New System.Windows.Forms.TextBox()
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
      Me.dbgRentalRates = New System.Windows.Forms.DataGrid()
      Me.Frame1.SuspendLayout()
      CType(Me.lblLabels, System.ComponentModel.ISupportInitialize).BeginInit()
      CType(Me.dbgRentalRates, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.SuspendLayout()
      '
      'cmdUpdate
      '
      Me.cmdUpdate.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
      Me.cmdUpdate.BackColor = System.Drawing.SystemColors.Control
      Me.cmdUpdate.Cursor = System.Windows.Forms.Cursors.Default
      Me.cmdUpdate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.cmdUpdate.ForeColor = System.Drawing.SystemColors.ControlText
      Me.cmdUpdate.Location = New System.Drawing.Point(532, 234)
      Me.cmdUpdate.Name = "cmdUpdate"
      Me.cmdUpdate.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.cmdUpdate.Size = New System.Drawing.Size(77, 26)
      Me.cmdUpdate.TabIndex = 10
      Me.cmdUpdate.Text = "&Save"
      Me.ToolTip1.SetToolTip(Me.cmdUpdate, "Save changes or new record")
      '
      'cmdClose
      '
      Me.cmdClose.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
      Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
      Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
      Me.cmdClose.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
      Me.cmdClose.Location = New System.Drawing.Point(532, 328)
      Me.cmdClose.Name = "cmdClose"
      Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.cmdClose.Size = New System.Drawing.Size(77, 26)
      Me.cmdClose.TabIndex = 13
      Me.cmdClose.Text = "&Close"
      Me.ToolTip1.SetToolTip(Me.cmdClose, "Close without saving current changes")
      '
      'cmdAdd
      '
      Me.cmdAdd.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
      Me.cmdAdd.BackColor = System.Drawing.SystemColors.Control
      Me.cmdAdd.Cursor = System.Windows.Forms.Cursors.Default
      Me.cmdAdd.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.cmdAdd.ForeColor = System.Drawing.SystemColors.ControlText
      Me.cmdAdd.Location = New System.Drawing.Point(532, 264)
      Me.cmdAdd.Name = "cmdAdd"
      Me.cmdAdd.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.cmdAdd.Size = New System.Drawing.Size(77, 26)
      Me.cmdAdd.TabIndex = 11
      Me.cmdAdd.Text = "&Add"
      '
      'cmdDelete
      '
      Me.cmdDelete.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
      Me.cmdDelete.BackColor = System.Drawing.SystemColors.Control
      Me.cmdDelete.Cursor = System.Windows.Forms.Cursors.Default
      Me.cmdDelete.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.cmdDelete.ForeColor = System.Drawing.SystemColors.ControlText
      Me.cmdDelete.Location = New System.Drawing.Point(532, 296)
      Me.cmdDelete.Name = "cmdDelete"
      Me.cmdDelete.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.cmdDelete.Size = New System.Drawing.Size(77, 26)
      Me.cmdDelete.TabIndex = 12
      Me.cmdDelete.Text = "&Delete"
      '
      'Frame1
      '
      Me.Frame1.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
      Me.Frame1.BackColor = System.Drawing.SystemColors.Control
      Me.Frame1.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtMinimum, Me.txtDeposit, Me.txtWeekEnd, Me.txtMonthly, Me.txtWeekly, Me.txtDaily, Me.txtHalfDay, Me.txtHour, Me.txtPrice_ID, Me.txtEquip_Name, Me._lblLabels_0, Me._lblLabels_1, Me._lblLabels_2, Me._lblLabels_3, Me._lblLabels_4, Me._lblLabels_5, Me._lblLabels_6, Me._lblLabels_7, Me._lblLabels_8, Me._lblLabels_9})
      Me.Frame1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
      Me.Frame1.Location = New System.Drawing.Point(10, 226)
      Me.Frame1.Name = "Frame1"
      Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.Frame1.Size = New System.Drawing.Size(507, 129)
      Me.Frame1.TabIndex = 15
      Me.Frame1.TabStop = False
      Me.Frame1.Text = "Edit Rates"
      '
      'txtMinimum
      '
      Me.txtMinimum.AcceptsReturn = True
      Me.txtMinimum.AutoSize = False
      Me.txtMinimum.BackColor = System.Drawing.SystemColors.Window
      Me.txtMinimum.Cursor = System.Windows.Forms.Cursors.IBeam
      Me.txtMinimum.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.txtMinimum.ForeColor = System.Drawing.SystemColors.WindowText
      Me.txtMinimum.Location = New System.Drawing.Point(408, 102)
      Me.txtMinimum.MaxLength = 0
      Me.txtMinimum.Name = "txtMinimum"
      Me.txtMinimum.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.txtMinimum.Size = New System.Drawing.Size(85, 19)
      Me.txtMinimum.TabIndex = 9
      Me.txtMinimum.Text = ""
      Me.txtMinimum.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      '
      'txtDeposit
      '
      Me.txtDeposit.AcceptsReturn = True
      Me.txtDeposit.AutoSize = False
      Me.txtDeposit.BackColor = System.Drawing.SystemColors.Window
      Me.txtDeposit.Cursor = System.Windows.Forms.Cursors.IBeam
      Me.txtDeposit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.txtDeposit.ForeColor = System.Drawing.SystemColors.WindowText
      Me.txtDeposit.Location = New System.Drawing.Point(408, 80)
      Me.txtDeposit.MaxLength = 0
      Me.txtDeposit.Name = "txtDeposit"
      Me.txtDeposit.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.txtDeposit.Size = New System.Drawing.Size(85, 19)
      Me.txtDeposit.TabIndex = 8
      Me.txtDeposit.Tag = "$#,##0.00;($#,##0.00)"
      Me.txtDeposit.Text = ""
      Me.txtDeposit.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      '
      'txtWeekEnd
      '
      Me.txtWeekEnd.AcceptsReturn = True
      Me.txtWeekEnd.AutoSize = False
      Me.txtWeekEnd.BackColor = System.Drawing.SystemColors.Window
      Me.txtWeekEnd.Cursor = System.Windows.Forms.Cursors.IBeam
      Me.txtWeekEnd.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.txtWeekEnd.ForeColor = System.Drawing.SystemColors.WindowText
      Me.txtWeekEnd.Location = New System.Drawing.Point(408, 58)
      Me.txtWeekEnd.MaxLength = 0
      Me.txtWeekEnd.Name = "txtWeekEnd"
      Me.txtWeekEnd.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.txtWeekEnd.Size = New System.Drawing.Size(85, 19)
      Me.txtWeekEnd.TabIndex = 7
      Me.txtWeekEnd.Tag = "$#,##0.00;($#,##0.00)"
      Me.txtWeekEnd.Text = ""
      Me.txtWeekEnd.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      '
      'txtMonthly
      '
      Me.txtMonthly.AcceptsReturn = True
      Me.txtMonthly.AutoSize = False
      Me.txtMonthly.BackColor = System.Drawing.SystemColors.Window
      Me.txtMonthly.Cursor = System.Windows.Forms.Cursors.IBeam
      Me.txtMonthly.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.txtMonthly.ForeColor = System.Drawing.SystemColors.WindowText
      Me.txtMonthly.Location = New System.Drawing.Point(408, 36)
      Me.txtMonthly.MaxLength = 0
      Me.txtMonthly.Name = "txtMonthly"
      Me.txtMonthly.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.txtMonthly.Size = New System.Drawing.Size(85, 19)
      Me.txtMonthly.TabIndex = 6
      Me.txtMonthly.Tag = "$#,##0.00;($#,##0.00)"
      Me.txtMonthly.Text = ""
      Me.txtMonthly.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      '
      'txtWeekly
      '
      Me.txtWeekly.AcceptsReturn = True
      Me.txtWeekly.AutoSize = False
      Me.txtWeekly.BackColor = System.Drawing.SystemColors.Window
      Me.txtWeekly.Cursor = System.Windows.Forms.Cursors.IBeam
      Me.txtWeekly.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.txtWeekly.ForeColor = System.Drawing.SystemColors.WindowText
      Me.txtWeekly.Location = New System.Drawing.Point(408, 14)
      Me.txtWeekly.MaxLength = 0
      Me.txtWeekly.Name = "txtWeekly"
      Me.txtWeekly.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.txtWeekly.Size = New System.Drawing.Size(85, 19)
      Me.txtWeekly.TabIndex = 5
      Me.txtWeekly.Tag = "$#,##0.00;($#,##0.00)"
      Me.txtWeekly.Text = ""
      Me.txtWeekly.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      '
      'txtDaily
      '
      Me.txtDaily.AcceptsReturn = True
      Me.txtDaily.AutoSize = False
      Me.txtDaily.BackColor = System.Drawing.SystemColors.Window
      Me.txtDaily.Cursor = System.Windows.Forms.Cursors.IBeam
      Me.txtDaily.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.txtDaily.ForeColor = System.Drawing.SystemColors.WindowText
      Me.txtDaily.Location = New System.Drawing.Point(86, 102)
      Me.txtDaily.MaxLength = 0
      Me.txtDaily.Name = "txtDaily"
      Me.txtDaily.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.txtDaily.Size = New System.Drawing.Size(85, 19)
      Me.txtDaily.TabIndex = 4
      Me.txtDaily.Tag = "$#,##0.00;($#,##0.00)"
      Me.txtDaily.Text = ""
      Me.txtDaily.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      '
      'txtHalfDay
      '
      Me.txtHalfDay.AcceptsReturn = True
      Me.txtHalfDay.AutoSize = False
      Me.txtHalfDay.BackColor = System.Drawing.SystemColors.Window
      Me.txtHalfDay.Cursor = System.Windows.Forms.Cursors.IBeam
      Me.txtHalfDay.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.txtHalfDay.ForeColor = System.Drawing.SystemColors.WindowText
      Me.txtHalfDay.Location = New System.Drawing.Point(86, 80)
      Me.txtHalfDay.MaxLength = 0
      Me.txtHalfDay.Name = "txtHalfDay"
      Me.txtHalfDay.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.txtHalfDay.Size = New System.Drawing.Size(85, 19)
      Me.txtHalfDay.TabIndex = 3
      Me.txtHalfDay.Tag = "$#,##0.00;($#,##0.00)"
      Me.txtHalfDay.Text = ""
      Me.txtHalfDay.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      '
      'txtHour
      '
      Me.txtHour.AcceptsReturn = True
      Me.txtHour.AutoSize = False
      Me.txtHour.BackColor = System.Drawing.SystemColors.Window
      Me.txtHour.Cursor = System.Windows.Forms.Cursors.IBeam
      Me.txtHour.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.txtHour.ForeColor = System.Drawing.SystemColors.WindowText
      Me.txtHour.Location = New System.Drawing.Point(86, 58)
      Me.txtHour.MaxLength = 0
      Me.txtHour.Name = "txtHour"
      Me.txtHour.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.txtHour.Size = New System.Drawing.Size(85, 19)
      Me.txtHour.TabIndex = 2
      Me.txtHour.Tag = "$#,##0.00;($#,##0.00)"
      Me.txtHour.Text = ""
      Me.txtHour.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      '
      'txtPrice_ID
      '
      Me.txtPrice_ID.AcceptsReturn = True
      Me.txtPrice_ID.AutoSize = False
      Me.txtPrice_ID.BackColor = System.Drawing.Color.Cyan
      Me.txtPrice_ID.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
      Me.txtPrice_ID.Cursor = System.Windows.Forms.Cursors.IBeam
      Me.txtPrice_ID.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.txtPrice_ID.ForeColor = System.Drawing.SystemColors.WindowText
      Me.txtPrice_ID.Location = New System.Drawing.Point(86, 36)
      Me.txtPrice_ID.MaxLength = 0
      Me.txtPrice_ID.Name = "txtPrice_ID"
      Me.txtPrice_ID.ReadOnly = True
      Me.txtPrice_ID.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.txtPrice_ID.Size = New System.Drawing.Size(87, 19)
      Me.txtPrice_ID.TabIndex = 14
      Me.txtPrice_ID.TabStop = False
      Me.txtPrice_ID.Text = ""
      Me.txtPrice_ID.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      '
      'txtEquip_Name
      '
      Me.txtEquip_Name.AcceptsReturn = True
      Me.txtEquip_Name.AutoSize = False
      Me.txtEquip_Name.BackColor = System.Drawing.SystemColors.Window
      Me.txtEquip_Name.Cursor = System.Windows.Forms.Cursors.IBeam
      Me.txtEquip_Name.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.txtEquip_Name.ForeColor = System.Drawing.SystemColors.WindowText
      Me.txtEquip_Name.Location = New System.Drawing.Point(86, 14)
      Me.txtEquip_Name.MaxLength = 0
      Me.txtEquip_Name.Name = "txtEquip_Name"
      Me.txtEquip_Name.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.txtEquip_Name.Size = New System.Drawing.Size(245, 19)
      Me.txtEquip_Name.TabIndex = 1
      Me.txtEquip_Name.Text = ""
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
      Me._lblLabels_0.Size = New System.Drawing.Size(71, 13)
      Me._lblLabels_0.TabIndex = 25
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
      Me._lblLabels_1.Location = New System.Drawing.Point(12, 39)
      Me._lblLabels_1.Name = "_lblLabels_1"
      Me._lblLabels_1.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me._lblLabels_1.Size = New System.Drawing.Size(49, 13)
      Me._lblLabels_1.TabIndex = 24
      Me._lblLabels_1.Text = "Price_ID:"
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
      Me._lblLabels_2.Size = New System.Drawing.Size(31, 13)
      Me._lblLabels_2.TabIndex = 23
      Me._lblLabels_2.Text = "Hour:"
      '
      '_lblLabels_3
      '
      Me._lblLabels_3.BackColor = System.Drawing.SystemColors.Control
      Me._lblLabels_3.Cursor = System.Windows.Forms.Cursors.Default
      Me._lblLabels_3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me._lblLabels_3.ForeColor = System.Drawing.SystemColors.ControlText
      Me.lblLabels.SetIndex(Me._lblLabels_3, CType(3, Short))
      Me._lblLabels_3.Location = New System.Drawing.Point(12, 78)
      Me._lblLabels_3.Name = "_lblLabels_3"
      Me._lblLabels_3.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me._lblLabels_3.Size = New System.Drawing.Size(46, 26)
      Me._lblLabels_3.TabIndex = 22
      Me._lblLabels_3.Text = "HalfDay/4 Hours"
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
      Me._lblLabels_4.Size = New System.Drawing.Size(32, 13)
      Me._lblLabels_4.TabIndex = 21
      Me._lblLabels_4.Text = "Daily:"
      '
      '_lblLabels_5
      '
      Me._lblLabels_5.AutoSize = True
      Me._lblLabels_5.BackColor = System.Drawing.SystemColors.Control
      Me._lblLabels_5.Cursor = System.Windows.Forms.Cursors.Default
      Me._lblLabels_5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me._lblLabels_5.ForeColor = System.Drawing.SystemColors.ControlText
      Me.lblLabels.SetIndex(Me._lblLabels_5, CType(5, Short))
      Me._lblLabels_5.Location = New System.Drawing.Point(358, 17)
      Me._lblLabels_5.Name = "_lblLabels_5"
      Me._lblLabels_5.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me._lblLabels_5.Size = New System.Drawing.Size(43, 13)
      Me._lblLabels_5.TabIndex = 20
      Me._lblLabels_5.Text = "Weekly:"
      '
      '_lblLabels_6
      '
      Me._lblLabels_6.AutoSize = True
      Me._lblLabels_6.BackColor = System.Drawing.SystemColors.Control
      Me._lblLabels_6.Cursor = System.Windows.Forms.Cursors.Default
      Me._lblLabels_6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me._lblLabels_6.ForeColor = System.Drawing.SystemColors.ControlText
      Me.lblLabels.SetIndex(Me._lblLabels_6, CType(6, Short))
      Me._lblLabels_6.Location = New System.Drawing.Point(358, 39)
      Me._lblLabels_6.Name = "_lblLabels_6"
      Me._lblLabels_6.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me._lblLabels_6.Size = New System.Drawing.Size(46, 13)
      Me._lblLabels_6.TabIndex = 19
      Me._lblLabels_6.Text = "Monthly:"
      '
      '_lblLabels_7
      '
      Me._lblLabels_7.AutoSize = True
      Me._lblLabels_7.BackColor = System.Drawing.SystemColors.Control
      Me._lblLabels_7.Cursor = System.Windows.Forms.Cursors.Default
      Me._lblLabels_7.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me._lblLabels_7.ForeColor = System.Drawing.SystemColors.ControlText
      Me.lblLabels.SetIndex(Me._lblLabels_7, CType(7, Short))
      Me._lblLabels_7.Location = New System.Drawing.Point(348, 60)
      Me._lblLabels_7.Name = "_lblLabels_7"
      Me._lblLabels_7.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me._lblLabels_7.Size = New System.Drawing.Size(55, 13)
      Me._lblLabels_7.TabIndex = 18
      Me._lblLabels_7.Text = "WeekEnd:"
      '
      '_lblLabels_8
      '
      Me._lblLabels_8.AutoSize = True
      Me._lblLabels_8.BackColor = System.Drawing.SystemColors.Control
      Me._lblLabels_8.Cursor = System.Windows.Forms.Cursors.Default
      Me._lblLabels_8.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me._lblLabels_8.ForeColor = System.Drawing.SystemColors.ControlText
      Me.lblLabels.SetIndex(Me._lblLabels_8, CType(8, Short))
      Me._lblLabels_8.Location = New System.Drawing.Point(358, 82)
      Me._lblLabels_8.Name = "_lblLabels_8"
      Me._lblLabels_8.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me._lblLabels_8.Size = New System.Drawing.Size(44, 13)
      Me._lblLabels_8.TabIndex = 17
      Me._lblLabels_8.Text = "Deposit:"
      '
      '_lblLabels_9
      '
      Me._lblLabels_9.AutoSize = True
      Me._lblLabels_9.BackColor = System.Drawing.SystemColors.Control
      Me._lblLabels_9.Cursor = System.Windows.Forms.Cursors.Default
      Me._lblLabels_9.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me._lblLabels_9.ForeColor = System.Drawing.SystemColors.ControlText
      Me.lblLabels.SetIndex(Me._lblLabels_9, CType(9, Short))
      Me._lblLabels_9.Location = New System.Drawing.Point(352, 104)
      Me._lblLabels_9.Name = "_lblLabels_9"
      Me._lblLabels_9.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me._lblLabels_9.Size = New System.Drawing.Size(52, 13)
      Me._lblLabels_9.TabIndex = 16
      Me._lblLabels_9.Text = "Minimum:"
      '
      'dbgRentalRates
      '
      Me.dbgRentalRates.AllowSorting = False
      Me.dbgRentalRates.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                  Or System.Windows.Forms.AnchorStyles.Left) _
                  Or System.Windows.Forms.AnchorStyles.Right)
      Me.dbgRentalRates.DataMember = ""
      Me.dbgRentalRates.HeaderForeColor = System.Drawing.SystemColors.ControlText
      Me.dbgRentalRates.Location = New System.Drawing.Point(11, 0)
      Me.dbgRentalRates.Name = "dbgRentalRates"
      Me.dbgRentalRates.Size = New System.Drawing.Size(596, 218)
      Me.dbgRentalRates.TabIndex = 16
      '
      'frmRentalRates
      '
      Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
      Me.ClientSize = New System.Drawing.Size(616, 369)
      Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.dbgRentalRates, Me.cmdAdd, Me.cmdUpdate, Me.cmdDelete, Me.cmdClose, Me.Frame1})
      Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
      Me.Location = New System.Drawing.Point(22, 139)
      Me.MinimizeBox = False
      Me.Name = "frmRentalRates"
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "Rental Rates"
      Me.Frame1.ResumeLayout(False)
      CType(Me.lblLabels, System.ComponentModel.ISupportInitialize).EndInit()
      CType(Me.dbgRentalRates, System.ComponentModel.ISupportInitialize).EndInit()
      Me.ResumeLayout(False)

   End Sub
#End Region
#Region " Module Variables "
   Dim msAddorEdit As String
   Dim mbDirty As Boolean
   Dim mbFormLoading As Boolean
   Private oDA As CDataAccess
   Private iHitRow As Integer = 0
   Private dtRR As New DataTable("dt")
   Private oCG As New CGrid()


#End Region

#Region " Private Methods "
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


   Private Sub LoadRates()
      Try

         With Me.dtRR.Rows(Me.iHitRow)
            Me.txtEquip_Name.Text = .Item("equip_name")
            Me.txtPrice_ID.Text = .Item("price_id")
            Me.txtHour.Text = FormatCurrency(UnFormat(.Item("hourrate")))
            Me.txtHalfDay.Text = FormatCurrency(UnFormat(.Item("halfday")))
            Me.txtDaily.Text = FormatCurrency(UnFormat(.Item("daily")))
            Me.txtWeekly.Text = FormatCurrency(UnFormat(.Item("weekly")))
            Me.txtMonthly.Text = FormatCurrency(UnFormat(.Item("monthly")))
            Me.txtWeekEnd.Text = FormatCurrency(UnFormat(.Item("WeekEnd")))
            Me.txtDeposit.Text = FormatCurrency(UnFormat(.Item("deposit")))
            Me.txtMinimum.Text = .Item("minimum")
            mbDirty = False
         End With

      Catch ex As System.Exception
         'StructuredErrorHandler(ex)
      End Try

   End Sub

   Private Sub ClearRates()
      With Me
         .txtEquip_Name.Text = ""
         .txtPrice_ID.Text = ""
         .txtHour.Text = ""
         .txtHalfDay.Text = ""
         .txtDaily.Text = ""
         .txtWeekly.Text = ""
         .txtMonthly.Text = ""
         .txtWeekEnd.Text = ""
         .txtDeposit.Text = ""
         .txtMinimum.Text = ""
         mbDirty = False
      End With
   End Sub


   Private Sub LoadTheGrid()
      Dim SQL As String


      Try
         dtRR = New DataTable("dt")
         SQL = "select Equip_Name, "
         SQL = SQL & "Price_Id, "
         SQL = SQL & "HourRate, "
         SQL = SQL & "HalfDay, "
         SQL = SQL & "Daily, "
         SQL = SQL & "weekly, "
         SQL = SQL & "Monthly, "
         SQL = SQL & "WeekEnd, "
         SQL = SQL & "Deposit, "
         SQL = SQL & "Minimum "
         SQL = SQL & "from rental_rates "
         SQL = SQL & " order by equip_name"

         oDA.SendQuery(SQL, dtRR, ConnectString, "dt")
         Me.dbgRentalRates.SetDataBinding(dtRR, "")
         Dim Formats() As String = _
             {"", "100", "T", "L", _
              "", "60", "T", "R", _
              "$#,##0.00", "60", "T", "R", _
              "$#,##0.00", "60", "T", "R", _
              "$#,##0.00", "60", "T", "R", _
              "$#,##0.00", "60", "T", "R", _
              "$#,##0.00", "60", "T", "R", _
              "$#,##0.00", "60", "T", "R", _
              "$#,##0.00", "60", "T", "R", _
              "0", "60", "T", "R"}
         If dtRR.Rows.Count > 0 Then
            oCG.SetTablesStyle(dtRR, Me.dbgRentalRates, Formats)

            Me.dbgRentalRates.SetDataBinding(dtRR, "")
            oCG.DisableAddNew(dbgRentalRates, Me)
         End If
         mbDirty = False
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

#End Region

#Region " Form & Control Events "

   Private Sub cmdAdd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAdd.Click
      Dim SQL As String
      Dim dt As New DataTable()

      If mbDirty Then
         If MsgBox("You have unsaved changes; do you want to add without saving your changes?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
            Exit Sub
         End If
      End If


      ClearRates()
      msAddorEdit = "A"

      SQL = "select max(price_id) from rental_rates"
      oDA.SendQuery(SQL, dt, ConnectString)
      If dt.Rows.Count = 0 Then
         Me.txtPrice_ID.Text = CStr(1)
      Else
         If IsDBNull(dt.Rows(0).Item(0)) Then
            Me.txtPrice_ID.Text = CStr(1)
         Else
            Me.txtPrice_ID.Text = CStr(Val(dt.Rows(0).Item(0)) + 1)
         End If
      End If
      Me.txtEquip_Name.Focus()
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
      Dim iRows As Integer

      Try
         Dim sErr As String = ""


         If MsgBox("Are you sure you want to delete the selected row?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
            Exit Sub
         End If
         SQL = "delete from rental_rates "
         SQL &= "where price_id = " & Me.dtRR.Rows(Me.iHitRow).Item("price_id") & " "
         iRows = oDA.SendActionSql(SQL, ConnectString, sErr)
         If iRows = 0 Then
            MsgBox("Delete of equipment item failed.  " & Chr(10) & sErr, MsgBoxStyle.Critical)
            Exit Sub
         End If
         Me.LoadTheGrid()

      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try

   End Sub

   Private Sub cmdUpdate_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdUpdate.Click
      Dim SQL As String
      Dim sErr As String


      Try
         With Me
            If msAddorEdit = "A" Then
               SQL = "insert into rental_rates "
               SQL = SQL & "(equip_name, price_id, hourrate, halfday, daily, "
               SQL = SQL & "weekly, monthly, weekend, deposit, minimum) "
               SQL = SQL & "values('"
               SQL = SQL & Replace(.txtEquip_Name.Text, "'", "''") & "', "
               SQL = SQL & .txtPrice_ID.Text & ", "
               SQL = SQL & UnFormat(.txtHour.Text) & ", "
               SQL = SQL & UnFormat(.txtHalfDay.Text) & ", "
               SQL = SQL & UnFormat(.txtDaily.Text) & ", "
               SQL = SQL & UnFormat(.txtWeekly.Text) & ", "
               SQL = SQL & UnFormat(.txtMonthly.Text) & ", "
               SQL = SQL & UnFormat(.txtWeekEnd.Text) & ", "
               SQL = SQL & UnFormat(.txtDeposit.Text) & ", "
               SQL = SQL & UnFormat(.txtMinimum.Text) & ") "
            Else
               SQL = "update rental_rates "
               SQL = SQL & "set equip_name = '" & .txtEquip_Name.Text & "', "
               SQL = SQL & "hourrate = " & UnFormat(.txtHour.Text) & ", "
               SQL = SQL & "halfday = " & UnFormat(.txtHalfDay.Text) & ", "
               SQL = SQL & "daily = " & UnFormat(.txtDaily.Text) & ", "
               SQL = SQL & "weekly = " & UnFormat(.txtWeekly.Text) & ", "
               SQL = SQL & "monthly = " & UnFormat(.txtMonthly.Text) & ", "
               SQL = SQL & "weekend = " & UnFormat(.txtWeekEnd.Text) & ", "
               SQL = SQL & "deposit = " & UnFormat(.txtDeposit.Text) & ", "
               SQL = SQL & "minimum = " & UnFormat(.txtMinimum.Text) & " "
               SQL = SQL & "where price_id = " & .txtPrice_ID.Text
            End If
         End With
         oDA.SendActionSql(SQL, ConnectString, sErr)

         msAddorEdit = "E"
         LoadTheGrid()
         mbDirty = False
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub



   Private Sub dbgRentalRates_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
      msAddorEdit = "E"
   End Sub

   Private Sub frmRentalRates_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
      If mbFormLoading Then
         mbFormLoading = False
         mbDirty = False
      End If
   End Sub

   Private Sub frmRentalRates_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
      'IncrementChildCount Me
      LoadTheGrid()
      LoadRates()
      msAddorEdit = "E"
      mbFormLoading = True

   End Sub


   Private Sub txtHour_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtHour.KeyPress
      Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
      If KeyAscii = 13 Then KeyAscii = 0
      KeyAscii = CkKeyPressNumeric(KeyAscii, txtHour)
      If KeyAscii = 0 Then
         eventArgs.Handled = True
      End If
   End Sub
   Private Sub txtHour_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtHour.KeyDown
      Dim KeyCode As Short = eventArgs.KeyCode
      Dim Shift As Short = eventArgs.KeyData \ &H10000
      On Error Resume Next
      If KeyCode = 13 Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
      If KeyCode = System.Windows.Forms.Keys.Up Then System.Windows.Forms.SendKeys.SendWait("+{TAB}")
      If KeyCode = System.Windows.Forms.Keys.Down Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
   End Sub
   Private Sub txtHour_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtHour.Enter
      txtHour.Text = UnFmt_T_B(txtHour)
      txtHour.SelectionStart = 0
      txtHour.SelectionLength = Len(Trim(txtHour.Text))
   End Sub
   Private Sub txtHour_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtHour.Leave
      txtHour.Text = Fmt_T_B(txtHour)
   End Sub

   Private Sub txtPrice_ID_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPrice_ID.KeyPress
      Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
      If KeyAscii = 13 Then KeyAscii = 0
      If KeyAscii = 0 Then
         eventArgs.Handled = True
      End If
   End Sub
   Private Sub txtPrice_ID_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtPrice_ID.KeyDown
      Dim KeyCode As Short = eventArgs.KeyCode
      Dim Shift As Short = eventArgs.KeyData \ &H10000
      On Error Resume Next
      If KeyCode = 13 Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
      If KeyCode = System.Windows.Forms.Keys.Up Then System.Windows.Forms.SendKeys.SendWait("+{TAB}")
      If KeyCode = System.Windows.Forms.Keys.Down Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
   End Sub
   Private Sub txtPrice_ID_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPrice_ID.Enter
      txtPrice_ID.SelectionStart = 0
      txtPrice_ID.SelectionLength = Len(Trim(txtPrice_ID.Text))
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

   Private Sub dbgRentalRates_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dbgRentalRates.MouseUp
      Try
         Dim pt = New Point(e.X, e.Y)
         Dim hti As DataGrid.HitTestInfo = Me.dbgRentalRates.HitTest(pt)
         Me.dbgRentalRates.Select(hti.Row)
         iHitRow = hti.Row
         Me.LoadRates()
         mbDirty = False
         msAddorEdit = "E"
      Catch ex As System.Exception
         'StructuredErrorHandler(ex)
      End Try
   End Sub


   Private Sub txtDaily_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDaily.TextChanged
      mbDirty = True
   End Sub

   Private Sub txtweekend_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWeekEnd.TextChanged
      mbDirty = True
   End Sub

   Private Sub txtDeposit_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDeposit.TextChanged
      mbDirty = True
   End Sub

   Private Sub txtEquip_Name_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEquip_Name.TextChanged
      mbDirty = True
   End Sub

   Private Sub txtHalfDay_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtHalfDay.TextChanged
      mbDirty = True
   End Sub

   Private Sub txtHour_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtHour.TextChanged
      mbDirty = True
   End Sub

   Private Sub txtMinimum_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMinimum.TextChanged
      mbDirty = True
   End Sub

   Private Sub txtMinimum_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtMinimum.KeyPress
      Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
      If KeyAscii = 13 Then KeyAscii = 0
      KeyAscii = CkKeyPressNumeric(KeyAscii, txtMinimum)
      If KeyAscii = 0 Then
         eventArgs.Handled = True
      End If
   End Sub
   Private Sub txtMinimum_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtMinimum.KeyDown
      Dim KeyCode As Short = eventArgs.KeyCode
      Dim Shift As Short = eventArgs.KeyData \ &H10000
      On Error Resume Next
      If KeyCode = 13 Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
      If KeyCode = System.Windows.Forms.Keys.Up Then System.Windows.Forms.SendKeys.SendWait("+{TAB}")
      If KeyCode = System.Windows.Forms.Keys.Down Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
   End Sub
   Private Sub txtMinimum_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMinimum.Enter
      txtMinimum.Text = UnFmt_T_B(txtMinimum)
      txtMinimum.SelectionStart = 0
      txtMinimum.SelectionLength = Len(Trim(txtMinimum.Text))
   End Sub
   Private Sub txtMinimum_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMinimum.Leave
      txtMinimum.Text = Fmt_T_B(txtMinimum)
   End Sub

   Private Sub txtDeposit_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDeposit.KeyPress
      Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
      If KeyAscii = 13 Then KeyAscii = 0
      KeyAscii = CkKeyPressNumeric(KeyAscii, txtDeposit)
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
      txtDeposit.Text = UnFmt_T_B(txtDeposit)
      txtDeposit.SelectionStart = 0
      txtDeposit.SelectionLength = Len(Trim(txtDeposit.Text))
   End Sub
   Private Sub txtDeposit_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDeposit.Leave
      txtDeposit.Text = Fmt_T_B(txtDeposit)
   End Sub

   Private Sub txtweekend_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtWeekEnd.KeyPress
      Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
      If KeyAscii = 13 Then KeyAscii = 0
      KeyAscii = CkKeyPressNumeric(KeyAscii, txtWeekEnd)
      If KeyAscii = 0 Then
         eventArgs.Handled = True
      End If
   End Sub
   Private Sub txtweekend_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtWeekEnd.KeyDown
      Dim KeyCode As Short = eventArgs.KeyCode
      Dim Shift As Short = eventArgs.KeyData \ &H10000
      On Error Resume Next
      If KeyCode = 13 Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
      If KeyCode = System.Windows.Forms.Keys.Up Then System.Windows.Forms.SendKeys.SendWait("+{TAB}")
      If KeyCode = System.Windows.Forms.Keys.Down Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
   End Sub
   Private Sub txtWeekEnd_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWeekEnd.Enter
      txtWeekEnd.Text = UnFmt_T_B(txtWeekEnd)
      txtWeekEnd.SelectionStart = 0
      txtWeekEnd.SelectionLength = Len(Trim(txtWeekEnd.Text))
   End Sub
   Private Sub txtWeekend_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWeekEnd.Leave
      txtWeekEnd.Text = Fmt_T_B(txtWeekEnd)
   End Sub

   Private Sub txtMonthly_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMonthly.TextChanged
      mbDirty = True
   End Sub

   Private Sub txtMonthly_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtMonthly.KeyPress
      Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
      If KeyAscii = 13 Then KeyAscii = 0
      'UPGRADE_ISSUE: Assignment not supported: KeyAscii to a non-zero value Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1058"'
      KeyAscii = CkKeyPressNumeric(KeyAscii, txtMonthly)
      If KeyAscii = 0 Then
         eventArgs.Handled = True
      End If
   End Sub
   Private Sub txtMonthly_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtMonthly.KeyDown
      Dim KeyCode As Short = eventArgs.KeyCode
      Dim Shift As Short = eventArgs.KeyData \ &H10000
      On Error Resume Next
      If KeyCode = 13 Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
      If KeyCode = System.Windows.Forms.Keys.Up Then System.Windows.Forms.SendKeys.SendWait("+{TAB}")
      If KeyCode = System.Windows.Forms.Keys.Down Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
   End Sub
   Private Sub txtMonthly_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMonthly.Enter
      txtMonthly.Text = UnFmt_T_B(txtMonthly)
      txtMonthly.SelectionStart = 0
      txtMonthly.SelectionLength = Len(Trim(txtMonthly.Text))
   End Sub
   Private Sub txtMonthly_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMonthly.Leave
      txtMonthly.Text = Fmt_T_B(txtMonthly)
   End Sub

   Private Sub txtWeekly_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWeekly.TextChanged
      mbDirty = True
   End Sub

   Private Sub txtWeekly_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtWeekly.KeyPress
      Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
      If KeyAscii = 13 Then KeyAscii = 0
      KeyAscii = CkKeyPressNumeric(KeyAscii, txtWeekly)
      If KeyAscii = 0 Then
         eventArgs.Handled = True
      End If
   End Sub
   Private Sub txtWeekly_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtWeekly.KeyDown
      Dim KeyCode As Short = eventArgs.KeyCode
      Dim Shift As Short = eventArgs.KeyData \ &H10000
      On Error Resume Next
      If KeyCode = 13 Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
      If KeyCode = System.Windows.Forms.Keys.Up Then System.Windows.Forms.SendKeys.SendWait("+{TAB}")
      If KeyCode = System.Windows.Forms.Keys.Down Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
   End Sub
   Private Sub txtWeekly_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWeekly.Enter

      txtWeekly.Text = UnFmt_T_B(txtWeekly)
      txtWeekly.SelectionStart = 0
      txtWeekly.SelectionLength = Len(Trim(txtWeekly.Text))
   End Sub
   Private Sub txtWeekly_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWeekly.Leave
      txtWeekly.Text = Fmt_T_B(txtWeekly)
   End Sub

   Private Sub txtDaily_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDaily.KeyPress
      Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
      If KeyAscii = 13 Then KeyAscii = 0
      KeyAscii = CkKeyPressNumeric(KeyAscii, txtDaily)
      If KeyAscii = 0 Then
         eventArgs.Handled = True
      End If
   End Sub
   Private Sub txtDaily_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtDaily.KeyDown
      Dim KeyCode As Short = eventArgs.KeyCode
      Dim Shift As Short = eventArgs.KeyData \ &H10000
      On Error Resume Next
      If KeyCode = 13 Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
      If KeyCode = System.Windows.Forms.Keys.Up Then System.Windows.Forms.SendKeys.SendWait("+{TAB}")
      If KeyCode = System.Windows.Forms.Keys.Down Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
   End Sub
   Private Sub txtDaily_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDaily.Enter
      txtDaily.Text = UnFmt_T_B(txtDaily)
      txtDaily.SelectionStart = 0
      txtDaily.SelectionLength = Len(Trim(txtDaily.Text))
   End Sub
   Private Sub txtDaily_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDaily.Leave
      txtDaily.Text = Fmt_T_B(txtDaily)
   End Sub

   Private Sub txtHalfDay_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtHalfDay.KeyPress
      Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
      If KeyAscii = 13 Then KeyAscii = 0
      KeyAscii = CkKeyPressNumeric(KeyAscii, txtHalfDay)
      If KeyAscii = 0 Then
         eventArgs.Handled = True
      End If
   End Sub
   Private Sub txtHalfDay_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtHalfDay.KeyDown
      Dim KeyCode As Short = eventArgs.KeyCode
      Dim Shift As Short = eventArgs.KeyData \ &H10000
      On Error Resume Next
      If KeyCode = 13 Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
      If KeyCode = System.Windows.Forms.Keys.Up Then System.Windows.Forms.SendKeys.SendWait("+{TAB}")
      If KeyCode = System.Windows.Forms.Keys.Down Then System.Windows.Forms.SendKeys.SendWait("{TAB}")
   End Sub
   Private Sub txtHalfDay_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtHalfDay.Enter
      txtHalfDay.Text = UnFmt_T_B(txtHalfDay)
      txtHalfDay.SelectionStart = 0
      txtHalfDay.SelectionLength = Len(Trim(txtHalfDay.Text))
   End Sub
   Private Sub txtHalfDay_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtHalfDay.Leave
      txtHalfDay.Text = Fmt_T_B(txtHalfDay)
   End Sub


#End Region



End Class