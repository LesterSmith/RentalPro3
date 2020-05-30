Option Strict Off
Option Explicit On 
Imports VB = Microsoft.VisualBasic
Imports System.Text.RegularExpressions
Friend Class frmCustomerMaintenance
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
    Friend WithEvents dbgCustomers As System.Windows.Forms.DataGrid
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents lblLabels As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents txtBillAdd2 As System.Windows.Forms.TextBox
    Public WithEvents txtBillAdd3 As System.Windows.Forms.TextBox
    Public WithEvents txtCity As System.Windows.Forms.TextBox
    Public WithEvents txtCompanyName As System.Windows.Forms.TextBox
    Public WithEvents txtContactName As System.Windows.Forms.TextBox
    Public WithEvents txtCustomerID As System.Windows.Forms.TextBox
    Public WithEvents txtEmail As System.Windows.Forms.TextBox
    Public WithEvents txtExtension As System.Windows.Forms.TextBox
    Public WithEvents txtFax As System.Windows.Forms.TextBox
    Public WithEvents txtPhone As System.Windows.Forms.TextBox
    Public WithEvents txtPostalCode As System.Windows.Forms.TextBox
    Public WithEvents txtState As System.Windows.Forms.TextBox
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    Friend WithEvents lblTaxid As System.Windows.Forms.Label
    Friend WithEvents txtSecCode As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtCardExpires As System.Windows.Forms.MaskedTextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtCreditCard As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents txtBillAdd1 As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtNote As System.Windows.Forms.TextBox
    Friend WithEvents txtDLNumber As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtTaxID As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCustomerMaintenance))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdUpdate = New System.Windows.Forms.Button()
        Me.cmdClose = New System.Windows.Forms.Button()
        Me.cmdAdd = New System.Windows.Forms.Button()
        Me.cmdDelete = New System.Windows.Forms.Button()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.txtDLNumber = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtNote = New System.Windows.Forms.TextBox()
        Me.txtSecCode = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtCardExpires = New System.Windows.Forms.MaskedTextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtCreditCard = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtTaxID = New System.Windows.Forms.TextBox()
        Me.lblTaxid = New System.Windows.Forms.Label()
        Me.txtEmail = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtFax = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtExtension = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtPhone = New System.Windows.Forms.TextBox()
        Me.txtContactName = New System.Windows.Forms.TextBox()
        Me.txtPostalCode = New System.Windows.Forms.TextBox()
        Me.txtState = New System.Windows.Forms.TextBox()
        Me.txtCity = New System.Windows.Forms.TextBox()
        Me.txtBillAdd3 = New System.Windows.Forms.TextBox()
        Me.txtBillAdd2 = New System.Windows.Forms.TextBox()
        Me.txtBillAdd1 = New System.Windows.Forms.TextBox()
        Me.txtCustomerID = New System.Windows.Forms.TextBox()
        Me.txtCompanyName = New System.Windows.Forms.TextBox()
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
        Me.dbgCustomers = New System.Windows.Forms.DataGrid()
        Me.Frame1.SuspendLayout()
        CType(Me.lblLabels, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dbgCustomers, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdUpdate
        '
        Me.cmdUpdate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdUpdate.BackColor = System.Drawing.SystemColors.Control
        Me.cmdUpdate.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdUpdate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdUpdate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdUpdate.Location = New System.Drawing.Point(549, 314)
        Me.cmdUpdate.Name = "cmdUpdate"
        Me.cmdUpdate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdUpdate.Size = New System.Drawing.Size(77, 26)
        Me.cmdUpdate.TabIndex = 0
        Me.cmdUpdate.Text = "&Save"
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
        Me.cmdClose.Location = New System.Drawing.Point(549, 408)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClose.Size = New System.Drawing.Size(77, 26)
        Me.cmdClose.TabIndex = 3
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
        Me.cmdAdd.Location = New System.Drawing.Point(549, 344)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAdd.Size = New System.Drawing.Size(77, 26)
        Me.cmdAdd.TabIndex = 1
        Me.cmdAdd.Text = "&Add"
        Me.cmdAdd.UseVisualStyleBackColor = False
        '
        'cmdDelete
        '
        Me.cmdDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdDelete.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDelete.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdDelete.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDelete.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDelete.Location = New System.Drawing.Point(549, 376)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDelete.Size = New System.Drawing.Size(77, 26)
        Me.cmdDelete.TabIndex = 2
        Me.cmdDelete.Text = "&Delete"
        Me.cmdDelete.UseVisualStyleBackColor = False
        '
        'Frame1
        '
        Me.Frame1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.txtDLNumber)
        Me.Frame1.Controls.Add(Me.Label8)
        Me.Frame1.Controls.Add(Me.Label7)
        Me.Frame1.Controls.Add(Me.txtNote)
        Me.Frame1.Controls.Add(Me.txtSecCode)
        Me.Frame1.Controls.Add(Me.Label6)
        Me.Frame1.Controls.Add(Me.txtCardExpires)
        Me.Frame1.Controls.Add(Me.Label5)
        Me.Frame1.Controls.Add(Me.txtCreditCard)
        Me.Frame1.Controls.Add(Me.Label4)
        Me.Frame1.Controls.Add(Me.txtTaxID)
        Me.Frame1.Controls.Add(Me.lblTaxid)
        Me.Frame1.Controls.Add(Me.txtEmail)
        Me.Frame1.Controls.Add(Me.Label3)
        Me.Frame1.Controls.Add(Me.txtFax)
        Me.Frame1.Controls.Add(Me.Label2)
        Me.Frame1.Controls.Add(Me.txtExtension)
        Me.Frame1.Controls.Add(Me.Label1)
        Me.Frame1.Controls.Add(Me.txtPhone)
        Me.Frame1.Controls.Add(Me.txtContactName)
        Me.Frame1.Controls.Add(Me.txtPostalCode)
        Me.Frame1.Controls.Add(Me.txtState)
        Me.Frame1.Controls.Add(Me.txtCity)
        Me.Frame1.Controls.Add(Me.txtBillAdd3)
        Me.Frame1.Controls.Add(Me.txtBillAdd2)
        Me.Frame1.Controls.Add(Me.txtBillAdd1)
        Me.Frame1.Controls.Add(Me.txtCustomerID)
        Me.Frame1.Controls.Add(Me.txtCompanyName)
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
        Me.Frame1.Location = New System.Drawing.Point(10, 147)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(528, 290)
        Me.Frame1.TabIndex = 15
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Edit Customer"
        '
        'txtDLNumber
        '
        Me.txtDLNumber.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDLNumber.Location = New System.Drawing.Point(313, 199)
        Me.txtDLNumber.Name = "txtDLNumber"
        Me.txtDLNumber.Size = New System.Drawing.Size(206, 22)
        Me.txtDLNumber.TabIndex = 41
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(277, 199)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(29, 14)
        Me.Label8.TabIndex = 40
        Me.Label8.Text = "DL #"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(52, 222)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(35, 14)
        Me.Label7.TabIndex = 39
        Me.Label7.Text = "Notes"
        '
        'txtNote
        '
        Me.txtNote.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNote.Location = New System.Drawing.Point(89, 225)
        Me.txtNote.Multiline = True
        Me.txtNote.Name = "txtNote"
        Me.txtNote.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtNote.Size = New System.Drawing.Size(430, 58)
        Me.txtNote.TabIndex = 18
        '
        'txtSecCode
        '
        Me.txtSecCode.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSecCode.Location = New System.Drawing.Point(425, 175)
        Me.txtSecCode.MaxLength = 4
        Me.txtSecCode.Name = "txtSecCode"
        Me.txtSecCode.Size = New System.Drawing.Size(37, 22)
        Me.txtSecCode.TabIndex = 38
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(368, 177)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(54, 14)
        Me.Label6.TabIndex = 37
        Me.Label6.Text = "Sec Code"
        '
        'txtCardExpires
        '
        Me.txtCardExpires.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCardExpires.Location = New System.Drawing.Point(315, 175)
        Me.txtCardExpires.Mask = "00/00"
        Me.txtCardExpires.Name = "txtCardExpires"
        Me.txtCardExpires.Size = New System.Drawing.Size(39, 22)
        Me.txtCardExpires.TabIndex = 36
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(260, 177)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(43, 14)
        Me.Label5.TabIndex = 35
        Me.Label5.Text = "Expires"
        '
        'txtCreditCard
        '
        Me.txtCreditCard.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCreditCard.Location = New System.Drawing.Point(90, 174)
        Me.txtCreditCard.MaxLength = 16
        Me.txtCreditCard.Name = "txtCreditCard"
        Me.txtCreditCard.Size = New System.Drawing.Size(163, 22)
        Me.txtCreditCard.TabIndex = 34
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(25, 177)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(61, 14)
        Me.Label4.TabIndex = 33
        Me.Label4.Text = "Credit Card"
        '
        'txtTaxID
        '
        Me.txtTaxID.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTaxID.Location = New System.Drawing.Point(313, 148)
        Me.txtTaxID.Name = "txtTaxID"
        Me.txtTaxID.Size = New System.Drawing.Size(206, 22)
        Me.txtTaxID.TabIndex = 32
        Me.txtTaxID.Tag = "(No Auto Formatting)"
        '
        'lblTaxid
        '
        Me.lblTaxid.AutoSize = True
        Me.lblTaxid.Location = New System.Drawing.Point(267, 152)
        Me.lblTaxid.Name = "lblTaxid"
        Me.lblTaxid.Size = New System.Drawing.Size(36, 14)
        Me.lblTaxid.TabIndex = 31
        Me.lblTaxid.Text = "Tax ID"
        '
        'txtEmail
        '
        Me.txtEmail.AcceptsReturn = True
        Me.txtEmail.BackColor = System.Drawing.SystemColors.Window
        Me.txtEmail.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtEmail.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEmail.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtEmail.Location = New System.Drawing.Point(313, 124)
        Me.txtEmail.MaxLength = 0
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtEmail.Size = New System.Drawing.Size(206, 22)
        Me.txtEmail.TabIndex = 12
        Me.txtEmail.Tag = "(No Auto Formatting)"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(272, 128)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(31, 14)
        Me.Label3.TabIndex = 30
        Me.Label3.Text = "Email"
        '
        'txtFax
        '
        Me.txtFax.AcceptsReturn = True
        Me.txtFax.BackColor = System.Drawing.SystemColors.Window
        Me.txtFax.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFax.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFax.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFax.Location = New System.Drawing.Point(423, 102)
        Me.txtFax.MaxLength = 0
        Me.txtFax.Name = "txtFax"
        Me.txtFax.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFax.Size = New System.Drawing.Size(98, 22)
        Me.txtFax.TabIndex = 11
        Me.txtFax.Tag = "(No Auto Formatting)"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(387, 103)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(28, 14)
        Me.Label2.TabIndex = 29
        Me.Label2.Text = "Fax "
        '
        'txtExtension
        '
        Me.txtExtension.AcceptsReturn = True
        Me.txtExtension.BackColor = System.Drawing.SystemColors.Window
        Me.txtExtension.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtExtension.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtExtension.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtExtension.Location = New System.Drawing.Point(422, 80)
        Me.txtExtension.MaxLength = 0
        Me.txtExtension.Name = "txtExtension"
        Me.txtExtension.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtExtension.Size = New System.Drawing.Size(98, 22)
        Me.txtExtension.TabIndex = 10
        Me.txtExtension.Tag = "(No Auto Formatting)"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(361, 83)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(54, 14)
        Me.Label1.TabIndex = 27
        Me.Label1.Text = "Extension"
        '
        'txtPhone
        '
        Me.txtPhone.AcceptsReturn = True
        Me.txtPhone.BackColor = System.Drawing.SystemColors.Window
        Me.txtPhone.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPhone.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPhone.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPhone.Location = New System.Drawing.Point(422, 58)
        Me.txtPhone.MaxLength = 0
        Me.txtPhone.Name = "txtPhone"
        Me.txtPhone.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPhone.Size = New System.Drawing.Size(98, 22)
        Me.txtPhone.TabIndex = 9
        Me.txtPhone.Tag = "(No Auto Formatting)"
        '
        'txtContactName
        '
        Me.txtContactName.AcceptsReturn = True
        Me.txtContactName.BackColor = System.Drawing.SystemColors.Window
        Me.txtContactName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtContactName.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtContactName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtContactName.Location = New System.Drawing.Point(422, 36)
        Me.txtContactName.MaxLength = 0
        Me.txtContactName.Name = "txtContactName"
        Me.txtContactName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtContactName.Size = New System.Drawing.Size(98, 22)
        Me.txtContactName.TabIndex = 8
        Me.txtContactName.Tag = "(No Auto Formatting)"
        '
        'txtPostalCode
        '
        Me.txtPostalCode.AcceptsReturn = True
        Me.txtPostalCode.BackColor = System.Drawing.SystemColors.Window
        Me.txtPostalCode.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPostalCode.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPostalCode.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPostalCode.Location = New System.Drawing.Point(422, 14)
        Me.txtPostalCode.MaxLength = 0
        Me.txtPostalCode.Name = "txtPostalCode"
        Me.txtPostalCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPostalCode.Size = New System.Drawing.Size(98, 22)
        Me.txtPostalCode.TabIndex = 7
        Me.txtPostalCode.Tag = "(No Auto Formatting)"
        '
        'txtState
        '
        Me.txtState.AcceptsReturn = True
        Me.txtState.BackColor = System.Drawing.SystemColors.Window
        Me.txtState.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtState.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtState.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtState.Location = New System.Drawing.Point(90, 148)
        Me.txtState.MaxLength = 2
        Me.txtState.Name = "txtState"
        Me.txtState.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtState.Size = New System.Drawing.Size(26, 22)
        Me.txtState.TabIndex = 6
        Me.txtState.Tag = "(No Auto Formatting)"
        '
        'txtCity
        '
        Me.txtCity.AcceptsReturn = True
        Me.txtCity.BackColor = System.Drawing.SystemColors.Window
        Me.txtCity.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCity.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCity.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCity.Location = New System.Drawing.Point(89, 125)
        Me.txtCity.MaxLength = 0
        Me.txtCity.Name = "txtCity"
        Me.txtCity.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCity.Size = New System.Drawing.Size(168, 22)
        Me.txtCity.TabIndex = 5
        Me.txtCity.Tag = "(No Auto Formatting)"
        '
        'txtBillAdd3
        '
        Me.txtBillAdd3.AcceptsReturn = True
        Me.txtBillAdd3.BackColor = System.Drawing.SystemColors.Window
        Me.txtBillAdd3.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtBillAdd3.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBillAdd3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtBillAdd3.Location = New System.Drawing.Point(89, 102)
        Me.txtBillAdd3.MaxLength = 0
        Me.txtBillAdd3.Name = "txtBillAdd3"
        Me.txtBillAdd3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtBillAdd3.Size = New System.Drawing.Size(168, 22)
        Me.txtBillAdd3.TabIndex = 4
        Me.txtBillAdd3.Tag = "(No Auto Formatting)"
        '
        'txtBillAdd2
        '
        Me.txtBillAdd2.AcceptsReturn = True
        Me.txtBillAdd2.BackColor = System.Drawing.SystemColors.Window
        Me.txtBillAdd2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtBillAdd2.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBillAdd2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtBillAdd2.Location = New System.Drawing.Point(89, 80)
        Me.txtBillAdd2.MaxLength = 0
        Me.txtBillAdd2.Name = "txtBillAdd2"
        Me.txtBillAdd2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtBillAdd2.Size = New System.Drawing.Size(168, 22)
        Me.txtBillAdd2.TabIndex = 3
        Me.txtBillAdd2.Tag = "(No Auto Formatting)"
        '
        'txtBillAdd1
        '
        Me.txtBillAdd1.AcceptsReturn = True
        Me.txtBillAdd1.BackColor = System.Drawing.SystemColors.Window
        Me.txtBillAdd1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtBillAdd1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBillAdd1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtBillAdd1.Location = New System.Drawing.Point(89, 58)
        Me.txtBillAdd1.MaxLength = 0
        Me.txtBillAdd1.Name = "txtBillAdd1"
        Me.txtBillAdd1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtBillAdd1.Size = New System.Drawing.Size(168, 22)
        Me.txtBillAdd1.TabIndex = 2
        Me.txtBillAdd1.Tag = "(No Auto Formatting)"
        '
        'txtCustomerID
        '
        Me.txtCustomerID.AcceptsReturn = True
        Me.txtCustomerID.BackColor = System.Drawing.Color.White
        Me.txtCustomerID.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCustomerID.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustomerID.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCustomerID.Location = New System.Drawing.Point(89, 36)
        Me.txtCustomerID.MaxLength = 0
        Me.txtCustomerID.Name = "txtCustomerID"
        Me.txtCustomerID.ReadOnly = True
        Me.txtCustomerID.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCustomerID.Size = New System.Drawing.Size(87, 22)
        Me.txtCustomerID.TabIndex = 1
        Me.txtCustomerID.TabStop = False
        Me.txtCustomerID.Tag = "(No Auto Formatting)"
        Me.txtCustomerID.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtCompanyName
        '
        Me.txtCompanyName.AcceptsReturn = True
        Me.txtCompanyName.BackColor = System.Drawing.SystemColors.Window
        Me.txtCompanyName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCompanyName.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCompanyName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCompanyName.Location = New System.Drawing.Point(89, 14)
        Me.txtCompanyName.MaxLength = 0
        Me.txtCompanyName.Name = "txtCompanyName"
        Me.txtCompanyName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCompanyName.Size = New System.Drawing.Size(245, 22)
        Me.txtCompanyName.TabIndex = 0
        Me.txtCompanyName.Tag = "(No Auto Formatting)"
        '
        '_lblLabels_0
        '
        Me._lblLabels_0.AutoSize = True
        Me._lblLabels_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblLabels_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_0, CType(0, Short))
        Me._lblLabels_0.Location = New System.Drawing.Point(4, 17)
        Me._lblLabels_0.Name = "_lblLabels_0"
        Me._lblLabels_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_0.Size = New System.Drawing.Size(82, 14)
        Me._lblLabels_0.TabIndex = 25
        Me._lblLabels_0.Text = "Company Name"
        '
        '_lblLabels_1
        '
        Me._lblLabels_1.AutoSize = True
        Me._lblLabels_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblLabels_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_1, CType(1, Short))
        Me._lblLabels_1.Location = New System.Drawing.Point(23, 39)
        Me._lblLabels_1.Name = "_lblLabels_1"
        Me._lblLabels_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_1.Size = New System.Drawing.Size(62, 14)
        Me._lblLabels_1.TabIndex = 24
        Me._lblLabels_1.Text = "CustomerID"
        '
        '_lblLabels_2
        '
        Me._lblLabels_2.AutoSize = True
        Me._lblLabels_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblLabels_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_2, CType(2, Short))
        Me._lblLabels_2.Location = New System.Drawing.Point(6, 61)
        Me._lblLabels_2.Name = "_lblLabels_2"
        Me._lblLabels_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_2.Size = New System.Drawing.Size(84, 14)
        Me._lblLabels_2.TabIndex = 23
        Me._lblLabels_2.Text = "Billing Address1"
        '
        '_lblLabels_3
        '
        Me._lblLabels_3.AutoSize = True
        Me._lblLabels_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblLabels_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_3, CType(3, Short))
        Me._lblLabels_3.Location = New System.Drawing.Point(1, 84)
        Me._lblLabels_3.Name = "_lblLabels_3"
        Me._lblLabels_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_3.Size = New System.Drawing.Size(84, 14)
        Me._lblLabels_3.TabIndex = 22
        Me._lblLabels_3.Text = "Billing Address2"
        '
        '_lblLabels_4
        '
        Me._lblLabels_4.AutoSize = True
        Me._lblLabels_4.BackColor = System.Drawing.SystemColors.Control
        Me._lblLabels_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_4, CType(4, Short))
        Me._lblLabels_4.Location = New System.Drawing.Point(1, 105)
        Me._lblLabels_4.Name = "_lblLabels_4"
        Me._lblLabels_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_4.Size = New System.Drawing.Size(84, 14)
        Me._lblLabels_4.TabIndex = 21
        Me._lblLabels_4.Text = "Billing Address3"
        '
        '_lblLabels_5
        '
        Me._lblLabels_5.AutoSize = True
        Me._lblLabels_5.BackColor = System.Drawing.SystemColors.Control
        Me._lblLabels_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_5, CType(5, Short))
        Me._lblLabels_5.Location = New System.Drawing.Point(59, 129)
        Me._lblLabels_5.Name = "_lblLabels_5"
        Me._lblLabels_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_5.Size = New System.Drawing.Size(25, 14)
        Me._lblLabels_5.TabIndex = 20
        Me._lblLabels_5.Text = "City"
        '
        '_lblLabels_6
        '
        Me._lblLabels_6.AutoSize = True
        Me._lblLabels_6.BackColor = System.Drawing.SystemColors.Control
        Me._lblLabels_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_6, CType(6, Short))
        Me._lblLabels_6.Location = New System.Drawing.Point(54, 151)
        Me._lblLabels_6.Name = "_lblLabels_6"
        Me._lblLabels_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_6.Size = New System.Drawing.Size(32, 14)
        Me._lblLabels_6.TabIndex = 19
        Me._lblLabels_6.Text = "State"
        '
        '_lblLabels_7
        '
        Me._lblLabels_7.AutoSize = True
        Me._lblLabels_7.BackColor = System.Drawing.SystemColors.Control
        Me._lblLabels_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_7.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_7, CType(7, Short))
        Me._lblLabels_7.Location = New System.Drawing.Point(394, 17)
        Me._lblLabels_7.Name = "_lblLabels_7"
        Me._lblLabels_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_7.Size = New System.Drawing.Size(22, 14)
        Me._lblLabels_7.TabIndex = 18
        Me._lblLabels_7.Text = "Zip"
        '
        '_lblLabels_8
        '
        Me._lblLabels_8.AutoSize = True
        Me._lblLabels_8.BackColor = System.Drawing.SystemColors.Control
        Me._lblLabels_8.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_8.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_8, CType(8, Short))
        Me._lblLabels_8.Location = New System.Drawing.Point(369, 39)
        Me._lblLabels_8.Name = "_lblLabels_8"
        Me._lblLabels_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_8.Size = New System.Drawing.Size(47, 14)
        Me._lblLabels_8.TabIndex = 17
        Me._lblLabels_8.Text = "Contact:"
        '
        '_lblLabels_9
        '
        Me._lblLabels_9.AutoSize = True
        Me._lblLabels_9.BackColor = System.Drawing.SystemColors.Control
        Me._lblLabels_9.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_9.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_9, CType(9, Short))
        Me._lblLabels_9.Location = New System.Drawing.Point(378, 61)
        Me._lblLabels_9.Name = "_lblLabels_9"
        Me._lblLabels_9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_9.Size = New System.Drawing.Size(37, 14)
        Me._lblLabels_9.TabIndex = 16
        Me._lblLabels_9.Text = "Phone"
        '
        'dbgCustomers
        '
        Me.dbgCustomers.AllowSorting = False
        Me.dbgCustomers.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dbgCustomers.CaptionVisible = False
        Me.dbgCustomers.DataMember = ""
        Me.dbgCustomers.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dbgCustomers.Location = New System.Drawing.Point(11, 0)
        Me.dbgCustomers.Name = "dbgCustomers"
        Me.dbgCustomers.Size = New System.Drawing.Size(606, 141)
        Me.dbgCustomers.TabIndex = 16
        '
        'frmCustomerMaintenance
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(633, 449)
        Me.Controls.Add(Me.dbgCustomers)
        Me.Controls.Add(Me.cmdAdd)
        Me.Controls.Add(Me.cmdUpdate)
        Me.Controls.Add(Me.cmdDelete)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.Frame1)
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(22, 139)
        Me.MinimizeBox = False
        Me.Name = "frmCustomerMaintenance"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Customer Data Maintenance"
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        CType(Me.lblLabels, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dbgCustomers, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
#End Region
#Region " Module Variables "
    Dim msAddorEdit As String
    Dim mbDirty As Boolean
    Dim mbFormLoading As Boolean
    Private oDA As CDataAccess
    Private iHitRow As Integer = 0
    Private dtCustomers As New DataTable("dt")
    Private oCG As New CGrid()


#End Region
#Region " Private Methods "
    Private Sub LoadTheGrid()
        Dim SQL As String


        Try
            dtCustomers = New DataTable("dt")
            SQL = "select CompanyName, "
            SQL &= "CustomerID, "
            SQL &= "ContactName, "
            SQL &= "BillingAddress1, "
            SQL &= "BillingAddress2, "
            SQL &= "BillingAddress3, "
            SQL &= "City, "
            SQL &= "State, "
            SQL &= "PostalCode, "
            SQL &= "PhoneNumber, "
            SQL &= "Extension, "
            SQL &= "FaxNumber, "
            SQL &= "EmailAddress,Tax_ID, "
            SQL &= "CreditCard, "
            SQL &= "CardExpires, "
            SQL &= "SecCode, "
            SQL &= "Notes, DLNumber "
            SQL &= "from customers "
            SQL &= "order by companyname"
            Me.dbgCustomers.SetDataBinding(dtCustomers, "")
            Me.dbgCustomers.Refresh()
            Dim iRows = oDA.SendQuery(SQL, dtCustomers, ConnectString, "dt")
            System.Windows.Forms.Application.DoEvents()
            Dim Formats() As String = _
                {"", "150", "T", "L", _
                 "", "60", "T", "R", _
                 "", "150", "T", "L", _
                 "", "150", "T", "L", _
                 "", "150", "T", "L", _
                 "", "60", "T", "L", _
                 "", "40", "T", "L", _
                 "", "40", "T", "L", _
                 "", "60", "T", "L", _
                 "", "60", "T", "L", _
                 "", "60", "T", "L", _
                 "", "60", "T", "L", _
                 "0", "60", "T", "L", _
                 "", "60", "T", "L", _
                 "", "150", "T", "L", _
                 "", "40", "T", "L", _
                 "", "40", "T", "L", _
                 "", "100", "T", "L", _
                 "", "150", "T", "L"}
            If dtCustomers.Rows.Count > 0 Then
                oCG.SetTablesStyle(dtCustomers, Me.dbgCustomers, Formats)

                Me.dbgCustomers.SetDataBinding(dtCustomers, "")
                oCG.DisableAddNew(dbgCustomers, Me)
            End If
            mbDirty = False
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub LoadTextBoxes()
        Try

            With Me.dtCustomers.Rows(Me.iHitRow)
                Me.txtCompanyName.Text = IIf(IsDBNull(.Item("companyname")), "", .Item("companyname"))
                Me.txtCustomerID.Text = IIf(IsDBNull(.Item("customerid")), "", .Item("customerid"))
                Me.txtBillAdd1.Text = IIf(IsDBNull(.Item("billingaddress1")), "", .Item("billingaddress1"))
                Me.txtBillAdd2.Text = IIf(IsDBNull(.Item("billingaddress2")), "", .Item("billingaddress2"))
                Me.txtBillAdd3.Text = IIf(IsDBNull(.Item("billingaddress3")), "", .Item("billingaddress3"))
                Me.txtCity.Text = IIf(IsDBNull(.Item("city")), "", .Item("city"))
                Me.txtContactName.Text = IIf(IsDBNull(.Item("contactname")), "", .Item("contactname"))
                Me.txtEmail.Text = IIf(IsDBNull(.Item("emailaddress")), "", .Item("emailaddress"))
                Me.txtExtension.Text = IIf(IsDBNull(.Item("extension")), "", .Item("extension"))
                Me.txtFax.Text = IIf(IsDBNull(.Item("faxnumber")), "", .Item("faxnumber"))
                Me.txtPhone.Text = IIf(IsDBNull(.Item("phonenumber")), "", .Item("phonenumber"))
                Me.txtPostalCode.Text = IIf(IsDBNull(.Item("postalcode")), "", .Item("postalcode"))
                Me.txtState.Text = IIf(IsDBNull(.Item("state")), "", .Item("state"))
                Try
                    If Not IsDBNull(.Item("tax_id")) AndAlso Not String.IsNullOrEmpty(.Item("tax_id")) Then
                        Me.txtTaxID.Text = StringEncryption.DecryptString(.Item("tax_id"))
                    Else
                        Me.txtTaxID.Text = String.Empty
                    End If
                Catch ex As System.Exception
                    ' if decryption fails the field was probably not encrypted yet
                    Me.txtTaxID.Text = .Item("tax_id")
                End Try
                If IsDBNull(.Item("CreditCard")) OrElse String.IsNullOrEmpty(.Item("CreditCard")) Then
                    Me.txtCreditCard.Text = String.Empty
                Else
                    Me.txtCreditCard.Text = StringEncryption.DecryptString(.Item("CreditCard"))
                End If
                Me.txtCardExpires.Text = IIf(IsDBNull(.Item("CardExpires")), "", .Item("CardExpires"))
                Me.txtSecCode.Text = IIf(IsDBNull(.Item("SecCode")), "", .Item("SecCode"))
                Me.txtNote.Text = IIf(IsDBNull(.Item("Notes")), "", .Item("Notes"))
                If IsDBNull(.Item("DLNumber")) OrElse String.IsNullOrEmpty(.Item("DLNumber")) Then
                    Me.txtDLNumber.Text = String.Empty
                Else
                    Me.txtDLNumber.Text = StringEncryption.DecryptString(.Item("DLNumber"))
                End If
                mbDirty = False
            End With

        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try

    End Sub

    Private Sub ClearTextBoxes()
        With Me
            .txtBillAdd1.Text = String.Empty
            .txtBillAdd2.Text = String.Empty
            .txtBillAdd3.Text = String.Empty
            .txtCity.Text = String.Empty
            .txtCompanyName.Text = String.Empty
            .txtContactName.Text = String.Empty
            .txtCustomerID.Text = String.Empty
            .txtEmail.Text = String.Empty
            .txtExtension.Text = String.Empty
            .txtFax.Text = String.Empty
            .txtPhone.Text = String.Empty
            .txtPostalCode.Text = String.Empty
            .txtState.Text = String.Empty
            .txtTaxID.Text = String.Empty
            .txtCardExpires.Text = String.Empty
            .txtCreditCard.Text = String.Empty
            .txtSecCode.Text = String.Empty
            .txtDLNumber.Text = String.Empty
            .txtNote.Text = String.Empty
            mbDirty = False
        End With
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


#End Region

#Region " Form & Control Events "
    Private Sub cmdAdd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAdd.Click
        Dim SQL As String
        Dim dt As New DataTable()
        Dim oCust As New CCustomer()


        Try
            If mbDirty Then
                If MsgBox("You have unsaved changes; do you want to add without saving your changes?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                    Exit Sub
                End If
            End If


            ClearTextBoxes()
            msAddorEdit = "A"
            Me.txtCustomerID.Text = oCust.GetNewCustID

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
        Dim iRows As Integer

        Try
            Dim sErr As String = ""
            SQL = "select invoiceid from invoices "
            SQL &= "where customerid = " & Me.dtCustomers.Rows(Me.iHitRow).Item("customerid") & " "
            If oDA.SendQuery(SQL, dt, ConnectString) > 0 Then
                MsgBox("You cannot delete a customer that has invoices assoicated with it.", MsgBoxStyle.Exclamation)
                Exit Sub
            End If

            If MsgBox("Are you sure you want to delete the selected row?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Exit Sub
            End If

            dt.Reset()
            SQL = "delete from customers "
            SQL &= "where customerid = " & Me.dtCustomers.Rows(Me.iHitRow).Item("customerid") & " "
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

        If Not ValidateCustData() Then Return

        Try
            With Me
                If msAddorEdit = "A" Then
                    SQL = "insert into customers "
                    SQL &= "(companyname, customerid, Contactname,Billingaddress1, billingaddress2, "
                    SQL &= "billingaddress3, city,state,postalcode,Phonenumber,extension, "
                    SQL &= "FAXnumber,emailaddress,CreditCard,CardExpires,SecCode,Notes,DLNumber,Tax_id) "
                    SQL &= "values('"
                    SQL &= Replace(.txtCompanyName.Text, "'", "''") & "', "
                    SQL &= .txtCustomerID.Text & ", "
                    SQL &= "'" & .txtContactName.Text & "', "
                    SQL &= "'" & .txtBillAdd1.Text & "', "
                    SQL &= "'" & .txtBillAdd2.Text & "', "
                    SQL &= "'" & .txtBillAdd3.Text & "', "
                    SQL &= "'" & .txtCity.Text & "', "
                    SQL &= "'" & .txtState.Text & "', "
                    SQL &= "'" & .txtPostalCode.Text & "', "
                    SQL &= "'" & .txtPhone.Text & "', "
                    SQL &= "'" & .txtExtension.Text & "', "
                    SQL &= "'" & .txtFax.Text & "', "
                    SQL &= "'" & .txtEmail.Text & "', "
                    SQL &= "'" & IIf(.txtCreditCard.Text.Trim() = String.Empty, String.Empty, StringEncryption.EncryptString(.txtCreditCard.Text)) & "', "
                    SQL &= "'" & .txtCardExpires.Text & "', "
                    SQL &= "'" & .txtSecCode.Text & "', "
                    SQL &= "'" & .txtNote.Text & "', "
                    SQL &= "'" & IIf(.txtDLNumber.Text.Trim() = String.Empty, String.Empty, StringEncryption.EncryptString(.txtDLNumber.Text)) & "', "
                    SQL &= "'" & IIf(.txtTaxID.Text.Trim = String.Empty, String.Empty, StringEncryption.EncryptString(.txtTaxID.Text)) & _
                        "') "
                Else
                    SQL = "update customers set "
                    SQL &= "companyname = '" & Replace(.txtCompanyName.Text, "'", "''") & "', "
                    SQL &= "customerid = " & .txtCustomerID.Text & ", "
                    SQL &= "contactname = '" & .txtContactName.Text & "', "
                    SQL &= "billingaddress1 = '" & .txtBillAdd1.Text & "', "
                    SQL &= "billingaddress2 = '" & .txtBillAdd2.Text & "', "
                    SQL &= "billingaddress3 = '" & .txtBillAdd3.Text & "', "
                    SQL &= "city = '" & .txtCity.Text & "', "
                    SQL &= "state = '" & .txtState.Text & "', "
                    SQL &= "postalcode = '" & .txtPostalCode.Text & "', "
                    SQL &= "phonenumber = '" & .txtPhone.Text & "', "
                    SQL &= "extension= '" & .txtExtension.Text & "', "
                    SQL &= "faxnumber = '" & .txtFax.Text & "', "
                    SQL &= "emailaddress = '" & .txtEmail.Text & "', "
                    SQL &= "CreditCard = '" & IIf(.txtCreditCard.Text.Trim = "", "", StringEncryption.EncryptString(.txtCreditCard.Text)) & "', "
                    SQL &= "CardExpires = '" & .txtCardExpires.Text & "', "
                    SQL &= "SecCode = '" & .txtSecCode.Text & "', "
                    SQL &= "Notes = '" & txtNote.Text & "', "
                    SQL &= "DLNumber = '" & IIf(.txtDLNumber.Text.Trim = "", "", StringEncryption.EncryptString(txtDLNumber.Text)) & "', "
                    SQL &= "tax_id = '" & IIf(.txtTaxID.Text.Trim = "", "", StringEncryption.EncryptString(.txtTaxID.Text)) & "' "
                    SQL = SQL & "where customerid = " & .txtCustomerID.Text
                End If
            End With
            If oDA.SendActionSql(SQL, ConnectString, sErr) < 1 Then
                MsgBox("Update of customer data failed.  " & Chr(10) & sErr, MsgBoxStyle.Critical)
                Exit Sub
            End If

            msAddorEdit = "E"
            LoadTheGrid()
            Me.iHitRow = 0
            LoadTextBoxes()
            mbDirty = False
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

    Private Function ValidateCustData() As Boolean
        Dim emailPattern As String = "^(([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)(\s*;\s*|\s*$))*$"
        Dim errors As String = String.Empty
        Dim eol = Environment.NewLine
        Dim re As Regex = New Regex(emailPattern, RegexOptions.IgnoreCase)
        Dim m As Match = re.Match(txtEmail.Text)
        If Not m.Success Then _
            errors &= "The EMail text box contains an invalid Email Address" & eol

        If Not String.IsNullOrEmpty(errors) Then
            MessageBox.Show(errors, "Invalid Data", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End If
        Return True
    End Function

    Private Sub dbgCustomers_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        msAddorEdit = "E"
    End Sub

    Private Sub frmCustomerMaintenance_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        If mbFormLoading Then
            mbFormLoading = False
            mbDirty = False
        End If
    End Sub

    Private Sub frmCustomerMaintenance_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'IncrementChildCount Me
        LoadTheGrid()
        LoadTextBoxes()
        msAddorEdit = "E"
        mbFormLoading = True

    End Sub


    Private Sub dbgCustomers_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dbgCustomers.MouseUp
        Try
            Dim pt = New Point(e.X, e.Y)
            Dim hti As DataGrid.HitTestInfo = Me.dbgCustomers.HitTest(pt)
            Me.dbgCustomers.Select(hti.Row)
            iHitRow = hti.Row
            LoadTextBoxes()
            mbDirty = False
            msAddorEdit = "E"
        Catch ex As System.Exception
            'StructuredErrorHandler(ex)
        End Try
    End Sub
    Private Sub txtCompanyName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCompanyName.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtCompanyName_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCompanyName.KeyDown
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
        If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
        If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
    End Sub
    Private Sub txtCompanyName_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCompanyName.Enter
        txtCompanyName.SelectionStart = 0
        txtCompanyName.SelectionLength = txtCompanyName.Text.Trim.Length
    End Sub
    Private Sub txtExtension_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtExtension.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtExtension_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtExtension.KeyDown
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
        If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
        If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
    End Sub
    Private Sub txtExtension_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtExtension.Enter
        txtExtension.SelectionStart = 0
        txtExtension.SelectionLength = txtExtension.Text.Trim.Length
    End Sub
    Private Sub txtBillAdd2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtBillAdd2.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtBillAdd2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtBillAdd2.KeyDown
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
        If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
        If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
    End Sub
    Private Sub txtBillAdd2_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBillAdd2.Enter
        txtBillAdd2.SelectionStart = 0
        txtBillAdd2.SelectionLength = txtBillAdd2.Text.Trim.Length
    End Sub
    Private Sub txtEmail_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEmail.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtEmail_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtEmail.KeyDown
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
        If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
        If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
    End Sub
    Private Sub txtEmail_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEmail.Enter
        txtEmail.SelectionStart = 0
        txtEmail.SelectionLength = txtEmail.Text.Trim.Length
    End Sub
    Private Sub txtBillAdd1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtBillAdd1.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtBillAdd1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtBillAdd1.KeyDown
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
        If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
        If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
    End Sub
    Private Sub txtBillAdd1_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBillAdd1.Enter
        txtBillAdd1.SelectionStart = 0
        txtBillAdd1.SelectionLength = txtBillAdd1.Text.Trim.Length
    End Sub
    Private Sub txtFax_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFax.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtFax_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFax.KeyDown
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
        If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
        If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
    End Sub
    Private Sub txtFax_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFax.Enter
        txtFax.SelectionStart = 0
        txtFax.SelectionLength = txtFax.Text.Trim.Length
    End Sub
    Private Sub txtCustomerID_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCustomerID.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtCustomerID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustomerID.KeyDown
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
        If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
        If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
    End Sub
    Private Sub txtCustomerID_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerID.Enter
        txtCustomerID.SelectionStart = 0
        txtCustomerID.SelectionLength = txtCustomerID.Text.Trim.Length
    End Sub
    Private Sub txtBillAdd3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtBillAdd3.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtBillAdd3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtBillAdd3.KeyDown
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
        If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
        If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
    End Sub
    Private Sub txtBillAdd3_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBillAdd3.Enter
        txtBillAdd3.SelectionStart = 0
        txtBillAdd3.SelectionLength = txtBillAdd3.Text.Trim.Length
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
    Private Sub txtState_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtState.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtState_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtState.KeyDown
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
        If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
        If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
    End Sub
    Private Sub txtState_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtState.Enter
        txtState.SelectionStart = 0
        txtState.SelectionLength = txtState.Text.Trim.Length
    End Sub
    Private Sub txtPhone_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPhone.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtPhone_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPhone.KeyDown
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
        If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
        If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
    End Sub
    Private Sub txtPhone_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPhone.Enter
        txtPhone.SelectionStart = 0
        txtPhone.SelectionLength = txtPhone.Text.Trim.Length
    End Sub
    Private Sub txtPostalCode_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPostalCode.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtPostalCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPostalCode.KeyDown
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
        If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
        If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
    End Sub
    Private Sub txtPostalCode_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPostalCode.Enter
        txtPostalCode.SelectionStart = 0
        txtPostalCode.SelectionLength = txtPostalCode.Text.Trim.Length
    End Sub
    Private Sub txtCity_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCity.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtCity_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCity.KeyDown
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
        If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
        If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
    End Sub
    Private Sub txtCity_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCity.Enter
        txtCity.SelectionStart = 0
        txtCity.SelectionLength = txtCity.Text.Trim.Length
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


#End Region

End Class
