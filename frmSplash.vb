Option Strict Off
Option Explicit On
Friend Class frmSplash
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
    Public WithEvents picLogo As System.Windows.Forms.PictureBox
    Public WithEvents lblCompanyProduct As System.Windows.Forms.Label
    Public WithEvents lblWarning As System.Windows.Forms.Label
    Public WithEvents fraMainFrame As System.Windows.Forms.GroupBox
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblVersion As System.Windows.Forms.Label
    Friend WithEvents lblDescription As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSplash))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.fraMainFrame = New System.Windows.Forms.GroupBox()
        Me.lblVersion = New System.Windows.Forms.Label()
        Me.lblDescription = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.picLogo = New System.Windows.Forms.PictureBox()
        Me.lblCompanyProduct = New System.Windows.Forms.Label()
        Me.lblWarning = New System.Windows.Forms.Label()
        Me.fraMainFrame.SuspendLayout()
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'fraMainFrame
        '
        Me.fraMainFrame.BackColor = System.Drawing.SystemColors.Control
        Me.fraMainFrame.Controls.Add(Me.lblVersion)
        Me.fraMainFrame.Controls.Add(Me.lblDescription)
        Me.fraMainFrame.Controls.Add(Me.Label1)
        Me.fraMainFrame.Controls.Add(Me.picLogo)
        Me.fraMainFrame.Controls.Add(Me.lblCompanyProduct)
        Me.fraMainFrame.Controls.Add(Me.lblWarning)
        Me.fraMainFrame.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraMainFrame.ForeColor = System.Drawing.Color.Red
        Me.fraMainFrame.Location = New System.Drawing.Point(3, -1)
        Me.fraMainFrame.Name = "fraMainFrame"
        Me.fraMainFrame.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraMainFrame.Size = New System.Drawing.Size(492, 306)
        Me.fraMainFrame.TabIndex = 0
        Me.fraMainFrame.TabStop = False
        '
        'lblVersion
        '
        Me.lblVersion.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVersion.ForeColor = System.Drawing.Color.Black
        Me.lblVersion.Location = New System.Drawing.Point(183, 223)
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.Size = New System.Drawing.Size(97, 16)
        Me.lblVersion.TabIndex = 6
        Me.lblVersion.Text = "Version 4.6"
        '
        'lblDescription
        '
        Me.lblDescription.Font = New System.Drawing.Font("Arial", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDescription.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDescription.Location = New System.Drawing.Point(176, 165)
        Me.lblDescription.Name = "lblDescription"
        Me.lblDescription.Size = New System.Drawing.Size(288, 48)
        Me.lblDescription.TabIndex = 5
        Me.lblDescription.Text = "Label2"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Arial", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(169, 136)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(200, 24)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Enterprise Rental System"
        '
        'picLogo
        '
        Me.picLogo.BackColor = System.Drawing.SystemColors.Control
        Me.picLogo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.picLogo.Cursor = System.Windows.Forms.Cursors.Default
        Me.picLogo.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.picLogo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.picLogo.Image = CType(resources.GetObject("picLogo.Image"), System.Drawing.Image)
        Me.picLogo.Location = New System.Drawing.Point(16, 40)
        Me.picLogo.Name = "picLogo"
        Me.picLogo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.picLogo.Size = New System.Drawing.Size(130, 119)
        Me.picLogo.TabIndex = 1
        Me.picLogo.TabStop = False
        '
        'lblCompanyProduct
        '
        Me.lblCompanyProduct.AutoSize = True
        Me.lblCompanyProduct.BackColor = System.Drawing.SystemColors.Control
        Me.lblCompanyProduct.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCompanyProduct.Font = New System.Drawing.Font("Arial", 24.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompanyProduct.ForeColor = System.Drawing.Color.Blue
        Me.lblCompanyProduct.Location = New System.Drawing.Point(167, 90)
        Me.lblCompanyProduct.Name = "lblCompanyProduct"
        Me.lblCompanyProduct.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCompanyProduct.Size = New System.Drawing.Size(185, 36)
        Me.lblCompanyProduct.TabIndex = 3
        Me.lblCompanyProduct.Tag = "CompanyProduct"
        Me.lblCompanyProduct.Text = "RentalPro3"
        '
        'lblWarning
        '
        Me.lblWarning.BackColor = System.Drawing.SystemColors.Control
        Me.lblWarning.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblWarning.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWarning.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblWarning.Location = New System.Drawing.Point(106, 248)
        Me.lblWarning.Name = "lblWarning"
        Me.lblWarning.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblWarning.Size = New System.Drawing.Size(307, 49)
        Me.lblWarning.TabIndex = 2
        Me.lblWarning.Tag = "Warning"
        Me.lblWarning.Text = "Warning: ...This program is protected by International Copyright Law.  It is the " &
    "property of HHI Software, Inc., Copyright 2001, and is liscensed to "
        '
        'frmSplash
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(497, 314)
        Me.ControlBox = False
        Me.Controls.Add(Me.fraMainFrame)
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(166, 155)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmSplash"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.fraMainFrame.ResumeLayout(False)
        Me.fraMainFrame.PerformLayout()
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
#End Region

#Region " Form & Control Events "
    Private Sub frmSplash_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        'lblVersion.Text = "Version " & App.Major & "." & App.Minor & "." & App.Revision
        '   lblProductName.Caption = App.Title
        Me.lblWarning.Text &= CorporateName
        Me.lblDescription.Text = "Licensed to " & CorporateName
        Me.lblVersion.Text = "Version: " &
                    System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileMajorPart &
                     "." & System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileMinorPart

    End Sub

    Private Sub fraMainFrame_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles fraMainFrame.Enter

    End Sub


#End Region

    Private Sub lblDescription_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblDescription.Click

    End Sub
End Class