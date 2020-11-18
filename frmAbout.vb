Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmAbout
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
    Public WithEvents picIcon As System.Windows.Forms.PictureBox
    Public WithEvents cmdOK As System.Windows.Forms.Button
    Public WithEvents cmdSysInfo As System.Windows.Forms.Button
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents lblDescription As System.Windows.Forms.Label
    Public WithEvents lblTitle As System.Windows.Forms.Label
    Public WithEvents _Line1_1 As System.Windows.Forms.Label
    Public WithEvents _Line1_0 As System.Windows.Forms.Label
    Public WithEvents lblVersion As System.Windows.Forms.Label
    Public WithEvents lblDisclaimer As System.Windows.Forms.Label
    Public WithEvents Line1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    Friend WithEvents HhiImage1 As HHIImage.HHIImage
    Friend WithEvents lblBuildDate As Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAbout))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.picIcon = New System.Windows.Forms.PictureBox()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.cmdSysInfo = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblDescription = New System.Windows.Forms.Label()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me._Line1_1 = New System.Windows.Forms.Label()
        Me._Line1_0 = New System.Windows.Forms.Label()
        Me.lblVersion = New System.Windows.Forms.Label()
        Me.lblDisclaimer = New System.Windows.Forms.Label()
        Me.Line1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.HhiImage1 = New HHIImage.HHIImage()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.lblBuildDate = New System.Windows.Forms.Label()
        CType(Me.picIcon, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Line1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'picIcon
        '
        Me.picIcon.BackColor = System.Drawing.SystemColors.Window
        Me.picIcon.Cursor = System.Windows.Forms.Cursors.Default
        Me.picIcon.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.picIcon.ForeColor = System.Drawing.SystemColors.WindowText
        Me.picIcon.Image = CType(resources.GetObject("picIcon.Image"), System.Drawing.Image)
        Me.picIcon.Location = New System.Drawing.Point(26, 23)
        Me.picIcon.Name = "picIcon"
        Me.picIcon.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.picIcon.Size = New System.Drawing.Size(32, 32)
        Me.picIcon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.picIcon.TabIndex = 2
        Me.picIcon.TabStop = False
        '
        'cmdOK
        '
        Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOK.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdOK.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOK.Location = New System.Drawing.Point(517, 256)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOK.Size = New System.Drawing.Size(157, 33)
        Me.cmdOK.TabIndex = 0
        Me.cmdOK.Tag = "OK"
        Me.cmdOK.Text = "OK"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        'cmdSysInfo
        '
        Me.cmdSysInfo.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSysInfo.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSysInfo.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSysInfo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSysInfo.Location = New System.Drawing.Point(518, 300)
        Me.cmdSysInfo.Name = "cmdSysInfo"
        Me.cmdSysInfo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSysInfo.Size = New System.Drawing.Size(156, 33)
        Me.cmdSysInfo.TabIndex = 1
        Me.cmdSysInfo.Tag = "&System Info..."
        Me.cmdSysInfo.Text = "&System Info..."
        Me.cmdSysInfo.UseVisualStyleBackColor = False
        Me.cmdSysInfo.Visible = False
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 14.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label1.Location = New System.Drawing.Point(112, 175)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(438, 33)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Developed by HHI Software"
        '
        'lblDescription
        '
        Me.lblDescription.BackColor = System.Drawing.SystemColors.Control
        Me.lblDescription.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDescription.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDescription.ForeColor = System.Drawing.Color.Black
        Me.lblDescription.Location = New System.Drawing.Point(112, 110)
        Me.lblDescription.Name = "lblDescription"
        Me.lblDescription.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDescription.Size = New System.Drawing.Size(566, 48)
        Me.lblDescription.TabIndex = 6
        Me.lblDescription.Tag = "App Description"
        Me.lblDescription.Text = "Rental Pro - Enterprise Rental Management System"
        '
        'lblTitle
        '
        Me.lblTitle.BackColor = System.Drawing.SystemColors.Control
        Me.lblTitle.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTitle.Font = New System.Drawing.Font("Arial", 24.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.ForeColor = System.Drawing.Color.Blue
        Me.lblTitle.Location = New System.Drawing.Point(112, 23)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTitle.Size = New System.Drawing.Size(573, 47)
        Me.lblTitle.TabIndex = 5
        Me.lblTitle.Tag = "Application Title"
        Me.lblTitle.Text = "RentalPro"
        '
        '_Line1_1
        '
        Me._Line1_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Line1.SetIndex(Me._Line1_1, CType(1, Short))
        Me._Line1_1.Location = New System.Drawing.Point(24, 237)
        Me._Line1_1.Name = "_Line1_1"
        Me._Line1_1.Size = New System.Drawing.Size(658, 1)
        Me._Line1_1.TabIndex = 8
        '
        '_Line1_0
        '
        Me._Line1_0.BackColor = System.Drawing.Color.White
        Me.Line1.SetIndex(Me._Line1_0, CType(0, Short))
        Me._Line1_0.Location = New System.Drawing.Point(26, 238)
        Me._Line1_0.Name = "_Line1_0"
        Me._Line1_0.Size = New System.Drawing.Size(657, 5)
        Me._Line1_0.TabIndex = 9
        '
        'lblVersion
        '
        Me.lblVersion.BackColor = System.Drawing.SystemColors.Control
        Me.lblVersion.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblVersion.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVersion.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVersion.Location = New System.Drawing.Point(112, 76)
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblVersion.Size = New System.Drawing.Size(142, 22)
        Me.lblVersion.TabIndex = 4
        Me.lblVersion.Tag = "Version"
        Me.lblVersion.Text = "Version 4.5"
        '
        'lblDisclaimer
        '
        Me.lblDisclaimer.BackColor = System.Drawing.SystemColors.Control
        Me.lblDisclaimer.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDisclaimer.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDisclaimer.ForeColor = System.Drawing.Color.Black
        Me.lblDisclaimer.Location = New System.Drawing.Point(27, 256)
        Me.lblDisclaimer.Name = "lblDisclaimer"
        Me.lblDisclaimer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDisclaimer.Size = New System.Drawing.Size(413, 80)
        Me.lblDisclaimer.TabIndex = 3
        Me.lblDisclaimer.Tag = "Warning: ..."
        Me.lblDisclaimer.Text = "Warning: ...This program is protected by International Copyright Law.  It is the " &
    "property of HHI Software, Inc., Copyright 2001-2025, and is liscensed to "
        '
        'HhiImage1
        '
        Me.HhiImage1.BackgroundImage = CType(resources.GetObject("HhiImage1.BackgroundImage"), System.Drawing.Image)
        Me.HhiImage1.Location = New System.Drawing.Point(573, 171)
        Me.HhiImage1.Name = "HhiImage1"
        Me.HhiImage1.Size = New System.Drawing.Size(35, 37)
        Me.HhiImage1.TabIndex = 10
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.SystemColors.ControlText
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox1.Location = New System.Drawing.Point(568, 167)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(47, 48)
        Me.PictureBox1.TabIndex = 11
        Me.PictureBox1.TabStop = False
        '
        'lblBuildDate
        '
        Me.lblBuildDate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBuildDate.Location = New System.Drawing.Point(277, 77)
        Me.lblBuildDate.Name = "lblBuildDate"
        Me.lblBuildDate.Size = New System.Drawing.Size(347, 20)
        Me.lblBuildDate.TabIndex = 12
        Me.lblBuildDate.Text = "Built: "
        '
        'frmAbout
        '
        Me.AcceptButton = Me.cmdOK
        Me.AutoScaleBaseSize = New System.Drawing.Size(8, 19)
        Me.CancelButton = Me.cmdOK
        Me.ClientSize = New System.Drawing.Size(704, 375)
        Me.ControlBox = False
        Me.Controls.Add(Me.lblBuildDate)
        Me.Controls.Add(Me.HhiImage1)
        Me.Controls.Add(Me.picIcon)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.cmdSysInfo)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblDescription)
        Me.Controls.Add(Me.lblTitle)
        Me.Controls.Add(Me._Line1_1)
        Me.Controls.Add(Me._Line1_0)
        Me.Controls.Add(Me.lblVersion)
        Me.Controls.Add(Me.lblDisclaimer)
        Me.Controls.Add(Me.PictureBox1)
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(159, 184)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmAbout"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Tag = "About Project1"
        CType(Me.picIcon, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Line1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region

    ' Reg Key Security Options...
    Const KEY_ALL_ACCESS As Integer = &H2003F


    ' Reg Key ROOT Types...
    Const HKEY_LOCAL_MACHINE As Integer = &H80000002
    Const ERROR_SUCCESS As Short = 0
    Const REG_SZ As Short = 1 ' Unicode nul terminated string
    Const REG_DWORD As Short = 4 ' 32-bit number


    Const gREGKEYSYSINFOLOC As String = "SOFTWARE\Microsoft\Shared Tools Location"
    Const gREGVALSYSINFOLOC As String = "MSINFO"
    Const gREGKEYSYSINFO As String = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
    Const gREGVALSYSINFO As String = "PATH"


    Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Integer, ByVal lpSubKey As String, ByVal ulOptions As Integer, ByVal samDesired As Integer, ByRef phkResult As Integer) As Integer
    Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByVal lpData As String, ByRef lpcbData As Integer) As Integer
    Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Integer) As Integer

    Private Sub frmAbout_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        lblVersion.Text = "Version " &
         System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileMajorPart &
         "." & System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileMinorPart
        'lblTitle.Text = CorporateName 'System.Reflection.Assembly.GetExecutingAssembly.GetName.Name
        Me.lblDescription.Text = "Licensed to " & CorporateName
        Me.lblDisclaimer.Text &= CorporateName
        Me.lblBuildDate.Text = "Built: " + BuildDate
    End Sub



    Private Sub cmdSysInfo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSysInfo.Click
        Call StartSysInfo()
    End Sub


    Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
        Me.Close()
        System.Windows.Forms.Application.DoEvents()
    End Sub


    Public Sub StartSysInfo()
        On Error GoTo SysInfoErr


        Dim rc As Integer
        Dim SysInfoPath As String


        ' Try To Get System Info Program Path\Name From Registry...
        If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
            ' Try To Get System Info Program Path Only From Registry...
        ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
            ' Validate Existance Of Known 32 Bit File Version
            If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
                SysInfoPath = SysInfoPath & "\MSINFO32.EXE"


                ' Error - File Can Not Be Found...
            Else
                GoTo SysInfoErr
            End If
            ' Error - Registry Entry Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If


        Call Shell(SysInfoPath, AppWinStyle.NormalFocus)


        Exit Sub
SysInfoErr:
        MsgBox("System Information Is Unavailable At This Time", MsgBoxStyle.OkOnly)
    End Sub


    Public Function GetKeyValue(ByRef KeyRoot As Integer, ByRef KeyName As String, ByRef SubKeyRef As String, ByRef KeyVal As String) As Boolean
        Dim i As Integer ' Loop Counter
        Dim rc As Integer ' Return Code
        Dim hKey As Integer ' Handle To An Open Registry Key
        Dim hDepth As Integer '
        Dim KeyValType As Integer ' Data Type Of A Registry Key
        Dim tmpVal As String ' Tempory Storage For A Registry Key Value
        Dim KeyValSize As Integer ' Size Of Registry Key Variable
        '------------------------------------------------------------
        ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
        '------------------------------------------------------------
        rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key


        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError ' Handle Error...


        tmpVal = New String(Chr(0), 1024) ' Allocate Variable Space
        KeyValSize = 1024 ' Mark Variable Size


        '------------------------------------------------------------
        ' Retrieve Registry Key Value...
        '------------------------------------------------------------
        rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize) ' Get/Create Key Value


        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError ' Handle Errors


        tmpVal = VB.Left(tmpVal, InStr(tmpVal, Chr(0)) - 1)
        '------------------------------------------------------------
        ' Determine Key Value Type For Conversion...
        '------------------------------------------------------------
        Select Case KeyValType ' Search Data Types...
            Case REG_SZ ' String Registry Key Data Type
                KeyVal = tmpVal ' Copy String Value
            Case REG_DWORD ' Double Word Registry Key Data Type
                For i = Len(tmpVal) To 1 Step -1 ' Convert Each Bit
                    KeyVal = KeyVal & Hex(Asc(Mid(tmpVal, i, 1))) ' Build Value Char. By Char.
                Next
                KeyVal = VB6.Format("&h" & KeyVal) ' Convert Double Word To String
        End Select


        GetKeyValue = True ' Return Success
        rc = RegCloseKey(hKey) ' Close Registry Key
        Exit Function ' Exit


GetKeyError:  ' Cleanup After An Error Has Occured...
        KeyVal = "" ' Set Return Val To Empty String
        GetKeyValue = False ' Return Failure
        rc = RegCloseKey(hKey) ' Close Registry Key
    End Function

    Private Sub HhiImage1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HhiImage1.Load

    End Sub
End Class