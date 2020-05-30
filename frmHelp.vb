Public Class frmHelp
   Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

   Public Sub New()
      MyBase.New()

      'This call is required by the Windows Form Designer.
      InitializeComponent()

      'Add any initialization after the InitializeComponent() call

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
   Friend WithEvents tbHelpInfo As System.Windows.Forms.TextBox
   Friend WithEvents btnClose As System.Windows.Forms.Button
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
      Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmHelp))
      Me.tbHelpInfo = New System.Windows.Forms.TextBox()
      Me.btnClose = New System.Windows.Forms.Button()
      Me.SuspendLayout()
      '
      'tbHelpInfo
      '
      Me.tbHelpInfo.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                  Or System.Windows.Forms.AnchorStyles.Left) _
                  Or System.Windows.Forms.AnchorStyles.Right)
      Me.tbHelpInfo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.tbHelpInfo.Multiline = True
      Me.tbHelpInfo.Name = "tbHelpInfo"
      Me.tbHelpInfo.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
      Me.tbHelpInfo.Size = New System.Drawing.Size(520, 264)
      Me.tbHelpInfo.TabIndex = 3
      Me.tbHelpInfo.Text = "To get help on a topic, click the desired Link Button below."
      '
      'btnClose
      '
      Me.btnClose.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
      Me.btnClose.Location = New System.Drawing.Point(432, 280)
      Me.btnClose.Name = "btnClose"
      Me.btnClose.Size = New System.Drawing.Size(64, 24)
      Me.btnClose.TabIndex = 5
      Me.btnClose.Text = "&Close"
      '
      'frmHelp
      '
      Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
      Me.ClientSize = New System.Drawing.Size(520, 317)
      Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnClose, Me.tbHelpInfo})
      Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
      Me.Name = "frmHelp"
      Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "On Line Help System"
      Me.TopMost = True
      Me.ResumeLayout(False)

   End Sub

#End Region
   ' if this property is populated, display the
   ' contents and hide the other help buttons.
   Public CannedMessage As String = ""

    'test comment


#Region " Form & Control Events "
   Private Sub frmHelp_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
      With Me
         If .CannedMessage.Length > 0 Then
            Me.tbHelpInfo.Text = .CannedMessage
         End If
      End With
   End Sub

   Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
      Me.Close()
      System.Windows.Forms.Application.DoEvents()
   End Sub

   Private Sub tbHelpInfo_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbHelpInfo.Enter
      With tbHelpInfo
         .SelectionLength = 0
      End With
   End Sub


#End Region


End Class
