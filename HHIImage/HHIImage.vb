Public Class HHIImage
   Inherits System.Windows.Forms.UserControl

#Region " Windows Form Designer generated code "

   Public Sub New()
      MyBase.New()

      'This call is required by the Windows Form Designer.
      InitializeComponent()

      'Add any initialization after the InitializeComponent() call

   End Sub

   'UserControl1 overrides dispose to clean up the component list.
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
   Friend WithEvents Timer1 As System.Windows.Forms.Timer
   Friend WithEvents picSpin0 As System.Windows.Forms.PictureBox
   Friend WithEvents picSpin1 As System.Windows.Forms.PictureBox
   Friend WithEvents picSpin2 As System.Windows.Forms.PictureBox
   Friend WithEvents picSpin3 As System.Windows.Forms.PictureBox
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
      Me.components = New System.ComponentModel.Container()
      Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(HHIImage))
      Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
      Me.picSpin0 = New System.Windows.Forms.PictureBox()
      Me.picSpin1 = New System.Windows.Forms.PictureBox()
      Me.picSpin2 = New System.Windows.Forms.PictureBox()
      Me.picSpin3 = New System.Windows.Forms.PictureBox()
      Me.SuspendLayout()
      '
      'Timer1
      '
      Me.Timer1.Enabled = True
      Me.Timer1.Interval = 500
      '
      'picSpin0
      '
      Me.picSpin0.Image = CType(resources.GetObject("picSpin0.Image"), System.Drawing.Bitmap)
      Me.picSpin0.Name = "picSpin0"
      Me.picSpin0.Size = New System.Drawing.Size(32, 32)
      Me.picSpin0.TabIndex = 0
      Me.picSpin0.TabStop = False
      '
      'picSpin1
      '
      Me.picSpin1.Image = CType(resources.GetObject("picSpin1.Image"), System.Drawing.Bitmap)
      Me.picSpin1.Name = "picSpin1"
      Me.picSpin1.Size = New System.Drawing.Size(32, 32)
      Me.picSpin1.TabIndex = 1
      Me.picSpin1.TabStop = False
      '
      'picSpin2
      '
      Me.picSpin2.Image = CType(resources.GetObject("picSpin2.Image"), System.Drawing.Bitmap)
      Me.picSpin2.Name = "picSpin2"
      Me.picSpin2.Size = New System.Drawing.Size(32, 32)
      Me.picSpin2.TabIndex = 2
      Me.picSpin2.TabStop = False
      '
      'picSpin3
      '
      Me.picSpin3.Image = CType(resources.GetObject("picSpin3.Image"), System.Drawing.Bitmap)
      Me.picSpin3.Name = "picSpin3"
      Me.picSpin3.Size = New System.Drawing.Size(32, 32)
      Me.picSpin3.TabIndex = 3
      Me.picSpin3.TabStop = False
      '
      'HHIImage
      '
      Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Bitmap)
      Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.picSpin1, Me.picSpin2, Me.picSpin3, Me.picSpin0})
      Me.Name = "HHIImage"
      Me.Size = New System.Drawing.Size(32, 32)
      Me.ResumeLayout(False)

   End Sub

#End Region

   Private Sub Timer1_Tick(ByVal sender As Object, _
      ByVal e As System.EventArgs) Handles Timer1.Tick
      Select Case True
         Case Me.picSpin0.Visible
            picSpin1.Visible = True
            picSpin2.Visible = False
            picSpin3.Visible = False
            picSpin0.Visible = False
         Case picSpin1.Visible
            picSpin1.Visible = False
            picSpin2.Visible = True
            picSpin3.Visible = False
            picSpin0.Visible = False
         Case picSpin2.Visible
            picSpin1.Visible = False
            picSpin2.Visible = False
            picSpin3.Visible = True
            picSpin0.Visible = False
         Case picSpin3.Visible
            picSpin1.Visible = False
            picSpin2.Visible = False
            picSpin3.Visible = False
            picSpin0.Visible = True
      End Select
   End Sub
End Class
