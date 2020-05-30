Public Class frmDeliveryReport
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
   Friend WithEvents lblEndDate As System.Windows.Forms.Label
   Friend WithEvents lblStDate As System.Windows.Forms.Label
   Friend WithEvents dtpEndDate As System.Windows.Forms.DateTimePicker
   Friend WithEvents dtpStDate As System.Windows.Forms.DateTimePicker
   Friend WithEvents btnPreview As System.Windows.Forms.Button
   Friend WithEvents btnPrint As System.Windows.Forms.Button
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
      Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmDeliveryReport))
      Me.lblEndDate = New System.Windows.Forms.Label()
      Me.lblStDate = New System.Windows.Forms.Label()
      Me.dtpEndDate = New System.Windows.Forms.DateTimePicker()
      Me.dtpStDate = New System.Windows.Forms.DateTimePicker()
      Me.btnPreview = New System.Windows.Forms.Button()
      Me.btnPrint = New System.Windows.Forms.Button()
      Me.SuspendLayout()
      '
      'lblEndDate
      '
      Me.lblEndDate.AutoSize = True
      Me.lblEndDate.Location = New System.Drawing.Point(18, 40)
      Me.lblEndDate.Name = "lblEndDate"
      Me.lblEndDate.Size = New System.Drawing.Size(51, 13)
      Me.lblEndDate.TabIndex = 17
      Me.lblEndDate.Text = "End Date"
      '
      'lblStDate
      '
      Me.lblStDate.AutoSize = True
      Me.lblStDate.Location = New System.Drawing.Point(18, 18)
      Me.lblStDate.Name = "lblStDate"
      Me.lblStDate.Size = New System.Drawing.Size(55, 13)
      Me.lblStDate.TabIndex = 16
      Me.lblStDate.Text = "Start Date"
      '
      'dtpEndDate
      '
      Me.dtpEndDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
      Me.dtpEndDate.Location = New System.Drawing.Point(74, 40)
      Me.dtpEndDate.Name = "dtpEndDate"
      Me.dtpEndDate.Size = New System.Drawing.Size(96, 20)
      Me.dtpEndDate.TabIndex = 15
      '
      'dtpStDate
      '
      Me.dtpStDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
      Me.dtpStDate.Location = New System.Drawing.Point(74, 16)
      Me.dtpStDate.Name = "dtpStDate"
      Me.dtpStDate.Size = New System.Drawing.Size(96, 20)
      Me.dtpStDate.TabIndex = 14
      '
      'btnPreview
      '
      Me.btnPreview.Location = New System.Drawing.Point(34, 74)
      Me.btnPreview.Name = "btnPreview"
      Me.btnPreview.Size = New System.Drawing.Size(64, 28)
      Me.btnPreview.TabIndex = 19
      Me.btnPreview.Text = "&Preview"
      '
      'btnPrint
      '
      Me.btnPrint.Location = New System.Drawing.Point(106, 74)
      Me.btnPrint.Name = "btnPrint"
      Me.btnPrint.Size = New System.Drawing.Size(64, 28)
      Me.btnPrint.TabIndex = 18
      Me.btnPrint.Text = "&Print"
      '
      'frmDeliveryReport
      '
      Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
      Me.ClientSize = New System.Drawing.Size(202, 120)
      Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnPreview, Me.btnPrint, Me.lblEndDate, Me.lblStDate, Me.dtpEndDate, Me.dtpStDate})
      Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
      Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
      Me.MaximizeBox = False
      Me.MinimizeBox = False
      Me.Name = "frmDeliveryReport"
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "Delivery & Pickup Report"
      Me.ResumeLayout(False)

   End Sub

#End Region
   Private m_AcctBasis As String


#Region "Form and Control Events"
   Private Sub frmDeliveryReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
      ' set start date and end date default

   End Sub

   Private Sub btnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPreview.Click
      Dim oPS As New CDeliveryReport()
      oPS.Preview = True
      oPS.AcctBasis = m_AcctBasis
      oPS.StDate = Me.dtpStDate.Value.Date
      oPS.EndDate = Me.dtpEndDate.Value.Date
      oPS.PrintDeliveryReport()
   End Sub

   Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
      Dim oPS As New CDeliveryReport()
      oPS.AcctBasis = m_AcctBasis
      oPS.Preview = False
      oPS.StDate = Me.dtpStDate.Value.Date
      oPS.EndDate = Me.dtpEndDate.Value.Date
      oPS.PrintDeliveryReport()
   End Sub
#End Region

#Region "Property Methods"

   Public Property AcctBasis() As String
      Get
         Return m_AcctBasis
      End Get
      Set(ByVal Value As String)
         m_AcctBasis = Value
      End Set
   End Property
#End Region
End Class
