Imports System.Windows.Forms.Application
Public Class frmReRentReport
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
   Friend WithEvents dtpStartDate As System.Windows.Forms.DateTimePicker
   Friend WithEvents dtpEndDate As System.Windows.Forms.DateTimePicker
   Friend WithEvents Label1 As System.Windows.Forms.Label
   Friend WithEvents Label2 As System.Windows.Forms.Label
   Friend WithEvents buttonPreview As System.Windows.Forms.Button
   Friend WithEvents buttonPrint As System.Windows.Forms.Button
   Friend WithEvents buttonCancel As System.Windows.Forms.Button
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
      Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmReRentReport))
      Me.dtpStartDate = New System.Windows.Forms.DateTimePicker()
      Me.dtpEndDate = New System.Windows.Forms.DateTimePicker()
      Me.Label1 = New System.Windows.Forms.Label()
      Me.Label2 = New System.Windows.Forms.Label()
      Me.buttonPreview = New System.Windows.Forms.Button()
      Me.buttonPrint = New System.Windows.Forms.Button()
      Me.buttonCancel = New System.Windows.Forms.Button()
      Me.SuspendLayout()
      '
      'dtpStartDate
      '
      Me.dtpStartDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
      Me.dtpStartDate.Location = New System.Drawing.Point(80, 8)
      Me.dtpStartDate.Name = "dtpStartDate"
      Me.dtpStartDate.Size = New System.Drawing.Size(96, 20)
      Me.dtpStartDate.TabIndex = 0
      '
      'dtpEndDate
      '
      Me.dtpEndDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
      Me.dtpEndDate.Location = New System.Drawing.Point(80, 40)
      Me.dtpEndDate.Name = "dtpEndDate"
      Me.dtpEndDate.Size = New System.Drawing.Size(96, 20)
      Me.dtpEndDate.TabIndex = 1
      '
      'Label1
      '
      Me.Label1.AutoSize = True
      Me.Label1.Location = New System.Drawing.Point(8, 8)
      Me.Label1.Name = "Label1"
      Me.Label1.Size = New System.Drawing.Size(55, 13)
      Me.Label1.TabIndex = 2
      Me.Label1.Text = "Start Date"
      '
      'Label2
      '
      Me.Label2.AutoSize = True
      Me.Label2.Location = New System.Drawing.Point(8, 40)
      Me.Label2.Name = "Label2"
      Me.Label2.Size = New System.Drawing.Size(51, 13)
      Me.Label2.TabIndex = 3
      Me.Label2.Text = "End Date"
      '
      'buttonPreview
      '
      Me.buttonPreview.Location = New System.Drawing.Point(80, 72)
      Me.buttonPreview.Name = "buttonPreview"
      Me.buttonPreview.Size = New System.Drawing.Size(88, 24)
      Me.buttonPreview.TabIndex = 4
      Me.buttonPreview.Text = "&Preview"
      '
      'buttonPrint
      '
      Me.buttonPrint.Location = New System.Drawing.Point(80, 104)
      Me.buttonPrint.Name = "buttonPrint"
      Me.buttonPrint.Size = New System.Drawing.Size(88, 24)
      Me.buttonPrint.TabIndex = 5
      Me.buttonPrint.Text = "P&rint"
      '
      'buttonCancel
      '
      Me.buttonCancel.Location = New System.Drawing.Point(80, 136)
      Me.buttonCancel.Name = "buttonCancel"
      Me.buttonCancel.Size = New System.Drawing.Size(88, 24)
      Me.buttonCancel.TabIndex = 6
      Me.buttonCancel.Text = "&Cancel"
      '
      'frmReRentReport
      '
      Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
      Me.ClientSize = New System.Drawing.Size(192, 174)
      Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.buttonCancel, Me.buttonPrint, Me.buttonPreview, Me.Label2, Me.Label1, Me.dtpEndDate, Me.dtpStartDate})
      Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
      Me.MaximizeBox = False
      Me.MinimizeBox = False
      Me.Name = "frmReRentReport"
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "Print ReRent Report"
      Me.TopMost = True
      Me.ResumeLayout(False)

   End Sub

#End Region

#Region " Form & Control Events "
   Private Sub buttonCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonCancel.Click
      Me.Close()
      DoEvents()
   End Sub

   Private Sub buttonPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonPrint.Click
      Dim oRR As New CPrintReRentReport(False, Me.dtpStartDate.Value, Me.dtpEndDate.Value)
      Me.Close()
      DoEvents()
      oRR.ReRentPrint()
   End Sub

   Private Sub buttonPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonPreview.Click
      Dim oRR As New CPrintReRentReport(True, Me.dtpStartDate.Value, Me.dtpEndDate.Value)
      Me.Close()
      DoEvents()
      oRR.ReRentPrint()
   End Sub


#End Region

End Class
