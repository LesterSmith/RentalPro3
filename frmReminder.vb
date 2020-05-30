Imports System.Windows.Forms.Application

Public Class frmReminder
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
   Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
   Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
   Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
   Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
   Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
   Friend WithEvents dgDueIn As System.Windows.Forms.DataGrid
   Friend WithEvents dgReserved As System.Windows.Forms.DataGrid
   Friend WithEvents dgOver17Days As System.Windows.Forms.DataGrid
   Friend WithEvents dgOver3Days As System.Windows.Forms.DataGrid
   Friend WithEvents buttonCancel As System.Windows.Forms.Button
   Friend WithEvents buttonPreview As System.Windows.Forms.Button
   Friend WithEvents buttonPrintSetup As System.Windows.Forms.Button
   Friend WithEvents PrintDialog1 As System.Windows.Forms.PrintDialog
   Friend WithEvents PageSetupDialog1 As System.Windows.Forms.PageSetupDialog
   Friend WithEvents TabPage5 As System.Windows.Forms.TabPage
   Friend WithEvents dgRerentsDue As System.Windows.Forms.DataGrid
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
      Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmReminder))
      Me.TabControl1 = New System.Windows.Forms.TabControl()
      Me.TabPage1 = New System.Windows.Forms.TabPage()
      Me.dgDueIn = New System.Windows.Forms.DataGrid()
      Me.TabPage2 = New System.Windows.Forms.TabPage()
      Me.dgReserved = New System.Windows.Forms.DataGrid()
      Me.TabPage3 = New System.Windows.Forms.TabPage()
      Me.dgOver17Days = New System.Windows.Forms.DataGrid()
      Me.TabPage4 = New System.Windows.Forms.TabPage()
      Me.dgOver3Days = New System.Windows.Forms.DataGrid()
      Me.buttonCancel = New System.Windows.Forms.Button()
      Me.buttonPreview = New System.Windows.Forms.Button()
      Me.buttonPrintSetup = New System.Windows.Forms.Button()
      Me.PrintDialog1 = New System.Windows.Forms.PrintDialog()
      Me.PageSetupDialog1 = New System.Windows.Forms.PageSetupDialog()
      Me.TabPage5 = New System.Windows.Forms.TabPage()
      Me.dgRerentsDue = New System.Windows.Forms.DataGrid()
      Me.TabControl1.SuspendLayout()
      Me.TabPage1.SuspendLayout()
      CType(Me.dgDueIn, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.TabPage2.SuspendLayout()
      CType(Me.dgReserved, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.TabPage3.SuspendLayout()
      CType(Me.dgOver17Days, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.TabPage4.SuspendLayout()
      CType(Me.dgOver3Days, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.TabPage5.SuspendLayout()
      CType(Me.dgRerentsDue, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.SuspendLayout()
      '
      'TabControl1
      '
      Me.TabControl1.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                  Or System.Windows.Forms.AnchorStyles.Left) _
                  Or System.Windows.Forms.AnchorStyles.Right)
      Me.TabControl1.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabPage1, Me.TabPage2, Me.TabPage3, Me.TabPage4, Me.TabPage5})
      Me.TabControl1.Name = "TabControl1"
      Me.TabControl1.SelectedIndex = 0
      Me.TabControl1.Size = New System.Drawing.Size(752, 312)
      Me.TabControl1.TabIndex = 0
      '
      'TabPage1
      '
      Me.TabPage1.Controls.AddRange(New System.Windows.Forms.Control() {Me.dgDueIn})
      Me.TabPage1.Location = New System.Drawing.Point(4, 22)
      Me.TabPage1.Name = "TabPage1"
      Me.TabPage1.Size = New System.Drawing.Size(744, 286)
      Me.TabPage1.TabIndex = 0
      Me.TabPage1.Text = "Items Due Back Today"
      '
      'dgDueIn
      '
      Me.dgDueIn.CaptionText = "Equipment Due Back Today"
      Me.dgDueIn.DataMember = ""
      Me.dgDueIn.Dock = System.Windows.Forms.DockStyle.Fill
      Me.dgDueIn.HeaderForeColor = System.Drawing.SystemColors.ControlText
      Me.dgDueIn.Name = "dgDueIn"
      Me.dgDueIn.Size = New System.Drawing.Size(744, 286)
      Me.dgDueIn.TabIndex = 0
      '
      'TabPage2
      '
      Me.TabPage2.Controls.AddRange(New System.Windows.Forms.Control() {Me.dgReserved})
      Me.TabPage2.Location = New System.Drawing.Point(4, 22)
      Me.TabPage2.Name = "TabPage2"
      Me.TabPage2.Size = New System.Drawing.Size(744, 286)
      Me.TabPage2.TabIndex = 1
      Me.TabPage2.Text = "Items Reserved Today"
      '
      'dgReserved
      '
      Me.dgReserved.CaptionText = "Equipment Reserved for Today"
      Me.dgReserved.DataMember = ""
      Me.dgReserved.Dock = System.Windows.Forms.DockStyle.Fill
      Me.dgReserved.HeaderForeColor = System.Drawing.SystemColors.ControlText
      Me.dgReserved.Name = "dgReserved"
      Me.dgReserved.Size = New System.Drawing.Size(744, 286)
      Me.dgReserved.TabIndex = 1
      '
      'TabPage3
      '
      Me.TabPage3.Controls.AddRange(New System.Windows.Forms.Control() {Me.dgOver17Days})
      Me.TabPage3.Location = New System.Drawing.Point(4, 22)
      Me.TabPage3.Name = "TabPage3"
      Me.TabPage3.Size = New System.Drawing.Size(744, 286)
      Me.TabPage3.TabIndex = 2
      Me.TabPage3.Text = "Rented Over 17 Days"
      '
      'dgOver17Days
      '
      Me.dgOver17Days.CaptionText = "Eauipment Rented Past Monthly Break"
      Me.dgOver17Days.DataMember = ""
      Me.dgOver17Days.Dock = System.Windows.Forms.DockStyle.Fill
      Me.dgOver17Days.HeaderForeColor = System.Drawing.SystemColors.ControlText
      Me.dgOver17Days.Name = "dgOver17Days"
      Me.dgOver17Days.Size = New System.Drawing.Size(744, 286)
      Me.dgOver17Days.TabIndex = 1
      '
      'TabPage4
      '
      Me.TabPage4.Controls.AddRange(New System.Windows.Forms.Control() {Me.dgOver3Days})
      Me.TabPage4.Location = New System.Drawing.Point(4, 22)
      Me.TabPage4.Name = "TabPage4"
      Me.TabPage4.Size = New System.Drawing.Size(744, 286)
      Me.TabPage4.TabIndex = 3
      Me.TabPage4.Text = "Rented Over 3 Days"
      '
      'dgOver3Days
      '
      Me.dgOver3Days.CaptionText = "Equipment Rented Past Weekly Break"
      Me.dgOver3Days.DataMember = ""
      Me.dgOver3Days.Dock = System.Windows.Forms.DockStyle.Fill
      Me.dgOver3Days.HeaderForeColor = System.Drawing.SystemColors.ControlText
      Me.dgOver3Days.Name = "dgOver3Days"
      Me.dgOver3Days.Size = New System.Drawing.Size(744, 286)
      Me.dgOver3Days.TabIndex = 1
      '
      'buttonCancel
      '
      Me.buttonCancel.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
      Me.buttonCancel.Location = New System.Drawing.Point(656, 320)
      Me.buttonCancel.Name = "buttonCancel"
      Me.buttonCancel.Size = New System.Drawing.Size(72, 24)
      Me.buttonCancel.TabIndex = 1
      Me.buttonCancel.Text = "&Close"
      '
      'buttonPreview
      '
      Me.buttonPreview.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
      Me.buttonPreview.Location = New System.Drawing.Point(552, 320)
      Me.buttonPreview.Name = "buttonPreview"
      Me.buttonPreview.Size = New System.Drawing.Size(72, 24)
      Me.buttonPreview.TabIndex = 2
      Me.buttonPreview.Text = "&Preview"
      Me.buttonPreview.Visible = False
      '
      'buttonPrintSetup
      '
      Me.buttonPrintSetup.Location = New System.Drawing.Point(8, 320)
      Me.buttonPrintSetup.Name = "buttonPrintSetup"
      Me.buttonPrintSetup.Size = New System.Drawing.Size(72, 24)
      Me.buttonPrintSetup.TabIndex = 3
      Me.buttonPrintSetup.Text = "Print Setup"
      Me.buttonPrintSetup.Visible = False
      '
      'TabPage5
      '
      Me.TabPage5.Controls.AddRange(New System.Windows.Forms.Control() {Me.dgRerentsDue})
      Me.TabPage5.Location = New System.Drawing.Point(4, 22)
      Me.TabPage5.Name = "TabPage5"
      Me.TabPage5.Size = New System.Drawing.Size(744, 286)
      Me.TabPage5.TabIndex = 4
      Me.TabPage5.Text = "ReRents Due To Rent"
      '
      'dgRerentsDue
      '
      Me.dgRerentsDue.DataMember = ""
      Me.dgRerentsDue.Dock = System.Windows.Forms.DockStyle.Fill
      Me.dgRerentsDue.HeaderForeColor = System.Drawing.SystemColors.ControlText
      Me.dgRerentsDue.Name = "dgRerentsDue"
      Me.dgRerentsDue.Size = New System.Drawing.Size(744, 286)
      Me.dgRerentsDue.TabIndex = 0
      '
      'frmReminder
      '
      Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
      Me.ClientSize = New System.Drawing.Size(752, 350)
      Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.buttonPrintSetup, Me.buttonPreview, Me.buttonCancel, Me.TabControl1})
      Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
      Me.Name = "frmReminder"
      Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "Daily Reminder of Equipment"
      Me.TabControl1.ResumeLayout(False)
      Me.TabPage1.ResumeLayout(False)
      CType(Me.dgDueIn, System.ComponentModel.ISupportInitialize).EndInit()
      Me.TabPage2.ResumeLayout(False)
      CType(Me.dgReserved, System.ComponentModel.ISupportInitialize).EndInit()
      Me.TabPage3.ResumeLayout(False)
      CType(Me.dgOver17Days, System.ComponentModel.ISupportInitialize).EndInit()
      Me.TabPage4.ResumeLayout(False)
      CType(Me.dgOver3Days, System.ComponentModel.ISupportInitialize).EndInit()
      Me.TabPage5.ResumeLayout(False)
      CType(Me.dgRerentsDue, System.ComponentModel.ISupportInitialize).EndInit()
      Me.ResumeLayout(False)

   End Sub

#End Region
#Region " Module Variables "
   Private oCR As New CReminder()
   Private formLoading As Boolean


#End Region

#Region " Form & Control Events "
   Private Sub buttonPrintSetup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonPrintSetup.Click

      Try
         With Me.PageSetupDialog1
            Dim pd As New Printing.PrintDocument()
            .Document = pd
            .ShowDialog()
         End With
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub



   Private Sub frmReminder_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
      If formLoading Then
         formLoading = False
         oCR.GetEquipDueInToday(Me.dgDueIn, Me)
         oCR.GetEquipReservedToday(Me.dgReserved, Me)
         oCR.GetInvoicesOver17Days(Me.dgOver17Days, Me)
         oCR.GetInvoicesOver3Days(Me.dgOver3Days, Me)
         oCR.GetRerentsDue(Me.dgRerentsDue, Me)
      End If
   End Sub

   Private Sub frmReminder_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
      formLoading = True
   End Sub

   Private Sub buttonCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonCancel.Click
      Me.Close()
      DoEvents()
   End Sub


#End Region

End Class
