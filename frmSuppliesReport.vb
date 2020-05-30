Imports System.Windows.Forms.Application
Imports System.Text

Public Class frmSuppliesReport
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
   Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
   Friend WithEvents optEquipNameSort As System.Windows.Forms.RadioButton
   Friend WithEvents optEquipIdSort As System.Windows.Forms.RadioButton
   Friend WithEvents buttonCancel As System.Windows.Forms.Button
   Friend WithEvents buttonPrint As System.Windows.Forms.Button
   Friend WithEvents buttonPreview As System.Windows.Forms.Button
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
      Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmSuppliesReport))
      Me.GroupBox2 = New System.Windows.Forms.GroupBox()
      Me.optEquipNameSort = New System.Windows.Forms.RadioButton()
      Me.optEquipIdSort = New System.Windows.Forms.RadioButton()
      Me.buttonCancel = New System.Windows.Forms.Button()
      Me.buttonPrint = New System.Windows.Forms.Button()
      Me.buttonPreview = New System.Windows.Forms.Button()
      Me.GroupBox2.SuspendLayout()
      Me.SuspendLayout()
      '
      'GroupBox2
      '
      Me.GroupBox2.Controls.AddRange(New System.Windows.Forms.Control() {Me.optEquipNameSort, Me.optEquipIdSort})
      Me.GroupBox2.Location = New System.Drawing.Point(17, 10)
      Me.GroupBox2.Name = "GroupBox2"
      Me.GroupBox2.Size = New System.Drawing.Size(231, 44)
      Me.GroupBox2.TabIndex = 15
      Me.GroupBox2.TabStop = False
      Me.GroupBox2.Text = "Report Sort Order"
      '
      'optEquipNameSort
      '
      Me.optEquipNameSort.Location = New System.Drawing.Point(120, 18)
      Me.optEquipNameSort.Name = "optEquipNameSort"
      Me.optEquipNameSort.Size = New System.Drawing.Size(94, 16)
      Me.optEquipNameSort.TabIndex = 1
      Me.optEquipNameSort.Text = "Product Name"
      '
      'optEquipIdSort
      '
      Me.optEquipIdSort.Checked = True
      Me.optEquipIdSort.Location = New System.Drawing.Point(9, 18)
      Me.optEquipIdSort.Name = "optEquipIdSort"
      Me.optEquipIdSort.Size = New System.Drawing.Size(99, 19)
      Me.optEquipIdSort.TabIndex = 0
      Me.optEquipIdSort.TabStop = True
      Me.optEquipIdSort.Text = "Product ID"
      '
      'buttonCancel
      '
      Me.buttonCancel.Location = New System.Drawing.Point(195, 80)
      Me.buttonCancel.Name = "buttonCancel"
      Me.buttonCancel.Size = New System.Drawing.Size(74, 24)
      Me.buttonCancel.TabIndex = 14
      Me.buttonCancel.Text = "&Cancel"
      '
      'buttonPrint
      '
      Me.buttonPrint.Location = New System.Drawing.Point(103, 80)
      Me.buttonPrint.Name = "buttonPrint"
      Me.buttonPrint.Size = New System.Drawing.Size(74, 24)
      Me.buttonPrint.TabIndex = 13
      Me.buttonPrint.Text = "P&rint"
      '
      'buttonPreview
      '
      Me.buttonPreview.Location = New System.Drawing.Point(11, 80)
      Me.buttonPreview.Name = "buttonPreview"
      Me.buttonPreview.Size = New System.Drawing.Size(74, 24)
      Me.buttonPreview.TabIndex = 12
      Me.buttonPreview.Text = "&Preview"
      '
      'frmSuppliesReport
      '
      Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
      Me.ClientSize = New System.Drawing.Size(280, 118)
      Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.GroupBox2, Me.buttonCancel, Me.buttonPrint, Me.buttonPreview})
      Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
      Me.MaximizeBox = False
      Me.MinimizeBox = False
      Me.Name = "frmSuppliesReport"
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "Supplies Inventory Report"
      Me.GroupBox2.ResumeLayout(False)
      Me.ResumeLayout(False)

   End Sub

#End Region
   Private _preview As Boolean

#Region " Form & Control Events "
   Private Sub buttonPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonPreview.Click
      _preview = True
      PrintReport()
   End Sub

   Private Sub buttonCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonCancel.Click
      Me.Close()
      DoEvents()
   End Sub

   Private Sub buttonPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonPrint.Click
      _preview = False
      PrintReport()
   End Sub


#End Region

#Region " Private Methods "
   ''' <summary>
   '''
   ''' </summary>
   Private Sub PrintReport()
      Dim sb As New StringBuilder()
      Dim SQL As String
      Dim oDA As New CDataAccess()

      Try
         Dim decTotal As Decimal = 0
         Dim i As Integer
         Dim dt As New DataTable()
         Dim dr As DataRow
         Dim d As Decimal

         SQL = "select * from products "
         SQL &= "order by "
         If Me.optEquipIdSort.Checked Then
            SQL &= "productid"
         Else
            SQL &= "productname"
         End If
         If oDA.SendQuery(SQL, dt, ConnectString) = 0 Then
            MsgBox("No records for report.", MsgBoxStyle.Information)
            Exit Sub
         End If

         Dim colHdr As String = _
             "Product ID".PadRight(12) & _
             "Product Name".PadRight(25) & _
             "In Stock".PadRight(10) & _
             "Price/Unit".PadRight(10) & _
             "Ext Price".PadRight(10)

         For i = 0 To dt.Rows.Count - 1
            dr = dt.Rows(i)

            sb.Append(CType(dr("productid"), String).PadRight(12))
            sb.Append(LS(CType(dr("productname"), String), 24).PadRight(25))
            sb.Append(Format(dr("unitsinstock"), "#,##0").PadLeft(10))
            d = MND(dr("priceperunit"))
            sb.Append(FormatCurrency(d).PadLeft(10))
            decTotal += (d * MNI(dr("unitsinstock")))
            sb.Append(FormatCurrency(d * MND(dr("unitsinstock"))).PadLeft(10))
            sb.Append(vbCrLf)
         Next

         sb.Append(vbCrLf & "Total Value of Inventory:".PadRight(12 + 25 + 10 + 10))
         sb.Append(FormatCurrency(decTotal).PadLeft(10) & vbCrLf)

         Dim ops As New CPrintStringNew()
         ops.TitleFontStyle = "BI"
         ops.TitleFontSize = REPORT_TITLE_FONT_SIZE
         If _preview Then
            ops.PrintPreview(80, sb.ToString, _
            ReportName, _
            "Product Inventory Report", _
            colHdr1:=colHdr)
         Else
            ops.StartPrint(80, sb.ToString, _
               ReportName, _
               "Product Inventory Report", _
               colHdr1:=colHdr)
         End If
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

#End Region



End Class
