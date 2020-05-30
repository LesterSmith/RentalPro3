Imports System.Windows.Forms.Application
Imports System.Text

Public Class frmMeterReadingReport
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
   Friend WithEvents buttonCancel As System.Windows.Forms.Button
   Friend WithEvents buttonPrint As System.Windows.Forms.Button
   Friend WithEvents buttonPreview As System.Windows.Forms.Button
   Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
   Friend WithEvents optLatestReading As System.Windows.Forms.RadioButton
   Friend WithEvents optAllItems As System.Windows.Forms.RadioButton
   Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
   Friend WithEvents optEquipIdSort As System.Windows.Forms.RadioButton
   Friend WithEvents optEquipNameSort As System.Windows.Forms.RadioButton
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
      Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMeterReadingReport))
      Me.buttonCancel = New System.Windows.Forms.Button()
      Me.buttonPrint = New System.Windows.Forms.Button()
      Me.buttonPreview = New System.Windows.Forms.Button()
      Me.GroupBox1 = New System.Windows.Forms.GroupBox()
      Me.optLatestReading = New System.Windows.Forms.RadioButton()
      Me.optAllItems = New System.Windows.Forms.RadioButton()
      Me.GroupBox2 = New System.Windows.Forms.GroupBox()
      Me.optEquipIdSort = New System.Windows.Forms.RadioButton()
      Me.optEquipNameSort = New System.Windows.Forms.RadioButton()
      Me.GroupBox1.SuspendLayout()
      Me.GroupBox2.SuspendLayout()
      Me.SuspendLayout()
      '
      'buttonCancel
      '
      Me.buttonCancel.Location = New System.Drawing.Point(198, 100)
      Me.buttonCancel.Name = "buttonCancel"
      Me.buttonCancel.Size = New System.Drawing.Size(74, 24)
      Me.buttonCancel.TabIndex = 9
      Me.buttonCancel.Text = "&Cancel"
      '
      'buttonPrint
      '
      Me.buttonPrint.Location = New System.Drawing.Point(106, 100)
      Me.buttonPrint.Name = "buttonPrint"
      Me.buttonPrint.Size = New System.Drawing.Size(74, 24)
      Me.buttonPrint.TabIndex = 8
      Me.buttonPrint.Text = "P&rint"
      '
      'buttonPreview
      '
      Me.buttonPreview.Location = New System.Drawing.Point(14, 100)
      Me.buttonPreview.Name = "buttonPreview"
      Me.buttonPreview.Size = New System.Drawing.Size(74, 24)
      Me.buttonPreview.TabIndex = 7
      Me.buttonPreview.Text = "&Preview"
      '
      'GroupBox1
      '
      Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.optLatestReading, Me.optAllItems})
      Me.GroupBox1.Location = New System.Drawing.Point(6, 5)
      Me.GroupBox1.Name = "GroupBox1"
      Me.GroupBox1.Size = New System.Drawing.Size(136, 80)
      Me.GroupBox1.TabIndex = 10
      Me.GroupBox1.TabStop = False
      Me.GroupBox1.Text = "Select Print Details"
      '
      'optLatestReading
      '
      Me.optLatestReading.Location = New System.Drawing.Point(4, 44)
      Me.optLatestReading.Name = "optLatestReading"
      Me.optLatestReading.Size = New System.Drawing.Size(128, 24)
      Me.optLatestReading.TabIndex = 3
      Me.optLatestReading.Text = "Latest Reading Only"
      '
      'optAllItems
      '
      Me.optAllItems.Checked = True
      Me.optAllItems.Location = New System.Drawing.Point(4, 20)
      Me.optAllItems.Name = "optAllItems"
      Me.optAllItems.Size = New System.Drawing.Size(128, 16)
      Me.optAllItems.TabIndex = 2
      Me.optAllItems.TabStop = True
      Me.optAllItems.Text = "All Items"
      '
      'GroupBox2
      '
      Me.GroupBox2.Controls.AddRange(New System.Windows.Forms.Control() {Me.optEquipNameSort, Me.optEquipIdSort})
      Me.GroupBox2.Location = New System.Drawing.Point(152, 5)
      Me.GroupBox2.Name = "GroupBox2"
      Me.GroupBox2.Size = New System.Drawing.Size(133, 80)
      Me.GroupBox2.TabIndex = 11
      Me.GroupBox2.TabStop = False
      Me.GroupBox2.Text = "Report Sort Order"
      '
      'optEquipIdSort
      '
      Me.optEquipIdSort.Checked = True
      Me.optEquipIdSort.Location = New System.Drawing.Point(9, 18)
      Me.optEquipIdSort.Name = "optEquipIdSort"
      Me.optEquipIdSort.Size = New System.Drawing.Size(99, 19)
      Me.optEquipIdSort.TabIndex = 0
      Me.optEquipIdSort.TabStop = True
      Me.optEquipIdSort.Text = "Equipment ID"
      '
      'optEquipNameSort
      '
      Me.optEquipNameSort.Location = New System.Drawing.Point(9, 48)
      Me.optEquipNameSort.Name = "optEquipNameSort"
      Me.optEquipNameSort.Size = New System.Drawing.Size(111, 16)
      Me.optEquipNameSort.TabIndex = 1
      Me.optEquipNameSort.Text = "Equipment Name"
      '
      'frmMeterReadingReport
      '
      Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
      Me.ClientSize = New System.Drawing.Size(292, 142)
      Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.GroupBox2, Me.GroupBox1, Me.buttonCancel, Me.buttonPrint, Me.buttonPreview})
      Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
      Me.MaximizeBox = False
      Me.MinimizeBox = False
      Me.Name = "frmMeterReadingReport"
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "Print Meter Reading Report"
      Me.GroupBox1.ResumeLayout(False)
      Me.GroupBox2.ResumeLayout(False)
      Me.ResumeLayout(False)

   End Sub

#End Region
   Private _Preview As Boolean

#Region "Processing Methods"
   ''' <summary>
   '''
   ''' </summary>
   Public Sub PrintReport()
      Dim dt As New DataTable()
      Dim i As Integer
      Dim sb As New StringBuilder()
      Dim SQL As String
      Dim oDA As New CDataAccess()
      Dim lastEquip As String = String.Empty
      Dim dr As DataRow
      Dim s As String

      Try

         SQL = "select m.*,e.equip_name from meter_reading m, equipment e "
         SQL &= "where m.equip_id = e.equip_id "
         If Me.optEquipIdSort.Checked Then
            SQL &= "order by m.equip_id, date_entered desc"
         Else
            SQL &= "order by e.equip_name,date_entered desc"
         End If

         If oDA.SendQuery(SQL, dt, ConnectString, "dt") = 0 Then
            MsgBox("No records to print.", MsgBoxStyle.Information)
            Exit Sub
         End If

         Dim colHdr As String = _
            "Equip ID".PadRight(11) & _
            "Equipment Name".PadRight(28) & _
            "Date Entered".PadRight(21) & _
            "Invoice".PadRight(8) & _
            "Entry Type".PadRight(11) & _
            "Reading".PadRight(9)

         'If Me.optEquipIdSort.Checked Then
         '   lastEquip = MNS(dt.Rows(0).Item("equip_id"))
         'Else
         '   lastEquip = MNS(dt.Rows(0).Item("equip_name"))
         'End If

         For i = 0 To dt.Rows.Count - 1
            dr = dt.Rows(i)
            ' ck to see if printing all records or just latest
            If Me.optEquipIdSort.Checked Then
               If Me.optLatestReading.Checked Then
                  If MNS(dr("equip_id")) = lastEquip Then
                     GoTo GetNextRecord
                  End If
                  lastEquip = MNS(dr("equip_id"))
               End If
            Else
               If Me.optLatestReading.Checked Then
                  If MNS(dr("equip_name")) = lastEquip Then
                     GoTo GetNextRecord
                  End If
                  lastEquip = MNS(dr("equip_name"))
               End If
            End If

            sb.Append(MNS(dr("equip_id")).PadRight(11))
            sb.Append(LS(MNS(dr("equip_name")), 27).PadRight(28))
            If IsDBNull(dr("date_entered")) Then
               sb.Append(Space(21))
            Else
               sb.Append(Format(dr("date_entered"), "MM/dd/yyyy hh:mm tt").PadRight(21))
            End If
            sb.Append(MNS(dr("Invoice_id")).PadRight(8))
            sb.Append(LS(MNS(dr("entry_type")), 10).PadRight(11))
            sb.Append(Format(MNSng(dr("meter_reading")), "#,##0.00").PadLeft(9))
            sb.Append(vbCrLf)
GetNextRecord:
         Next i

         Dim oPS As New CPrintStringNew()
         oPS.TitleFontStyle = "BI"
         oPS.TitleFontSize = REPORT_TITLE_FONT_SIZE
         If _Preview Then
            oPS.PrintPreview(80, sb.ToString, _
            ReportName, _
            "Meter Reading Report", _
            colHdr1:=colHdr)
         Else
            oPS.StartPrint(80, sb.ToString, _
               ReportName, _
               "Meter Reading Report", _
               colHdr1:=colHdr)
         End If

      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub
#End Region


#Region "Form and Control Events"
   Private Sub buttonPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonPrint.Click
      _Preview = False
      Me.PrintReport()
   End Sub

   Private Sub buttonPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonPreview.Click
      _Preview = True
      PrintReport()
   End Sub

   Private Sub buttonCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonCancel.Click
      Me.Close()
      DoEvents()
   End Sub

#End Region

End Class
