Public Class frmMeterReadingMaint
   Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

   Public Sub New()
      MyBase.New()

      'This call is required by the Windows Form Designer.
      InitializeComponent()

      'Add any initialization after the InitializeComponent() call
      oDA = New CDataAccess()
      oCG = New CGrid()
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
   Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
   Friend WithEvents lblProductId As System.Windows.Forms.Label
   Friend WithEvents lblProductName As System.Windows.Forms.Label
   Friend WithEvents lblProductDesc As System.Windows.Forms.Label
   Friend WithEvents Label4 As System.Windows.Forms.Label
   Public WithEvents cmdAdd As System.Windows.Forms.Button
   Public WithEvents cmdUpdate As System.Windows.Forms.Button
   Public WithEvents cmdDelete As System.Windows.Forms.Button
   Public WithEvents cmdClose As System.Windows.Forms.Button
   Friend WithEvents lblUIS As System.Windows.Forms.Label
   Friend WithEvents lblRL As System.Windows.Forms.Label
   Friend WithEvents dgMeters As System.Windows.Forms.DataGrid
   Friend WithEvents textCustomerID As System.Windows.Forms.TextBox
   Friend WithEvents textDateEntered As System.Windows.Forms.TextBox
   Friend WithEvents textMeterReading As System.Windows.Forms.TextBox
   Friend WithEvents textEquipName As System.Windows.Forms.TextBox
   Friend WithEvents textEquipID As System.Windows.Forms.TextBox
   Friend WithEvents cbEntryType As System.Windows.Forms.ComboBox
   Friend WithEvents dgEquip As System.Windows.Forms.DataGrid
   Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
   Friend WithEvents lblID As System.Windows.Forms.Label
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMeterReadingMaint))
        Me.dgMeters = New System.Windows.Forms.DataGrid()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.lblID = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cbEntryType = New System.Windows.Forms.ComboBox()
        Me.lblRL = New System.Windows.Forms.Label()
        Me.textCustomerID = New System.Windows.Forms.TextBox()
        Me.lblUIS = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.lblProductDesc = New System.Windows.Forms.Label()
        Me.lblProductName = New System.Windows.Forms.Label()
        Me.lblProductId = New System.Windows.Forms.Label()
        Me.textDateEntered = New System.Windows.Forms.TextBox()
        Me.textMeterReading = New System.Windows.Forms.TextBox()
        Me.textEquipName = New System.Windows.Forms.TextBox()
        Me.textEquipID = New System.Windows.Forms.TextBox()
        Me.cmdAdd = New System.Windows.Forms.Button()
        Me.cmdUpdate = New System.Windows.Forms.Button()
        Me.cmdDelete = New System.Windows.Forms.Button()
        Me.cmdClose = New System.Windows.Forms.Button()
        Me.dgEquip = New System.Windows.Forms.DataGrid()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        CType(Me.dgMeters, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgEquip, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgMeters
        '
        Me.dgMeters.AllowSorting = False
        Me.dgMeters.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgMeters.CaptionText = "Click Row to Update or Add Button to Add Meter Reading for Equipment"
        Me.dgMeters.DataMember = ""
        Me.dgMeters.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dgMeters.Location = New System.Drawing.Point(1, 196)
        Me.dgMeters.Name = "dgMeters"
        Me.dgMeters.Size = New System.Drawing.Size(565, 177)
        Me.dgMeters.TabIndex = 17
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.lblID)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.cbEntryType)
        Me.GroupBox1.Controls.Add(Me.lblRL)
        Me.GroupBox1.Controls.Add(Me.textCustomerID)
        Me.GroupBox1.Controls.Add(Me.lblUIS)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.lblProductDesc)
        Me.GroupBox1.Controls.Add(Me.lblProductName)
        Me.GroupBox1.Controls.Add(Me.lblProductId)
        Me.GroupBox1.Controls.Add(Me.textDateEntered)
        Me.GroupBox1.Controls.Add(Me.textMeterReading)
        Me.GroupBox1.Controls.Add(Me.textEquipName)
        Me.GroupBox1.Controls.Add(Me.textEquipID)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 377)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(320, 164)
        Me.GroupBox1.TabIndex = 18
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Selected Items"
        '
        'lblID
        '
        Me.lblID.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblID.Location = New System.Drawing.Point(240, 137)
        Me.lblID.Name = "lblID"
        Me.lblID.Size = New System.Drawing.Size(48, 19)
        Me.lblID.TabIndex = 15
        Me.lblID.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(208, 142)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(18, 13)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "ID"
        '
        'cbEntryType
        '
        Me.cbEntryType.Items.AddRange(New Object() {"In", "Out", "Initial"})
        Me.cbEntryType.Location = New System.Drawing.Point(97, 111)
        Me.cbEntryType.Name = "cbEntryType"
        Me.cbEntryType.Size = New System.Drawing.Size(103, 21)
        Me.cbEntryType.TabIndex = 5
        '
        'lblRL
        '
        Me.lblRL.AutoSize = True
        Me.lblRL.Location = New System.Drawing.Point(18, 137)
        Me.lblRL.Name = "lblRL"
        Me.lblRL.Size = New System.Drawing.Size(56, 13)
        Me.lblRL.TabIndex = 13
        Me.lblRL.Text = "Invoice ID"
        '
        'textCustomerID
        '
        Me.textCustomerID.Location = New System.Drawing.Point(97, 137)
        Me.textCustomerID.Name = "textCustomerID"
        Me.textCustomerID.Size = New System.Drawing.Size(100, 20)
        Me.textCustomerID.TabIndex = 6
        Me.textCustomerID.Tag = "(No Auto Formatting)"
        Me.textCustomerID.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblUIS
        '
        Me.lblUIS.AutoSize = True
        Me.lblUIS.Location = New System.Drawing.Point(27, 115)
        Me.lblUIS.Name = "lblUIS"
        Me.lblUIS.Size = New System.Drawing.Size(58, 13)
        Me.lblUIS.TabIndex = 12
        Me.lblUIS.Text = "Entry Type"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(15, 90)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(70, 13)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Date Entered"
        '
        'lblProductDesc
        '
        Me.lblProductDesc.AutoSize = True
        Me.lblProductDesc.Location = New System.Drawing.Point(8, 66)
        Me.lblProductDesc.Name = "lblProductDesc"
        Me.lblProductDesc.Size = New System.Drawing.Size(77, 13)
        Me.lblProductDesc.TabIndex = 10
        Me.lblProductDesc.Text = "Meter Reading"
        '
        'lblProductName
        '
        Me.lblProductName.AutoSize = True
        Me.lblProductName.Location = New System.Drawing.Point(20, 42)
        Me.lblProductName.Name = "lblProductName"
        Me.lblProductName.Size = New System.Drawing.Size(65, 13)
        Me.lblProductName.TabIndex = 8
        Me.lblProductName.Text = "Equip Name"
        '
        'lblProductId
        '
        Me.lblProductId.AutoSize = True
        Me.lblProductId.Location = New System.Drawing.Point(38, 18)
        Me.lblProductId.Name = "lblProductId"
        Me.lblProductId.Size = New System.Drawing.Size(48, 13)
        Me.lblProductId.TabIndex = 7
        Me.lblProductId.Text = "Equip ID"
        '
        'textDateEntered
        '
        Me.textDateEntered.Location = New System.Drawing.Point(97, 88)
        Me.textDateEntered.Name = "textDateEntered"
        Me.textDateEntered.Size = New System.Drawing.Size(158, 20)
        Me.textDateEntered.TabIndex = 4
        Me.textDateEntered.Tag = "(No Auto Formatting)"
        '
        'textMeterReading
        '
        Me.textMeterReading.Location = New System.Drawing.Point(97, 63)
        Me.textMeterReading.Name = "textMeterReading"
        Me.textMeterReading.Size = New System.Drawing.Size(100, 20)
        Me.textMeterReading.TabIndex = 3
        Me.textMeterReading.Tag = "(No Auto Formatting)"
        Me.textMeterReading.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'textEquipName
        '
        Me.textEquipName.Location = New System.Drawing.Point(97, 38)
        Me.textEquipName.Name = "textEquipName"
        Me.textEquipName.Size = New System.Drawing.Size(209, 20)
        Me.textEquipName.TabIndex = 1
        Me.textEquipName.Tag = "(No Auto Formatting)"
        '
        'textEquipID
        '
        Me.textEquipID.Location = New System.Drawing.Point(97, 14)
        Me.textEquipID.Name = "textEquipID"
        Me.textEquipID.Size = New System.Drawing.Size(100, 20)
        Me.textEquipID.TabIndex = 0
        Me.textEquipID.Tag = "(No Auto Formatting)"
        '
        'cmdAdd
        '
        Me.cmdAdd.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdAdd.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAdd.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAdd.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAdd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAdd.Location = New System.Drawing.Point(447, 449)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAdd.Size = New System.Drawing.Size(101, 26)
        Me.cmdAdd.TabIndex = 1
        Me.cmdAdd.Text = "Setup to &Add"
        Me.ToolTip1.SetToolTip(Me.cmdAdd, "Click to selected equipment in the upper grid to the lower grid, enter the meter " & _
                "reading, and click Save")
        Me.cmdAdd.UseVisualStyleBackColor = False
        '
        'cmdUpdate
        '
        Me.cmdUpdate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdUpdate.BackColor = System.Drawing.SystemColors.Control
        Me.cmdUpdate.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdUpdate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdUpdate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdUpdate.Location = New System.Drawing.Point(447, 417)
        Me.cmdUpdate.Name = "cmdUpdate"
        Me.cmdUpdate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdUpdate.Size = New System.Drawing.Size(101, 26)
        Me.cmdUpdate.TabIndex = 0
        Me.cmdUpdate.Text = "&Save Updates"
        Me.ToolTip1.SetToolTip(Me.cmdUpdate, "Click to Save any update or a new meter reading")
        Me.cmdUpdate.UseVisualStyleBackColor = False
        '
        'cmdDelete
        '
        Me.cmdDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdDelete.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDelete.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdDelete.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDelete.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDelete.Location = New System.Drawing.Point(447, 481)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDelete.Size = New System.Drawing.Size(101, 26)
        Me.cmdDelete.TabIndex = 2
        Me.cmdDelete.Text = "&Delete Selected"
        Me.cmdDelete.UseVisualStyleBackColor = False
        '
        'cmdClose
        '
        Me.cmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
        Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdClose.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdClose.Location = New System.Drawing.Point(447, 513)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClose.Size = New System.Drawing.Size(101, 26)
        Me.cmdClose.TabIndex = 3
        Me.cmdClose.Text = "&Close"
        Me.cmdClose.UseVisualStyleBackColor = False
        '
        'dgEquip
        '
        Me.dgEquip.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.dgEquip.CaptionVisible = False
        Me.dgEquip.DataMember = ""
        Me.dgEquip.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dgEquip.Location = New System.Drawing.Point(8, 16)
        Me.dgEquip.Name = "dgEquip"
        Me.dgEquip.Size = New System.Drawing.Size(424, 168)
        Me.dgEquip.TabIndex = 24
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.dgEquip)
        Me.GroupBox2.Location = New System.Drawing.Point(1, 0)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(439, 192)
        Me.GroupBox2.TabIndex = 25
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Select Equipment to Maintain"
        '
        'frmMeterReadingMaint
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(568, 550)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.cmdAdd)
        Me.Controls.Add(Me.cmdUpdate)
        Me.Controls.Add(Me.cmdDelete)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.dgMeters)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmMeterReadingMaint"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Maintain Meter Reading Records"
        CType(Me.dgMeters, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.dgEquip, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Module Variables "
   Dim msAddorEdit As String
   Dim mbDirty As Boolean
   Dim mbFormLoading As Boolean
   Private oDA As CDataAccess
   Private iHitRow As Integer
   Private dtMeters As New DataTable("dt")
   Private oCG As CGrid
   Private equipHitRow As Integer
   Private dtEquip As DataTable


#End Region

#Region " Private Methods "
   Private Sub ClearTextBoxes()
      Dim dt As New DataTable()
      Dim sql As String


      Try
         With Me
            .textEquipID.Text = String.Empty
            .textEquipName.Text = String.Empty
            .cbEntryType.Text = String.Empty
            .textMeterReading.Text = String.Empty
            .textCustomerID.Text = "0"
            If msAddorEdit = "A" Then
               .textEquipName.Text = dtEquip.Rows(equipHitRow).Item("equip_name")
               .textEquipID.Text = dtEquip.Rows(equipHitRow).Item("equip_id")
               .textDateEntered.Text = Now.ToString
               .cbEntryType.Text = "Initial"
            End If
         End With
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

   Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
      Dim SQL As String
      Dim dt As New DataTable()
      Dim oDA As New CDataAccess()
      Dim iRows As Integer

      Try
         Dim sErr As String = ""

         If MsgBox("Are you sure you want to delete the selected row?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
            Exit Sub
         End If
         SQL = "delete from meter_reading "
         SQL &= "where unique_id = " & Me.lblID.Text
         iRows = oDA.SendActionSql(SQL, ConnectString, sErr)
         If iRows = 0 Then
            MsgBox("Delete of equipment item failed.  " & Chr(10) & sErr, MsgBoxStyle.Critical)
            Exit Sub
         End If
         Me.LoadTheGrid()
         iHitRow = 0
         LoadTextBoxes()
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

   Private Sub LoadTextBoxes()
      Try

         With Me.dtMeters.Rows(Me.iHitRow)
            Me.textEquipID.Text = MNS(.Item("equip_id"))
            Me.textEquipName.Text = MNS(.Item("equip_name"))
            Me.textMeterReading.Text = Format(MNSng(.Item("meter_reading")), "0.00")
            Me.textDateEntered.Text = IIf(IsDBNull(.Item("date_entered")), "", Format(.Item("date_entered"), "MM/dd/yyyy HH:mm tt"))
            Me.textCustomerID.Text = MNI(.Item("invoice_id"))
            Me.cbEntryType.Text = MNS(.Item("entry_type"))
            Me.lblID.Text = MNI(.Item("unique_id"))
            mbDirty = False
         End With

      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try

   End Sub

   Private Sub LoadTheGrid()
      Dim SQL As String


      Try
         oCG.InitializeDatatableForStyles(dtMeters)
         oCG.BindDataTableToGrid(dtMeters, Me.dgMeters)

         SQL = "select m.equip_id,e.equip_name,m.meter_reading,m.date_entered, "
         SQL &= "m.invoice_id,m.entry_type, m.unique_id "
         SQL &= "from meter_reading m, equipment e "
         SQL &= "where m.equip_id = e.equip_id "
         SQL &= "and m.equip_id = '" & dtEquip.Rows(equipHitRow).Item("equip_id") & "' "
         SQL &= " order by m.equip_id"
         oDA.SendQuery(SQL, dtMeters, ConnectString, "dt")
         If dtMeters.Rows.Count > 0 Then
                Dim Formats() As String = _
                   {"", "60", "T", "L", _
                   "", "150", "T", "L", _
                   "##0.00", "60", "T", "L", _
                   "MM/dd/yyyy HH:mm tt", "120", "T", "L", _
                   "", "60", "T", "L", _
                   "", "60", "T", "L", _
                   "", "60", "T", "L"}
            oCG.SetTablesStyle(dtMeters, Me.dgMeters, Formats)
            oCG.BindDataTableToGrid(dtMeters, Me.dgMeters)
            oCG.DisableAddNew(Me.dgMeters, Me)
         End If
         mbDirty = False
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

   ''' <summary>
   ''' Load the equipment grid with items that require meters.
   ''' </summary>
   Private Sub LoadEquipGrid()
      Dim sql As String
      sql = "select equip_id,equip_name,serial_number from equipment "
      sql &= "where meter_required  "
      sql &= "order by equip_name"
      dtEquip = New DataTable("dt")
      If oDA.SendQuery(sql, dtEquip, ConnectString, "dt") > 0 Then
         Dim formats() As String = {"", "60", "T", "L", _
                                    "", "200", "T", "L", _
                                    "", "120", "T", "L"}

         oCG.SetTablesStyle(dtEquip, dgEquip, formats)
         oCG.BindDataTableToGrid(dtEquip, dgEquip)
         oCG.DisableAddNew(dgEquip, Me)
         equipHitRow = 0
         LoadTheGrid()
      End If
   End Sub

#End Region

#Region " Form & Control Events "
   Private Sub cmdUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdate.Click
      Dim SQL As String
      Dim sErr As String


      Try
         With Me
            If msAddorEdit = "A" Then
               SQL = "insert into  meter_reading"
               SQL &= "(equip_id, meter_reading, date_entered, invoice_id,entry_type) "
               SQL &= "values("
               SQL &= "'" & .textEquipID.Text & "', "
               SQL &= .textMeterReading.Text & ", "
               SQL &= "#" & .textDateEntered.Text & "#, "
               SQL &= .textCustomerID.Text & ", "
               SQL &= "'" & .cbEntryType.Text & "' "
               SQL &= ")"
            Else
               SQL = "update meter_reading "
               SQL &= "set meter_reading = " & .textMeterReading.Text & ", "
               SQL &= "date_entered = #" & .textDateEntered.Text & "#, "
               SQL &= "invoice_id = " & .textCustomerID.Text & ", "
               SQL &= "entry_type = '" & .cbEntryType.Text & "' "
               SQL &= "where unique_id = " & Me.lblID.Text
            End If
         End With
         If oDA.SendActionSql(SQL, ConnectString, sErr) < 1 Then
            MsgBox("Update failed: " & Chr(10) & sErr, MsgBoxStyle.Critical)
         End If

         msAddorEdit = "E"
         LoadTheGrid()
         iHitRow = 0
         LoadTextBoxes()
         mbDirty = False
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try

   End Sub

   Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
      Dim SQL As String
      Dim dt As New DataTable()


      Try
         If mbDirty Then
            If MsgBox("You have unsaved changes; do you want to add without saving your changes?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
               Exit Sub
            End If
         End If

         msAddorEdit = "A"

         ClearTextBoxes()
         SQL = "select max(unique_id) from meter_reading "
         If oDA.SendQuery(SQL, dt, ConnectString) = 0 Then
            Throw New System.Exception("Can't read meter_reading table to get unique ID.")
         End If
            Me.lblID.Text = MNI(dt.Rows(0).Item(0)) + 1

      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

   Private Sub dgMeters_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgMeters.MouseUp
      Try
         Dim pt = New Point(e.X, e.Y)
         Dim hti As DataGrid.HitTestInfo = Me.dgMeters.HitTest(pt)
         Me.dgMeters.Select(hti.Row)
         iHitRow = hti.Row
         Me.LoadTextBoxes()
         mbDirty = False
         msAddorEdit = "E"
      Catch ex As System.Exception
         'StructuredErrorHandler(ex)
      End Try
   End Sub

   Private Sub cmdClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdClose.Click
      If mbDirty Then
         If MsgBox("You have unsaved changes; do you want to close without saving your changes?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
            Exit Sub
         End If
      End If

      Me.Close()
      System.Windows.Forms.Application.DoEvents()
   End Sub
   Private Sub textEquipID_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles textEquipID.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
   End Sub
   Private Sub textEquipID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles textEquipID.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub textEquipID_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles textEquipID.Enter
      textEquipID.SelectionStart = 0
      textEquipID.SelectionLength = textEquipID.Text.Trim.Length
   End Sub
   Private Sub textEquipName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles textEquipName.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
   End Sub
   Private Sub textEquipName_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles textEquipName.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub textEquipName_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles textEquipName.Enter
      textEquipName.SelectionStart = 0
      textEquipName.SelectionLength = textEquipName.Text.Trim.Length
   End Sub
   Private Sub textMeterReading_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles textMeterReading.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
   End Sub
   Private Sub textMeterReading_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles textMeterReading.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub textMeterReading_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles textMeterReading.Enter
      textMeterReading.SelectionStart = 0
      textMeterReading.SelectionLength = textMeterReading.Text.Trim.Length
   End Sub
   Private Sub textCustomerID_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles textCustomerID.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
   End Sub
   Private Sub textCustomerID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles textCustomerID.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub textCustomerID_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles textCustomerID.Enter
      textCustomerID.SelectionStart = 0
      textCustomerID.SelectionLength = textCustomerID.Text.Trim.Length
   End Sub
   Private Sub textDateEntered_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles textDateEntered.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
   End Sub
   Private Sub textDateEntered_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles textDateEntered.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub textDateEntered_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles textDateEntered.Enter
      textDateEntered.SelectionStart = 0
      textDateEntered.SelectionLength = textDateEntered.Text.Trim.Length
   End Sub

   Private Sub textDateEntered_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles textDateEntered.Leave
      'textDateEntered.Text = Format(textDateEntered.Text, "MM/dd/yyyy HH:mm tt")
   End Sub


   Private Sub frmMeterReadingMaint_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
      LoadEquipGrid()
   End Sub

   Private Sub dgEquip_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgEquip.MouseUp
      equipHitRow = dgEquip.CurrentCell.RowNumber
      oCG.SelectCkBoxRow(dgEquip, equipHitRow)
      LoadTheGrid()
   End Sub


#End Region

End Class
