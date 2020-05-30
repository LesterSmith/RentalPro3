Public Class frmCategoryMaint
   Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

   Public Sub New()
      MyBase.New()

      'This call is required by the Windows Form Designer.
      InitializeComponent()

      'Add any initialization after the InitializeComponent() call
      oDA = New CDataAccess()
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
   Public WithEvents _lblLabels_1 As System.Windows.Forms.Label
   Public WithEvents _lblLabels_0 As System.Windows.Forms.Label
   Public WithEvents cmdClose As System.Windows.Forms.Button
   Public WithEvents cmdDelete As System.Windows.Forms.Button
   Public WithEvents cmdUpdate As System.Windows.Forms.Button
   Public WithEvents cmdAdd As System.Windows.Forms.Button
   Friend WithEvents dbgCategory As System.Windows.Forms.DataGrid
   Public WithEvents txtTypeID As System.Windows.Forms.TextBox
   Public WithEvents txtEquipType As System.Windows.Forms.TextBox
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
      Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmCategoryMaint))
      Me.txtTypeID = New System.Windows.Forms.TextBox()
      Me.txtEquipType = New System.Windows.Forms.TextBox()
      Me._lblLabels_1 = New System.Windows.Forms.Label()
      Me._lblLabels_0 = New System.Windows.Forms.Label()
      Me.dbgCategory = New System.Windows.Forms.DataGrid()
      Me.cmdClose = New System.Windows.Forms.Button()
      Me.cmdDelete = New System.Windows.Forms.Button()
      Me.cmdUpdate = New System.Windows.Forms.Button()
      Me.cmdAdd = New System.Windows.Forms.Button()
      CType(Me.dbgCategory, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.SuspendLayout()
      '
      'txtTypeID
      '
      Me.txtTypeID.AcceptsReturn = True
      Me.txtTypeID.AutoSize = False
      Me.txtTypeID.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(255, Byte))
      Me.txtTypeID.Cursor = System.Windows.Forms.Cursors.IBeam
      Me.txtTypeID.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.txtTypeID.ForeColor = System.Drawing.SystemColors.WindowText
      Me.txtTypeID.Location = New System.Drawing.Point(75, 198)
      Me.txtTypeID.MaxLength = 0
      Me.txtTypeID.Name = "txtTypeID"
      Me.txtTypeID.ReadOnly = True
      Me.txtTypeID.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.txtTypeID.Size = New System.Drawing.Size(78, 19)
      Me.txtTypeID.TabIndex = 7
      Me.txtTypeID.Tag = "(No Auto Formatting)"
      Me.txtTypeID.Text = ""
      Me.txtTypeID.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      '
      'txtEquipType
      '
      Me.txtEquipType.AcceptsReturn = True
      Me.txtEquipType.AutoSize = False
      Me.txtEquipType.BackColor = System.Drawing.SystemColors.Window
      Me.txtEquipType.Cursor = System.Windows.Forms.Cursors.IBeam
      Me.txtEquipType.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.txtEquipType.ForeColor = System.Drawing.SystemColors.WindowText
      Me.txtEquipType.Location = New System.Drawing.Point(75, 170)
      Me.txtEquipType.MaxLength = 0
      Me.txtEquipType.Name = "txtEquipType"
      Me.txtEquipType.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.txtEquipType.Size = New System.Drawing.Size(214, 19)
      Me.txtEquipType.TabIndex = 5
      Me.txtEquipType.Tag = "(No Auto Formatting)"
      Me.txtEquipType.Text = ""
      '
      '_lblLabels_1
      '
      Me._lblLabels_1.AutoSize = True
      Me._lblLabels_1.BackColor = System.Drawing.SystemColors.Control
      Me._lblLabels_1.Cursor = System.Windows.Forms.Cursors.Default
      Me._lblLabels_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me._lblLabels_1.ForeColor = System.Drawing.SystemColors.ControlText
      Me._lblLabels_1.Location = New System.Drawing.Point(30, 198)
      Me._lblLabels_1.Name = "_lblLabels_1"
      Me._lblLabels_1.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me._lblLabels_1.Size = New System.Drawing.Size(46, 13)
      Me._lblLabels_1.TabIndex = 6
      Me._lblLabels_1.Text = "Type ID:"
      '
      '_lblLabels_0
      '
      Me._lblLabels_0.AutoSize = True
      Me._lblLabels_0.BackColor = System.Drawing.SystemColors.Control
      Me._lblLabels_0.Cursor = System.Windows.Forms.Cursors.Default
      Me._lblLabels_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me._lblLabels_0.ForeColor = System.Drawing.SystemColors.ControlText
      Me._lblLabels_0.Location = New System.Drawing.Point(12, 170)
      Me._lblLabels_0.Name = "_lblLabels_0"
      Me._lblLabels_0.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me._lblLabels_0.Size = New System.Drawing.Size(63, 13)
      Me._lblLabels_0.TabIndex = 4
      Me._lblLabels_0.Text = "Equip Type:"
      '
      'dbgCategory
      '
      Me.dbgCategory.AllowSorting = False
      Me.dbgCategory.DataMember = ""
      Me.dbgCategory.HeaderForeColor = System.Drawing.SystemColors.ControlText
      Me.dbgCategory.Location = New System.Drawing.Point(8, 1)
      Me.dbgCategory.Name = "dbgCategory"
      Me.dbgCategory.Size = New System.Drawing.Size(392, 148)
      Me.dbgCategory.TabIndex = 8
      '
      'cmdClose
      '
      Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
      Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
      Me.cmdClose.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
      Me.cmdClose.Location = New System.Drawing.Point(325, 256)
      Me.cmdClose.Name = "cmdClose"
      Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.cmdClose.Size = New System.Drawing.Size(73, 29)
      Me.cmdClose.TabIndex = 13
      Me.cmdClose.Text = "&Close"
      '
      'cmdDelete
      '
      Me.cmdDelete.BackColor = System.Drawing.SystemColors.Control
      Me.cmdDelete.Cursor = System.Windows.Forms.Cursors.Default
      Me.cmdDelete.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.cmdDelete.ForeColor = System.Drawing.SystemColors.ControlText
      Me.cmdDelete.Location = New System.Drawing.Point(325, 223)
      Me.cmdDelete.Name = "cmdDelete"
      Me.cmdDelete.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.cmdDelete.Size = New System.Drawing.Size(73, 29)
      Me.cmdDelete.TabIndex = 12
      Me.cmdDelete.Text = "&Delete"
      '
      'cmdUpdate
      '
      Me.cmdUpdate.BackColor = System.Drawing.SystemColors.Control
      Me.cmdUpdate.Cursor = System.Windows.Forms.Cursors.Default
      Me.cmdUpdate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.cmdUpdate.ForeColor = System.Drawing.SystemColors.ControlText
      Me.cmdUpdate.Location = New System.Drawing.Point(325, 190)
      Me.cmdUpdate.Name = "cmdUpdate"
      Me.cmdUpdate.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.cmdUpdate.Size = New System.Drawing.Size(73, 29)
      Me.cmdUpdate.TabIndex = 11
      Me.cmdUpdate.Text = "&Update"
      '
      'cmdAdd
      '
      Me.cmdAdd.BackColor = System.Drawing.SystemColors.Control
      Me.cmdAdd.Cursor = System.Windows.Forms.Cursors.Default
      Me.cmdAdd.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.cmdAdd.ForeColor = System.Drawing.SystemColors.ControlText
      Me.cmdAdd.Location = New System.Drawing.Point(325, 157)
      Me.cmdAdd.Name = "cmdAdd"
      Me.cmdAdd.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.cmdAdd.Size = New System.Drawing.Size(73, 29)
      Me.cmdAdd.TabIndex = 10
      Me.cmdAdd.Text = "&Add"
      '
      'frmCategoryMaint
      '
      Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
      Me.ClientSize = New System.Drawing.Size(408, 293)
      Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdClose, Me.cmdDelete, Me.cmdUpdate, Me.cmdAdd, Me.dbgCategory, Me.txtTypeID, Me.txtEquipType, Me._lblLabels_1, Me._lblLabels_0})
      Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
      Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
      Me.Name = "frmCategoryMaint"
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "Maintain Equipment Category"
      CType(Me.dbgCategory, System.ComponentModel.ISupportInitialize).EndInit()
      Me.ResumeLayout(False)

   End Sub

#End Region
   Dim msAddorEdit As String
   Dim mbDirty As Boolean
   Dim mbFormLoading As Boolean
   Private oDA As CDataAccess
   Private iHitRow As Integer
   Private dtCM As New DataTable("dt")


   Private Sub cmdUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdate.Click
      Dim SQL As String
      Dim sErr As String


      Try
         With Me
            If msAddorEdit = "A" Then
               SQL = "insert into equipment_type "
               SQL = SQL & "(equip_type,equip_type_id) "
               SQL = SQL & "values('"
               SQL = SQL & Replace(.txtEquipType.Text, "'", "''") & "', "
               SQL = SQL & .txtTypeID.Text & ") "
            Else
               SQL = "update equipment_type "
               SQL = SQL & "set equip_type = '" & Replace(.txtEquipType.Text, "'", "''") & "'  "
               SQL = SQL & "where equip_type_id = " & .txtTypeID.Text
            End If
         End With
         oDA.SendActionSql(SQL, ConnectString, sErr)

         msAddorEdit = "E"
         LoadTheGrid()
         mbDirty = False
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

   Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
      Dim SQL As String
      Dim dt As New DataTable()

      If mbDirty Then
         If MsgBox("You have unsaved changes; do you want to add without saving your changes?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
            Exit Sub
         End If
      End If


      With Me
         .txtEquipType.Text = ""
         mbDirty = False
      End With
      msAddorEdit = "A"
      SQL = "select max(equip_type_id) from equipment_type"
      oDA.SendQuery(SQL, dt, ConnectString)
      If dt.Rows.Count = 0 Then
         Me.txtTypeID.Text = CStr(1)
      Else
         If IsDBNull(dt.Rows(0).Item(0)) Then
            Me.txtTypeID.Text = CStr(1)
         Else
            Me.txtTypeID.Text = CStr(Val(dt.Rows(0).Item(0)) + 1)
         End If
      End If

      Me.txtEquipType.Focus()
   End Sub

   Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
      If mbDirty Then
         If MsgBox("You have unsaved changes; do you want to close without saving your changes?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
            Exit Sub
         End If
      End If

      Me.Close()
      System.Windows.Forms.Application.DoEvents()
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
         SQL = "delete from equipment_type "
         SQL &= "where equip_type_id = " & Me.dtCM.Rows(Me.iHitRow).Item("equip_type_id") & " "
         iRows = oDA.SendActionSql(SQL, ConnectString, sErr)
         If iRows = 0 Then
            MsgBox("Delete of equipment item failed.  " & Chr(10) & sErr, MsgBoxStyle.Critical)
            Exit Sub
         End If
         Me.LoadTheGrid()

      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try

   End Sub

   Private Sub dbgCategory_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles dbgCategory.Click
      msAddorEdit = "E"
   End Sub

   Private Sub LoadTheGrid()
      Dim SQL As String


      Try
         dtCM = New DataTable("dt")
         SQL = "select equip_type_id, "
         SQL = SQL & "equip_type "
         SQL = SQL & "from equipment_type "
         SQL = SQL & " order by equip_type_id "
         oDA.SendQuery(SQL, dtCM, ConnectString)
         If dtCM.Rows.Count > 0 Then
            Me.dbgCategory.SetDataBinding(dtCM, "")

         End If
         mbDirty = False
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

   Private Sub frmCategoryMaint_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
      'IncrementChildCount Me
      LoadTheGrid()
      Me.txtEquipType.Text = Me.dtCM.Rows(0).Item("equip_type")
      Me.txtTypeID.Text = Me.dtCM.Rows(0).Item("equip_type_id")
      msAddorEdit = "E"
      mbFormLoading = True
   End Sub
   Private Sub txtEquipType_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEquipType.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
   End Sub
   Private Sub txtEquipType_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtEquipType.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub txtEquipType_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEquipType.Enter
      txtEquipType.SelectionStart = 0
      txtEquipType.SelectionLength = txtEquipType.Text.Trim.Length
   End Sub
   Private Sub txtTypeID_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTypeID.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
   End Sub
   Private Sub txtTypeID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTypeID.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub txtTypeID_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTypeID.Enter
      txtTypeID.SelectionStart = 0
      txtTypeID.SelectionLength = txtTypeID.Text.Trim.Length
   End Sub

   Private Sub dbgCategory_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dbgCategory.MouseUp
      Try
         Dim pt = New Point(e.X, e.Y)
         Dim hti As DataGrid.HitTestInfo = Me.dbgCategory.HitTest(pt)
         Me.dbgCategory.Select(hti.Row)
         iHitRow = hti.Row
         Me.txtEquipType.Text = dtCM.Rows(iHitRow).Item("equip_type")
         Me.txtTypeID.Text = dtCM.Rows(iHitRow).Item("equip_type_id")
         mbDirty = False
         msAddorEdit = "E"
      Catch ex As System.Exception
         'StructuredErrorHandler(ex)
      End Try
   End Sub
End Class
