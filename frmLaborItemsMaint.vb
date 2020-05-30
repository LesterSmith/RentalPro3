Public Class frmLaborItemsMaint
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
   Friend WithEvents txtProductDescription As System.Windows.Forms.TextBox
   Friend WithEvents txtPricePerUnit As System.Windows.Forms.TextBox
   Friend WithEvents lblProductDesc As System.Windows.Forms.Label
   Friend WithEvents Label4 As System.Windows.Forms.Label
   Public WithEvents cmdAdd As System.Windows.Forms.Button
   Public WithEvents cmdUpdate As System.Windows.Forms.Button
   Public WithEvents cmdDelete As System.Windows.Forms.Button
   Public WithEvents cmdClose As System.Windows.Forms.Button
   Friend WithEvents dbgSaleItems As System.Windows.Forms.DataGrid
   Friend WithEvents lblID As System.Windows.Forms.Label
   Friend WithEvents txtID As System.Windows.Forms.TextBox
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
      Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmLaborItemsMaint))
      Me.dbgSaleItems = New System.Windows.Forms.DataGrid()
      Me.GroupBox1 = New System.Windows.Forms.GroupBox()
      Me.Label4 = New System.Windows.Forms.Label()
      Me.lblProductDesc = New System.Windows.Forms.Label()
      Me.txtPricePerUnit = New System.Windows.Forms.TextBox()
      Me.txtProductDescription = New System.Windows.Forms.TextBox()
      Me.cmdAdd = New System.Windows.Forms.Button()
      Me.cmdUpdate = New System.Windows.Forms.Button()
      Me.cmdDelete = New System.Windows.Forms.Button()
      Me.cmdClose = New System.Windows.Forms.Button()
      Me.lblID = New System.Windows.Forms.Label()
      Me.txtID = New System.Windows.Forms.TextBox()
      CType(Me.dbgSaleItems, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.GroupBox1.SuspendLayout()
      Me.SuspendLayout()
      '
      'dbgSaleItems
      '
      Me.dbgSaleItems.AllowSorting = False
      Me.dbgSaleItems.DataMember = ""
      Me.dbgSaleItems.HeaderForeColor = System.Drawing.SystemColors.ControlText
      Me.dbgSaleItems.Name = "dbgSaleItems"
      Me.dbgSaleItems.Size = New System.Drawing.Size(485, 128)
      Me.dbgSaleItems.TabIndex = 17
      '
      'GroupBox1
      '
      Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtID, Me.lblID, Me.Label4, Me.lblProductDesc, Me.txtPricePerUnit, Me.txtProductDescription})
      Me.GroupBox1.Location = New System.Drawing.Point(8, 137)
      Me.GroupBox1.Name = "GroupBox1"
      Me.GroupBox1.Size = New System.Drawing.Size(320, 103)
      Me.GroupBox1.TabIndex = 18
      Me.GroupBox1.TabStop = False
      Me.GroupBox1.Text = "Selected Items"
      '
      'Label4
      '
      Me.Label4.AutoSize = True
      Me.Label4.Location = New System.Drawing.Point(44, 46)
      Me.Label4.Name = "Label4"
      Me.Label4.Size = New System.Drawing.Size(74, 13)
      Me.Label4.TabIndex = 7
      Me.Label4.Text = "Price Per Unit"
      '
      'lblProductDesc
      '
      Me.lblProductDesc.AutoSize = True
      Me.lblProductDesc.Location = New System.Drawing.Point(13, 22)
      Me.lblProductDesc.Name = "lblProductDesc"
      Me.lblProductDesc.Size = New System.Drawing.Size(93, 13)
      Me.lblProductDesc.TabIndex = 6
      Me.lblProductDesc.Text = "Labor Description"
      '
      'txtPricePerUnit
      '
      Me.txtPricePerUnit.Location = New System.Drawing.Point(122, 44)
      Me.txtPricePerUnit.Name = "txtPricePerUnit"
      Me.txtPricePerUnit.TabIndex = 3
      Me.txtPricePerUnit.Tag = "(No Auto Formatting)"
      Me.txtPricePerUnit.Text = ""
      Me.txtPricePerUnit.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      '
      'txtProductDescription
      '
      Me.txtProductDescription.Location = New System.Drawing.Point(122, 19)
      Me.txtProductDescription.MaxLength = 50
      Me.txtProductDescription.Name = "txtProductDescription"
      Me.txtProductDescription.Size = New System.Drawing.Size(188, 20)
      Me.txtProductDescription.TabIndex = 2
      Me.txtProductDescription.Tag = "(No Auto Formatting)"
      Me.txtProductDescription.Text = ""
      '
      'cmdAdd
      '
      Me.cmdAdd.BackColor = System.Drawing.SystemColors.Control
      Me.cmdAdd.Cursor = System.Windows.Forms.Cursors.Default
      Me.cmdAdd.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.cmdAdd.ForeColor = System.Drawing.SystemColors.ControlText
      Me.cmdAdd.Location = New System.Drawing.Point(384, 174)
      Me.cmdAdd.Name = "cmdAdd"
      Me.cmdAdd.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.cmdAdd.Size = New System.Drawing.Size(77, 26)
      Me.cmdAdd.TabIndex = 20
      Me.cmdAdd.Text = "&Add"
      '
      'cmdUpdate
      '
      Me.cmdUpdate.BackColor = System.Drawing.SystemColors.Control
      Me.cmdUpdate.Cursor = System.Windows.Forms.Cursors.Default
      Me.cmdUpdate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.cmdUpdate.ForeColor = System.Drawing.SystemColors.ControlText
      Me.cmdUpdate.Location = New System.Drawing.Point(384, 142)
      Me.cmdUpdate.Name = "cmdUpdate"
      Me.cmdUpdate.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.cmdUpdate.Size = New System.Drawing.Size(77, 26)
      Me.cmdUpdate.TabIndex = 19
      Me.cmdUpdate.Text = "&Save"
      '
      'cmdDelete
      '
      Me.cmdDelete.BackColor = System.Drawing.SystemColors.Control
      Me.cmdDelete.Cursor = System.Windows.Forms.Cursors.Default
      Me.cmdDelete.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.cmdDelete.ForeColor = System.Drawing.SystemColors.ControlText
      Me.cmdDelete.Location = New System.Drawing.Point(384, 206)
      Me.cmdDelete.Name = "cmdDelete"
      Me.cmdDelete.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.cmdDelete.Size = New System.Drawing.Size(77, 26)
      Me.cmdDelete.TabIndex = 21
      Me.cmdDelete.Text = "&Delete"
      '
      'cmdClose
      '
      Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
      Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
      Me.cmdClose.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
      Me.cmdClose.Location = New System.Drawing.Point(384, 238)
      Me.cmdClose.Name = "cmdClose"
      Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.cmdClose.Size = New System.Drawing.Size(77, 26)
      Me.cmdClose.TabIndex = 22
      Me.cmdClose.Text = "&Close"
      '
      'lblID
      '
      Me.lblID.AutoSize = True
      Me.lblID.Location = New System.Drawing.Point(96, 72)
      Me.lblID.Name = "lblID"
      Me.lblID.Size = New System.Drawing.Size(15, 13)
      Me.lblID.TabIndex = 8
      Me.lblID.Text = "ID"
      '
      'txtID
      '
      Me.txtID.Location = New System.Drawing.Point(121, 69)
      Me.txtID.Name = "txtID"
      Me.txtID.ReadOnly = True
      Me.txtID.Size = New System.Drawing.Size(64, 20)
      Me.txtID.TabIndex = 9
      Me.txtID.Text = ""
      Me.txtID.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      '
      'frmLaborItemsMaint
      '
      Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
      Me.ClientSize = New System.Drawing.Size(488, 272)
      Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdAdd, Me.cmdUpdate, Me.cmdDelete, Me.cmdClose, Me.GroupBox1, Me.dbgSaleItems})
      Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
      Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
      Me.MaximizeBox = False
      Me.MinimizeBox = False
      Me.Name = "frmLaborItemsMaint"
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "Maintain Labor Items"
      CType(Me.dbgSaleItems, System.ComponentModel.ISupportInitialize).EndInit()
      Me.GroupBox1.ResumeLayout(False)
      Me.ResumeLayout(False)

   End Sub

#End Region
#Region " Module Variables "
   Dim msAddorEdit As String
   Dim mbDirty As Boolean
   Dim mbFormLoading As Boolean
   Private oDA As CDataAccess
   Private iHitRow As Integer
   Private dtSI As New DataTable("dt")
   Private oCG As CGrid


#End Region


#Region " Form & Control Events "
    Private Sub dbgSALEitems_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dbgSaleItems.MouseUp
        Try
            Dim pt = New Point(e.X, e.Y)
            Dim hti As DataGrid.HitTestInfo = Me.dbgSaleItems.HitTest(pt)
            Me.dbgSaleItems.Select(hti.Row)
            iHitRow = hti.Row
            Me.LoadRates()
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
   Private Sub txtPricePerUnit_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPricePerUnit.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
   End Sub
   Private Sub txtPricePerUnit_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPricePerUnit.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub txtPricePerUnit_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPricePerUnit.Enter
      txtPricePerUnit.SelectionStart = 0
      txtPricePerUnit.SelectionLength = txtPricePerUnit.Text.Trim.Length
   End Sub
   Private Sub txtProductDescription_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtProductDescription.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
   End Sub
   Private Sub txtProductDescription_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtProductDescription.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub txtProductDescription_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProductDescription.Enter
      txtProductDescription.SelectionStart = 0
      txtProductDescription.SelectionLength = txtProductDescription.Text.Trim.Length
   End Sub

   Private Sub frmLaborItemsMaint_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
      Me.LoadTheGrid()
      Me.iHitRow = 0
      Me.LoadRates()
   End Sub


   Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
      Dim SQL As String
      Dim dt As New DataTable()

      If mbDirty Then
         If MsgBox("You have unsaved changes; do you want to add without saving your changes?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
            Exit Sub
         End If
      End If


      ClearTextBoxes()
      msAddorEdit = "A"

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
         SQL = "delete from labor_charges "
         SQL &= "where unique_id= " & Me.dtSI.Rows(Me.iHitRow).Item("unique_id") & " "
         iRows = oDA.SendActionSql(SQL, ConnectString, sErr)
         If iRows = 0 Then
            MsgBox("Delete of labor item failed.  " & Chr(10) & sErr, MsgBoxStyle.Critical)
            Exit Sub
         End If
         Me.LoadTheGrid()
         iHitRow = 0
         LoadRates()
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub


   Private Sub cmdUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdate.Click
      Dim SQL As String
      Dim sErr As String


      Try
         With Me
            If msAddorEdit = "A" Then
               SQL = "insert into  labor_charges"
               SQL &= "(labor_type,price) "
               SQL &= "values("
               SQL &= "'" & Replace(.txtProductDescription.Text, "'", "''") & "', "
               SQL &= UnFormat(.txtPricePerUnit.Text) & " "
               SQL &= ")"
            Else
                    SQL = "update labor_charges set"
               SQL &= "labor_type = '" & Replace(.txtProductDescription.Text, "'", "''") & "', "
               SQL &= "price = " & UnFormat(.txtPricePerUnit.Text) & " "
               SQL = SQL & "where unique_id = " & .txtID.Text & " "
            End If
         End With
         If oDA.SendActionSql(SQL, ConnectString, sErr) < 1 Then
            MsgBox("Update failed: " & Chr(10) & sErr, MsgBoxStyle.Critical)
         End If

         msAddorEdit = "E"
         LoadTheGrid()
         iHitRow = 0
         LoadRates()
         mbDirty = False
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try

   End Sub

#End Region


#Region " Private Methods "
   Private Sub LoadRates()
      Try

         With Me.dtSI.Rows(Me.iHitRow)
            If Not IsDBNull(.Item("Price")) Then
               Me.txtPricePerUnit.Text = FormatCurrency(.Item("Price"))
            Else
               Me.txtPricePerUnit.Text = "$0.00"
            End If

            Me.txtProductDescription.Text = IIf(IsDBNull(.Item("labor_type")), "", .Item("labor_type"))
            Me.txtID.Text = .Item("unique_id")
            mbDirty = False
         End With

      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try

   End Sub
   Private Sub LoadTheGrid()
      Dim SQL As String


      Try
         oCG.InitializeDatatableForStyles(dtSI)
         oCG.BindDataTableToGrid(dtSI, Me.dbgSaleItems)

         SQL = "select Labor_Type, "
         SQL = SQL & "Price, "
         SQL = SQL & "Unique_Id "
         SQL = SQL & "from labor_charges "
         SQL = SQL & " order by labor_type"
         oDA.SendQuery(SQL, dtSI, ConnectString, "dt")
         If dtSI.Rows.Count > 0 Then
            Dim Formats() As String = _
               {"", "200", "T", "L", _
               "$#,##0.00", "60", "T", "R", _
               "", "60", "T", "R"}
            oCG.SetTablesStyle(dtSI, Me.dbgSaleItems, Formats)
            oCG.BindDataTableToGrid(dtSI, Me.dbgSaleItems)
            oCG.DisableAddNew(Me.dbgSaleItems, Me)
         End If
         mbDirty = False
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub


   Private Sub ClearTextBoxes()
      With Me
         .txtPricePerUnit.Text = ""
         .txtProductDescription.Text = ""
         .txtID.Text = ""
      End With
   End Sub

#End Region



End Class
