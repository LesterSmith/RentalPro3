Public Class frmSaleItemsMaintenance
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
   Friend WithEvents txtProductID As System.Windows.Forms.TextBox
   Friend WithEvents txtProductName As System.Windows.Forms.TextBox
   Friend WithEvents txtProductDescription As System.Windows.Forms.TextBox
   Friend WithEvents txtPricePerUnit As System.Windows.Forms.TextBox
   Friend WithEvents lblProductId As System.Windows.Forms.Label
   Friend WithEvents lblProductName As System.Windows.Forms.Label
   Friend WithEvents lblProductDesc As System.Windows.Forms.Label
   Friend WithEvents Label4 As System.Windows.Forms.Label
   Public WithEvents cmdAdd As System.Windows.Forms.Button
   Public WithEvents cmdUpdate As System.Windows.Forms.Button
   Public WithEvents cmdDelete As System.Windows.Forms.Button
   Public WithEvents cmdClose As System.Windows.Forms.Button
   Friend WithEvents dbgSaleItems As System.Windows.Forms.DataGrid
   Friend WithEvents lblUIS As System.Windows.Forms.Label
   Friend WithEvents txtUnitsInStock As System.Windows.Forms.TextBox
   Friend WithEvents lblRL As System.Windows.Forms.Label
   Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
   Public WithEvents cmdRefresh As System.Windows.Forms.Button
   Friend WithEvents chkRefresh As System.Windows.Forms.CheckBox
   Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
      Me.components = New System.ComponentModel.Container()
      Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmSaleItemsMaintenance))
      Me.dbgSaleItems = New System.Windows.Forms.DataGrid()
      Me.GroupBox1 = New System.Windows.Forms.GroupBox()
      Me.lblRL = New System.Windows.Forms.Label()
      Me.TextBox1 = New System.Windows.Forms.TextBox()
      Me.lblUIS = New System.Windows.Forms.Label()
      Me.txtUnitsInStock = New System.Windows.Forms.TextBox()
      Me.Label4 = New System.Windows.Forms.Label()
      Me.lblProductDesc = New System.Windows.Forms.Label()
      Me.lblProductName = New System.Windows.Forms.Label()
      Me.lblProductId = New System.Windows.Forms.Label()
      Me.txtPricePerUnit = New System.Windows.Forms.TextBox()
      Me.txtProductDescription = New System.Windows.Forms.TextBox()
      Me.txtProductName = New System.Windows.Forms.TextBox()
      Me.txtProductID = New System.Windows.Forms.TextBox()
      Me.cmdAdd = New System.Windows.Forms.Button()
      Me.cmdUpdate = New System.Windows.Forms.Button()
      Me.cmdDelete = New System.Windows.Forms.Button()
      Me.cmdClose = New System.Windows.Forms.Button()
      Me.cmdRefresh = New System.Windows.Forms.Button()
      Me.chkRefresh = New System.Windows.Forms.CheckBox()
      Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
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
      Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.lblRL, Me.TextBox1, Me.lblUIS, Me.txtUnitsInStock, Me.Label4, Me.lblProductDesc, Me.lblProductName, Me.lblProductId, Me.txtPricePerUnit, Me.txtProductDescription, Me.txtProductName, Me.txtProductID})
      Me.GroupBox1.Location = New System.Drawing.Point(8, 137)
      Me.GroupBox1.Name = "GroupBox1"
      Me.GroupBox1.Size = New System.Drawing.Size(320, 183)
      Me.GroupBox1.TabIndex = 18
      Me.GroupBox1.TabStop = False
      Me.GroupBox1.Text = "Selected Items"
      '
      'lblRL
      '
      Me.lblRL.AutoSize = True
      Me.lblRL.Location = New System.Drawing.Point(42, 139)
      Me.lblRL.Name = "lblRL"
      Me.lblRL.Size = New System.Drawing.Size(75, 13)
      Me.lblRL.TabIndex = 11
      Me.lblRL.Text = "Reorder Level"
      '
      'TextBox1
      '
      Me.TextBox1.Location = New System.Drawing.Point(122, 139)
      Me.TextBox1.Name = "TextBox1"
      Me.TextBox1.TabIndex = 10
      Me.TextBox1.Tag = "(No Auto Formatting)"
      Me.TextBox1.Text = ""
      Me.TextBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      '
      'lblUIS
      '
      Me.lblUIS.AutoSize = True
      Me.lblUIS.Location = New System.Drawing.Point(44, 117)
      Me.lblUIS.Name = "lblUIS"
      Me.lblUIS.Size = New System.Drawing.Size(74, 13)
      Me.lblUIS.TabIndex = 9
      Me.lblUIS.Text = "Units In Stock"
      '
      'txtUnitsInStock
      '
      Me.txtUnitsInStock.Location = New System.Drawing.Point(122, 115)
      Me.txtUnitsInStock.Name = "txtUnitsInStock"
      Me.txtUnitsInStock.TabIndex = 8
      Me.txtUnitsInStock.Tag = "(No Auto Formatting)"
      Me.txtUnitsInStock.Text = ""
      Me.txtUnitsInStock.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      '
      'Label4
      '
      Me.Label4.AutoSize = True
      Me.Label4.Location = New System.Drawing.Point(44, 92)
      Me.Label4.Name = "Label4"
      Me.Label4.Size = New System.Drawing.Size(74, 13)
      Me.Label4.TabIndex = 7
      Me.Label4.Text = "Price Per Unit"
      '
      'lblProductDesc
      '
      Me.lblProductDesc.AutoSize = True
      Me.lblProductDesc.Location = New System.Drawing.Point(13, 68)
      Me.lblProductDesc.Name = "lblProductDesc"
      Me.lblProductDesc.Size = New System.Drawing.Size(103, 13)
      Me.lblProductDesc.TabIndex = 6
      Me.lblProductDesc.Text = "Product Description"
      '
      'lblProductName
      '
      Me.lblProductName.AutoSize = True
      Me.lblProductName.Location = New System.Drawing.Point(40, 44)
      Me.lblProductName.Name = "lblProductName"
      Me.lblProductName.Size = New System.Drawing.Size(76, 13)
      Me.lblProductName.TabIndex = 5
      Me.lblProductName.Text = "Product Name"
      '
      'lblProductId
      '
      Me.lblProductId.AutoSize = True
      Me.lblProductId.Location = New System.Drawing.Point(58, 20)
      Me.lblProductId.Name = "lblProductId"
      Me.lblProductId.Size = New System.Drawing.Size(58, 13)
      Me.lblProductId.TabIndex = 4
      Me.lblProductId.Text = "Product ID"
      '
      'txtPricePerUnit
      '
      Me.txtPricePerUnit.Location = New System.Drawing.Point(122, 90)
      Me.txtPricePerUnit.Name = "txtPricePerUnit"
      Me.txtPricePerUnit.TabIndex = 3
      Me.txtPricePerUnit.Tag = "(No Auto Formatting)"
      Me.txtPricePerUnit.Text = ""
      Me.txtPricePerUnit.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      '
      'txtProductDescription
      '
      Me.txtProductDescription.Location = New System.Drawing.Point(122, 65)
      Me.txtProductDescription.Name = "txtProductDescription"
      Me.txtProductDescription.Size = New System.Drawing.Size(188, 20)
      Me.txtProductDescription.TabIndex = 2
      Me.txtProductDescription.Tag = "(No Auto Formatting)"
      Me.txtProductDescription.Text = ""
      '
      'txtProductName
      '
      Me.txtProductName.Location = New System.Drawing.Point(122, 40)
      Me.txtProductName.Name = "txtProductName"
      Me.txtProductName.Size = New System.Drawing.Size(188, 20)
      Me.txtProductName.TabIndex = 1
      Me.txtProductName.Tag = "(No Auto Formatting)"
      Me.txtProductName.Text = ""
      '
      'txtProductID
      '
      Me.txtProductID.Location = New System.Drawing.Point(122, 16)
      Me.txtProductID.Name = "txtProductID"
      Me.txtProductID.TabIndex = 0
      Me.txtProductID.Tag = "(No Auto Formatting)"
      Me.txtProductID.Text = ""
      '
      'cmdAdd
      '
      Me.cmdAdd.BackColor = System.Drawing.SystemColors.Control
      Me.cmdAdd.Cursor = System.Windows.Forms.Cursors.Default
      Me.cmdAdd.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.cmdAdd.ForeColor = System.Drawing.SystemColors.ControlText
      Me.cmdAdd.Location = New System.Drawing.Point(384, 195)
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
      Me.cmdUpdate.Location = New System.Drawing.Point(384, 163)
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
      Me.cmdDelete.Location = New System.Drawing.Point(384, 227)
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
      Me.cmdClose.Location = New System.Drawing.Point(384, 291)
      Me.cmdClose.Name = "cmdClose"
      Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.cmdClose.Size = New System.Drawing.Size(77, 26)
      Me.cmdClose.TabIndex = 22
      Me.cmdClose.Text = "&Close"
      '
      'cmdRefresh
      '
      Me.cmdRefresh.BackColor = System.Drawing.SystemColors.Control
      Me.cmdRefresh.Cursor = System.Windows.Forms.Cursors.Default
      Me.cmdRefresh.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.cmdRefresh.ForeColor = System.Drawing.SystemColors.ControlText
      Me.cmdRefresh.Location = New System.Drawing.Point(384, 259)
      Me.cmdRefresh.Name = "cmdRefresh"
      Me.cmdRefresh.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.cmdRefresh.Size = New System.Drawing.Size(77, 26)
      Me.cmdRefresh.TabIndex = 23
      Me.cmdRefresh.Text = "&Refresh"
      Me.ToolTip1.SetToolTip(Me.cmdRefresh, "Refresh the grid with all changes")
      Me.cmdRefresh.Visible = False
      '
      'chkRefresh
      '
      Me.chkRefresh.Location = New System.Drawing.Point(352, 136)
      Me.chkRefresh.Name = "chkRefresh"
      Me.chkRefresh.Size = New System.Drawing.Size(120, 16)
      Me.chkRefresh.TabIndex = 24
      Me.chkRefresh.Text = "Refresh on Update"
      Me.ToolTip1.SetToolTip(Me.chkRefresh, "Refresh grid with each update, moves grid to top row")
      '
      'frmSaleItemsMaintenance
      '
      Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
      Me.ClientSize = New System.Drawing.Size(488, 334)
      Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.chkRefresh, Me.cmdRefresh, Me.cmdAdd, Me.cmdUpdate, Me.cmdDelete, Me.cmdClose, Me.GroupBox1, Me.dbgSaleItems})
      Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
      Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
      Me.MaximizeBox = False
      Me.MinimizeBox = False
      Me.Name = "frmSaleItemsMaintenance"
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "Maintain Sale Items"
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


#Region " Private Methods "
   Private Sub LoadTheGrid()
      Dim SQL As String


      Try
         oCG.InitializeDatatableForStyles(dtSI)
         oCG.BindDataTableToGrid(dtSI, Me.dbgSaleItems)

         SQL = "select ProductID, "
         SQL = SQL & "ProductName, "
         SQL = SQL & "ProductDescription, "
         SQL = SQL & "PricePerUnit,UnitsInStock,ReorderLevel "
         SQL = SQL & "from products "
         SQL = SQL & " order by ProductName"
         oDA.SendQuery(SQL, dtSI, ConnectString, "dt")
         If dtSI.Rows.Count > 0 Then
            Dim Formats() As String = _
               {"", "60", "T", "L", _
               "", "100", "T", "L", _
               "", "100", "T", "L", _
               "$#,##0.00", "60", "T", "R", _
               "", "60", "T", "R", _
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

   Private Sub LoadRates()
      Try

         With Me.dtSI.Rows(Me.iHitRow)
            If Not IsDBNull(.Item("Priceperunit")) Then
               Me.txtPricePerUnit.Text = FormatCurrency(.Item("Priceperunit"))
            Else
               Me.txtPricePerUnit.Text = "$0.00"
            End If

            Me.txtProductDescription.Text = IIf(IsDBNull(.Item("productdescription")), "", .Item("productdescription"))
            Me.txtProductID.Text = IIf(IsDBNull(.Item("productid")), "", .Item("productid"))
            Me.txtProductName.Text = IIf(IsDBNull(.Item("productname")), "", .Item("productname"))
            Me.txtUnitsInStock.Text = IIf(IsDBNull(.Item("UnitsInStock")), "", .Item("UnitsInStock"))
            Me.TextBox1.Text = IIf(IsDBNull(.Item("ReorderLevel")), 0, .Item("ReorderLevel"))
            mbDirty = False
         End With

      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try

   End Sub

#End Region


#Region " Form & Control Events "

   Private Sub cmdUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdate.Click
      Dim SQL As String
      Dim sErr As String = String.Empty


      Try
         With Me
            If msAddorEdit = "A" Then
               SQL = "insert into products "
               SQL &= "(productid, productname, productdescription, priceperunit,unitsinstock,reorderlevel) "
               SQL &= "values('"
               SQL &= .txtProductID.Text & "', "
               SQL &= "'" & Replace(.txtProductName.Text, "'", "''") & "', "
               SQL &= "'" & .txtProductDescription.Text & "', "
               SQL &= UnFormat(.txtPricePerUnit.Text) & ", "
               SQL &= .txtUnitsInStock.Text & ", "
               SQL &= .TextBox1.Text & ")"
            Else
               SQL = "update products "
               SQL &= "set productname = '" & Replace(.txtProductName.Text, "'", "''") & "', "
               SQL &= "productdescription = '" & .txtProductDescription.Text & "', "
               SQL &= "priceperunit = " & UnFormat(.txtPricePerUnit.Text) & ", "
               SQL &= "Unitsinstock = " & .txtUnitsInStock.Text & ", "
               SQL &= "reorderlevel = " & .TextBox1.Text & " "
               SQL = SQL & "where productid = '" & .txtProductID.Text & "'"
            End If
         End With
         If oDA.SendActionSql(SQL, ConnectString, sErr) < 1 Then
            If sErr.Length > 0 Then
               MsgBox("Update failed: " & Chr(10) & sErr, MsgBoxStyle.Critical)
               WriteErrLog(sErr)
            Else
               MsgBox("No records updated, product ID was not found.", MsgBoxStyle.Exclamation)
            End If
         End If
         msAddorEdit = "E"
         iHitRow = 0
         mbDirty = False
         If Me.chkRefresh.Checked Then
            LoadTheGrid()
            LoadRates()
         End If
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try

   End Sub

   Private Sub ClearTextBoxes()
      With Me
         .txtPricePerUnit.Text = ""
         .txtProductDescription.Text = ""
         .txtProductID.Text = ""
         .txtProductName.Text = ""
      End With
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
         SQL = "delete from products "
         SQL &= "where productid= '" & Me.dtSI.Rows(Me.iHitRow).Item("productid") & "' "
         iRows = oDA.SendActionSql(SQL, ConnectString, sErr)
         If iRows = 0 Then
            MsgBox("Delete of equipment item failed.  " & Chr(10) & sErr, MsgBoxStyle.Critical)
            Exit Sub
         End If
         Me.LoadTheGrid()
         iHitRow = 0
         LoadRates()
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

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
   Private Sub txtProductName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtProductName.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
   End Sub
   Private Sub txtProductName_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtProductName.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub txtProductName_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProductName.Enter
      txtProductName.SelectionStart = 0
      txtProductName.SelectionLength = txtProductName.Text.Trim.Length
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
   Private Sub txtProductID_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtProductID.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
   End Sub
   Private Sub txtProductID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtProductID.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub txtProductID_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProductID.Enter
      txtProductID.SelectionStart = 0
      txtProductID.SelectionLength = txtProductID.Text.Trim.Length
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

   Private Sub frmSaleItemsMaintenance_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
      Me.LoadTheGrid()
      Me.iHitRow = 0
      Me.LoadRates()
   End Sub
   Private Sub txtUnitsInStock_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtUnitsInStock.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
   End Sub
   Private Sub txtUnitsInStock_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUnitsInStock.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub txtUnitsInStock_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUnitsInStock.Enter
      txtUnitsInStock.SelectionStart = 0
      txtUnitsInStock.SelectionLength = txtUnitsInStock.Text.Trim.Length
   End Sub
   Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
   End Sub
   Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub TextBox1_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.Enter
      TextBox1.SelectionStart = 0
      TextBox1.SelectionLength = TextBox1.Text.Trim.Length
   End Sub
   Private Sub cmdRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRefresh.Click
      LoadTheGrid()
      LoadRates()

   End Sub

   Private Sub chkRefresh_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkRefresh.CheckedChanged
      If chkRefresh.Checked Then
         Me.cmdRefresh.Enabled = False
         LoadTheGrid()
         LoadRates()
      Else
         Me.cmdRefresh.Enabled = True
      End If
   End Sub


#End Region

End Class
