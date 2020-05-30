Imports System.Windows.Forms.Application
Public Class frmSupplies
   Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

   Public Sub New()
      MyBase.New()

      'This call is required by the Windows Form Designer.
      InitializeComponent()

      'Add any initialization after the InitializeComponent() call
      oDA = New CDataAccess()
      oCG = New CGrid()
      colSaleItems = New Collection()
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
   Friend WithEvents dbgToolnSupplies As System.Windows.Forms.DataGrid
   Friend WithEvents Label1 As System.Windows.Forms.Label
   Friend WithEvents Label2 As System.Windows.Forms.Label
   Friend WithEvents Label3 As System.Windows.Forms.Label
   Friend WithEvents Label4 As System.Windows.Forms.Label
   Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
   Friend WithEvents txtPricePerUnit As System.Windows.Forms.TextBox
   Friend WithEvents txtExtendedPrice As System.Windows.Forms.TextBox
   Friend WithEvents cboQty As System.Windows.Forms.ComboBox
   Friend WithEvents cboDiscount As System.Windows.Forms.ComboBox
   Friend WithEvents lblPromoCode As System.Windows.Forms.Label
   Public WithEvents btnAddToCart As System.Windows.Forms.Button
   Public WithEvents btnSelectItem As System.Windows.Forms.Button
   Public WithEvents cmdCancel As System.Windows.Forms.Button
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
      Me.components = New System.ComponentModel.Container()
      Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmSupplies))
      Me.dbgToolnSupplies = New System.Windows.Forms.DataGrid()
      Me.cboQty = New System.Windows.Forms.ComboBox()
      Me.Label1 = New System.Windows.Forms.Label()
      Me.Label2 = New System.Windows.Forms.Label()
      Me.txtPricePerUnit = New System.Windows.Forms.TextBox()
      Me.Label3 = New System.Windows.Forms.Label()
      Me.txtExtendedPrice = New System.Windows.Forms.TextBox()
      Me.Label4 = New System.Windows.Forms.Label()
      Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
      Me.btnAddToCart = New System.Windows.Forms.Button()
      Me.btnSelectItem = New System.Windows.Forms.Button()
      Me.cmdCancel = New System.Windows.Forms.Button()
      Me.cboDiscount = New System.Windows.Forms.ComboBox()
      Me.lblPromoCode = New System.Windows.Forms.Label()
      CType(Me.dbgToolnSupplies, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.SuspendLayout()
      '
      'dbgToolnSupplies
      '
      Me.dbgToolnSupplies.AllowSorting = False
      Me.dbgToolnSupplies.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                  Or System.Windows.Forms.AnchorStyles.Left) _
                  Or System.Windows.Forms.AnchorStyles.Right)
      Me.dbgToolnSupplies.CaptionVisible = False
      Me.dbgToolnSupplies.DataMember = ""
      Me.dbgToolnSupplies.HeaderForeColor = System.Drawing.SystemColors.ControlText
      Me.dbgToolnSupplies.Location = New System.Drawing.Point(0, 19)
      Me.dbgToolnSupplies.Name = "dbgToolnSupplies"
      Me.dbgToolnSupplies.ReadOnly = True
      Me.dbgToolnSupplies.Size = New System.Drawing.Size(608, 223)
      Me.dbgToolnSupplies.TabIndex = 0
      '
      'cboQty
      '
      Me.cboQty.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
      Me.cboQty.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "30", "40", "50", "75", "100"})
      Me.cboQty.Location = New System.Drawing.Point(7, 264)
      Me.cboQty.Name = "cboQty"
      Me.cboQty.Size = New System.Drawing.Size(63, 21)
      Me.cboQty.TabIndex = 1
      Me.cboQty.Text = "1"
      '
      'Label1
      '
      Me.Label1.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
      Me.Label1.AutoSize = True
      Me.Label1.Location = New System.Drawing.Point(7, 251)
      Me.Label1.Name = "Label1"
      Me.Label1.Size = New System.Drawing.Size(46, 13)
      Me.Label1.TabIndex = 2
      Me.Label1.Text = "Quantity"
      '
      'Label2
      '
      Me.Label2.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
      Me.Label2.AutoSize = True
      Me.Label2.Location = New System.Drawing.Point(81, 252)
      Me.Label2.Name = "Label2"
      Me.Label2.Size = New System.Drawing.Size(53, 13)
      Me.Label2.TabIndex = 3
      Me.Label2.Text = "Unit Price"
      '
      'txtPricePerUnit
      '
      Me.txtPricePerUnit.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
      Me.txtPricePerUnit.Location = New System.Drawing.Point(81, 265)
      Me.txtPricePerUnit.Name = "txtPricePerUnit"
      Me.txtPricePerUnit.Size = New System.Drawing.Size(80, 20)
      Me.txtPricePerUnit.TabIndex = 4
      Me.txtPricePerUnit.Tag = "(No Auto Formatting)"
      Me.txtPricePerUnit.Text = ""
      Me.txtPricePerUnit.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      Me.ToolTip1.SetToolTip(Me.txtPricePerUnit, "Change this price manually to give a price break")
      '
      'Label3
      '
      Me.Label3.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
      Me.Label3.AutoSize = True
      Me.Label3.Location = New System.Drawing.Point(172, 252)
      Me.Label3.Name = "Label3"
      Me.Label3.Size = New System.Drawing.Size(81, 13)
      Me.Label3.TabIndex = 5
      Me.Label3.Text = "Extended Price"
      '
      'txtExtendedPrice
      '
      Me.txtExtendedPrice.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
      Me.txtExtendedPrice.Location = New System.Drawing.Point(172, 265)
      Me.txtExtendedPrice.Name = "txtExtendedPrice"
      Me.txtExtendedPrice.ReadOnly = True
      Me.txtExtendedPrice.Size = New System.Drawing.Size(80, 20)
      Me.txtExtendedPrice.TabIndex = 6
      Me.txtExtendedPrice.Text = ""
      Me.txtExtendedPrice.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      '
      'Label4
      '
      Me.Label4.AutoSize = True
      Me.Label4.Location = New System.Drawing.Point(3, 4)
      Me.Label4.Name = "Label4"
      Me.Label4.Size = New System.Drawing.Size(135, 13)
      Me.Label4.TabIndex = 7
      Me.Label4.Text = "Select Tools and Supplies"
      '
      'btnAddToCart
      '
      Me.btnAddToCart.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
      Me.btnAddToCart.BackColor = System.Drawing.SystemColors.Control
      Me.btnAddToCart.Cursor = System.Windows.Forms.Cursors.Default
      Me.btnAddToCart.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.btnAddToCart.ForeColor = System.Drawing.SystemColors.ControlText
      Me.btnAddToCart.ImageAlign = System.Drawing.ContentAlignment.TopCenter
      Me.btnAddToCart.Location = New System.Drawing.Point(449, 260)
      Me.btnAddToCart.Name = "btnAddToCart"
      Me.btnAddToCart.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.btnAddToCart.Size = New System.Drawing.Size(87, 26)
      Me.btnAddToCart.TabIndex = 37
      Me.btnAddToCart.TabStop = False
      Me.btnAddToCart.Text = "&Add to Cart"
      Me.ToolTip1.SetToolTip(Me.btnAddToCart, "Add selected items to car")
      '
      'btnSelectItem
      '
      Me.btnSelectItem.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
      Me.btnSelectItem.BackColor = System.Drawing.SystemColors.Control
      Me.btnSelectItem.Cursor = System.Windows.Forms.Cursors.Default
      Me.btnSelectItem.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.btnSelectItem.ForeColor = System.Drawing.SystemColors.ControlText
      Me.btnSelectItem.ImageAlign = System.Drawing.ContentAlignment.TopCenter
      Me.btnSelectItem.Location = New System.Drawing.Point(354, 260)
      Me.btnSelectItem.Name = "btnSelectItem"
      Me.btnSelectItem.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.btnSelectItem.Size = New System.Drawing.Size(87, 26)
      Me.btnSelectItem.TabIndex = 38
      Me.btnSelectItem.TabStop = False
      Me.btnSelectItem.Text = "&Select Item"
      Me.ToolTip1.SetToolTip(Me.btnSelectItem, "Select this item and select another if desired")
      '
      'cmdCancel
      '
      Me.cmdCancel.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
      Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
      Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
      Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
      Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.TopCenter
      Me.cmdCancel.Location = New System.Drawing.Point(543, 260)
      Me.cmdCancel.Name = "cmdCancel"
      Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.cmdCancel.Size = New System.Drawing.Size(55, 26)
      Me.cmdCancel.TabIndex = 39
      Me.cmdCancel.TabStop = False
      Me.cmdCancel.Text = "&Cancel"
      Me.ToolTip1.SetToolTip(Me.cmdCancel, "Cancel all items")
      '
      'cboDiscount
      '
      Me.cboDiscount.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "15", "20"})
      Me.cboDiscount.Location = New System.Drawing.Point(259, 264)
      Me.cboDiscount.Name = "cboDiscount"
      Me.cboDiscount.Size = New System.Drawing.Size(72, 21)
      Me.cboDiscount.TabIndex = 36
      Me.cboDiscount.Visible = False
      '
      'lblPromoCode
      '
      Me.lblPromoCode.AutoSize = True
      Me.lblPromoCode.Location = New System.Drawing.Point(262, 252)
      Me.lblPromoCode.Name = "lblPromoCode"
      Me.lblPromoCode.Size = New System.Drawing.Size(68, 13)
      Me.lblPromoCode.TabIndex = 35
      Me.lblPromoCode.Text = "Promo Code"
      Me.lblPromoCode.Visible = False
      '
      'frmSupplies
      '
      Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
      Me.ClientSize = New System.Drawing.Size(608, 302)
      Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdCancel, Me.btnSelectItem, Me.btnAddToCart, Me.cboDiscount, Me.lblPromoCode, Me.Label4, Me.Label3, Me.Label2, Me.Label1, Me.txtExtendedPrice, Me.txtPricePerUnit, Me.cboQty, Me.dbgToolnSupplies})
      Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
      Me.MaximizeBox = False
      Me.MinimizeBox = False
      Me.Name = "frmSupplies"
      Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "Select Tool and Supplies"
      CType(Me.dbgToolnSupplies, System.ComponentModel.ISupportInitialize).EndInit()
      Me.ResumeLayout(False)

   End Sub

#End Region

#Region " Module Variables "
   Private SQL As String
   Private oDA As CDataAccess
   Dim dtTS As DataTable
   Dim oCG As CGrid
   Dim iHitRow As Integer
   Private colSaleItems As Collection
   Private mbFormLoading As Boolean = True
   Private mbDirty As Boolean


#End Region

#Region " Private Methods "
   Private Sub LoadTheGrid()
      Dim i As Integer

      Try
         SQL = "select ProductId,ProductName,ProductDescription,"
         SQL &= "PricePerUnit,UnitsInStock,ReorderLevel "
         SQL &= "from Products order by productname"
         oCG.ClearDataTableForRebinding(dtTS)

         If oDA.SendQuery(SQL, dtTS, ConnectString, "dt") > 0 Then
            Dim formats() As String = _
               {"", "60", "T", "L", _
                "", "150", "T", "L", _
                "", "150", "T", "L", _
                "$#,##0.00", "60", "T", "R", _
                "", "60", "T", "R", _
                "", "80", "T", "R"}
            oCG.SetTablesStyle(dtTS, Me.dbgToolnSupplies, formats)
            oCG.BindDataTableToGrid(dtTS, Me.dbgToolnSupplies)
            oCG.DisableAddNew(Me.dbgToolnSupplies, Me)
         End If
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

   Private Sub Recompute()
      With Me
         .txtExtendedPrice.Text = FormatCurrency(UnFormat(Me.txtPricePerUnit.Text) * Val(Me.cboQty.Text))
      End With
   End Sub

   Private Sub LoadTextBoxes()
      Try
         With Me
            Me.txtPricePerUnit.Text = FormatCurrency(dtTS.Rows(iHitRow).Item("priceperunit"))
            Me.txtExtendedPrice.Text = FormatCurrency(dtTS.Rows(iHitRow).Item("priceperunit") * Val(Me.cboQty.Text))
         End With
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

#End Region

#Region " Form & Control Events "
   ''' <summary>
   '''
   ''' </summary>
   ''' <param name = "sender"></param>
   ''' <param name = "e"></param>
   Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

      If mbDirty Then
         Dim sMsg As String
         Dim iRV As Integer
         sMsg = "You have selected items to sell.  Are you sure" & Chr(10)
         sMsg &= "you want to cancel the sale items?" & Chr(10)
         sMsg &= "" & Chr(10)
         sMsg &= "Click Yes to Cancel the sale of the selected items." & Chr(10)
         sMsg &= "Click No to leave the form displayed." & Chr(10)
         sMsg &= "" & Chr(10)
         iRV = MsgBox(sMsg, CType(36, Microsoft.VisualBasic.MsgBoxStyle), "Confirm Cancel")

         If iRV = 6 Then
            ' Yes Code goes here
         Else
            ' No code goes here
            Exit Sub
         End If
      End If
      Me.Close()
      DoEvents()
   End Sub

   Private Sub frmSupplies_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
      LoadTheGrid()
      iHitRow = 0
      LoadTextBoxes()
      mbFormLoading = False
   End Sub


   Private Sub txtPricePerUnit_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPricePerUnit.Leave
      With Me
         txtPricePerUnit.Text = FormatCurrency(txtPricePerUnit.Text)
         Recompute()
      End With
   End Sub

   Private Sub cboQty_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboQty.Leave
      If Not mbFormLoading Then
         'Me.LoadTextBoxes()
         Me.Recompute()
      End If
   End Sub


   Private Sub dbgToolnSupplies_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dbgToolnSupplies.MouseUp
      iHitRow = oCG.GetClickedRow(e, Me.dbgToolnSupplies)
      Try
         dbgToolnSupplies.Select(iHitRow)
         LoadTextBoxes()
      Catch
      End Try
   End Sub

   Private Sub btnAddSelectedItems_Click(ByVal sender As System.Object, _
      ByVal e As System.EventArgs) Handles btnSelectItem.Click
      Dim o As New CToolsAndSupplies()
      Dim discount As Decimal = Val(Me.cboDiscount.Text)
      Dim shtUnitsAvail As Short

      Try
         shtUnitsAvail = dtTS.Rows(iHitRow).Item("UnitsInStock")
         With o
            If Val(Me.cboQty.Text) > shtUnitsAvail Then
               MsgBox("There are only " & shtUnitsAvail & " available!", MsgBoxStyle.Information)
               Exit Sub
            End If
            .ItemCount = Val(Me.cboQty.Text)
            .ItemId = dtTS.Rows(iHitRow).Item("Productid")
            .ItemPrice = UnFormat(Me.txtPricePerUnit.Text) '* discount
            .ItemExtendedPrice = UnFormat(Me.txtExtendedPrice.Text)
            .Deposit = 0
            .RentOrSale = SALE
            .ItemName = dtTS.Rows(iHitRow).Item("productname")
         End With
         Me.colSaleItems.Add(o)
         Me.cboQty.Text = "1"
         mbDirty = True
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub


   Private Sub btnAddItemsAndClose_Click(ByVal sender As System.Object, _
     ByVal e As System.EventArgs) Handles btnAddToCart.Click


      Try
         Dim serr As String = ""
         With Me
            If .colSaleItems.Count > 0 Then
               Dim o As New CToolsAndSupplies()
               Dim i As Integer
               For i = 1 To Me.colSaleItems.Count
                  o = colSaleItems(i)

                  SQL = "insert into tempitems "
                  SQL &= "(itemid,itemname,itemcount,itemprice,itemextendedprice,"
                  SQL &= "rentorsale,itemdeposit,itemperiod,meter_required,hour_meter,user_id) "
                  SQL &= "values('" & o.ItemId & "', "
                  SQL &= "'" & Replace(o.ItemName, "'", "''") & "', "
                  SQL &= o.ItemCount.ToString & ", "
                  SQL &= o.ItemPrice.ToString & ", "
                  SQL &= o.ItemExtendedPrice.ToString & ", "
                  SQL &= "'" & SALE & "', "
                  SQL &= o.Deposit.ToString & ",'" & SALE & "', "
                  SQL &= False & ", 0, "
                  SQL &= "'" & UserName & "' "
                  SQL &= ")"
                  If oDA.SendActionSql(SQL, ConnectString, serr) < 1 Then
                     MsgBox("Can't add item to temp table, " & serr, MsgBoxStyle.Critical)
                     Exit Sub
                  End If
               Next i
            End If
         End With
         Me.Close()
         DoEvents()
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub


   Private Sub cboQty_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboQty.SelectedIndexChanged
      If Not mbFormLoading Then
         'Me.LoadTextBoxes()
         Me.Recompute()
      End If
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


#End Region

End Class
