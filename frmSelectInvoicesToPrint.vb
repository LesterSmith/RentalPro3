Imports System.Windows.Forms.Application
Public Class frmSelectInvoicesToPrint
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
   Friend WithEvents btnCancel As System.Windows.Forms.Button
   Friend WithEvents CustID As System.Windows.Forms.ColumnHeader
   Friend WithEvents Customer As System.Windows.Forms.ColumnHeader
   Friend WithEvents dbgEquipment As System.Windows.Forms.DataGrid
   Friend WithEvents Label1 As System.Windows.Forms.Label
   Friend WithEvents Label2 As System.Windows.Forms.Label
   Friend WithEvents lvInvoices As System.Windows.Forms.ListView
   Friend WithEvents btnPrint As System.Windows.Forms.Button
   Friend WithEvents chkSelectAllInvoices As System.Windows.Forms.CheckBox
   Friend WithEvents dtpStDate As System.Windows.Forms.DateTimePicker
   Friend WithEvents lblStDate As System.Windows.Forms.Label
   Friend WithEvents lblEndDate As System.Windows.Forms.Label
   Friend WithEvents grpInvoicesToShow As System.Windows.Forms.GroupBox
   Friend WithEvents optAllInvoices As System.Windows.Forms.RadioButton
   Friend WithEvents optOpenItems As System.Windows.Forms.RadioButton
   Friend WithEvents dtpEndDate As System.Windows.Forms.DateTimePicker
   Friend WithEvents btnPreview As System.Windows.Forms.Button
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSelectInvoicesToPrint))
        Me.dbgEquipment = New System.Windows.Forms.DataGrid()
        Me.lvInvoices = New System.Windows.Forms.ListView()
        Me.Customer = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.CustID = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnPrint = New System.Windows.Forms.Button()
        Me.chkSelectAllInvoices = New System.Windows.Forms.CheckBox()
        Me.dtpStDate = New System.Windows.Forms.DateTimePicker()
        Me.dtpEndDate = New System.Windows.Forms.DateTimePicker()
        Me.lblStDate = New System.Windows.Forms.Label()
        Me.lblEndDate = New System.Windows.Forms.Label()
        Me.grpInvoicesToShow = New System.Windows.Forms.GroupBox()
        Me.optAllInvoices = New System.Windows.Forms.RadioButton()
        Me.optOpenItems = New System.Windows.Forms.RadioButton()
        Me.btnPreview = New System.Windows.Forms.Button()
        CType(Me.dbgEquipment, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpInvoicesToShow.SuspendLayout()
        Me.SuspendLayout()
        '
        'dbgEquipment
        '
        Me.dbgEquipment.AllowSorting = False
        Me.dbgEquipment.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dbgEquipment.DataMember = ""
        Me.dbgEquipment.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dbgEquipment.Location = New System.Drawing.Point(8, 230)
        Me.dbgEquipment.Name = "dbgEquipment"
        Me.dbgEquipment.Size = New System.Drawing.Size(479, 169)
        Me.dbgEquipment.TabIndex = 0
        '
        'lvInvoices
        '
        Me.lvInvoices.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvInvoices.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.Customer, Me.CustID})
        Me.lvInvoices.FullRowSelect = True
        Me.lvInvoices.GridLines = True
        Me.lvInvoices.HideSelection = False
        Me.lvInvoices.Location = New System.Drawing.Point(8, 56)
        Me.lvInvoices.MultiSelect = False
        Me.lvInvoices.Name = "lvInvoices"
        Me.lvInvoices.Size = New System.Drawing.Size(479, 152)
        Me.lvInvoices.TabIndex = 1
        Me.lvInvoices.UseCompatibleStateImageBehavior = False
        Me.lvInvoices.View = System.Windows.Forms.View.Details
        '
        'Customer
        '
        Me.Customer.Text = "Customer"
        Me.Customer.Width = 204
        '
        'CustID
        '
        Me.CustID.Text = "Cust ID"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 38)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(75, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Select Invoice"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(13, 214)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(110, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Equipment on Invoice"
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(413, 408)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(64, 28)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "&Cancel"
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnPrint.Location = New System.Drawing.Point(341, 408)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(64, 28)
        Me.btnPrint.TabIndex = 5
        Me.btnPrint.Text = "&Print"
        '
        'chkSelectAllInvoices
        '
        Me.chkSelectAllInvoices.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.chkSelectAllInvoices.Location = New System.Drawing.Point(21, 408)
        Me.chkSelectAllInvoices.Name = "chkSelectAllInvoices"
        Me.chkSelectAllInvoices.Size = New System.Drawing.Size(136, 24)
        Me.chkSelectAllInvoices.TabIndex = 7
        Me.chkSelectAllInvoices.Text = "Select All Invoices"
        '
        'dtpStDate
        '
        Me.dtpStDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpStDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpStDate.Location = New System.Drawing.Point(240, 3)
        Me.dtpStDate.Name = "dtpStDate"
        Me.dtpStDate.Size = New System.Drawing.Size(96, 20)
        Me.dtpStDate.TabIndex = 10
        '
        'dtpEndDate
        '
        Me.dtpEndDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpEndDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpEndDate.Location = New System.Drawing.Point(240, 28)
        Me.dtpEndDate.Name = "dtpEndDate"
        Me.dtpEndDate.Size = New System.Drawing.Size(96, 20)
        Me.dtpEndDate.TabIndex = 11
        '
        'lblStDate
        '
        Me.lblStDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblStDate.AutoSize = True
        Me.lblStDate.Location = New System.Drawing.Point(177, 7)
        Me.lblStDate.Name = "lblStDate"
        Me.lblStDate.Size = New System.Drawing.Size(55, 13)
        Me.lblStDate.TabIndex = 12
        Me.lblStDate.Text = "Start Date"
        '
        'lblEndDate
        '
        Me.lblEndDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblEndDate.AutoSize = True
        Me.lblEndDate.Location = New System.Drawing.Point(177, 30)
        Me.lblEndDate.Name = "lblEndDate"
        Me.lblEndDate.Size = New System.Drawing.Size(52, 13)
        Me.lblEndDate.TabIndex = 13
        Me.lblEndDate.Text = "End Date"
        '
        'grpInvoicesToShow
        '
        Me.grpInvoicesToShow.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpInvoicesToShow.Controls.Add(Me.optAllInvoices)
        Me.grpInvoicesToShow.Controls.Add(Me.optOpenItems)
        Me.grpInvoicesToShow.Location = New System.Drawing.Point(344, 0)
        Me.grpInvoicesToShow.Name = "grpInvoicesToShow"
        Me.grpInvoicesToShow.Size = New System.Drawing.Size(144, 56)
        Me.grpInvoicesToShow.TabIndex = 14
        Me.grpInvoicesToShow.TabStop = False
        Me.grpInvoicesToShow.Text = "Invoices to Show"
        '
        'optAllInvoices
        '
        Me.optAllInvoices.Location = New System.Drawing.Point(7, 34)
        Me.optAllInvoices.Name = "optAllInvoices"
        Me.optAllInvoices.Size = New System.Drawing.Size(104, 16)
        Me.optAllInvoices.TabIndex = 11
        Me.optAllInvoices.Text = "All Invoices"
        '
        'optOpenItems
        '
        Me.optOpenItems.Checked = True
        Me.optOpenItems.Location = New System.Drawing.Point(7, 17)
        Me.optOpenItems.Name = "optOpenItems"
        Me.optOpenItems.Size = New System.Drawing.Size(112, 14)
        Me.optOpenItems.TabIndex = 10
        Me.optOpenItems.TabStop = True
        Me.optOpenItems.Text = "Open Invoices"
        '
        'btnPreview
        '
        Me.btnPreview.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnPreview.Location = New System.Drawing.Point(269, 408)
        Me.btnPreview.Name = "btnPreview"
        Me.btnPreview.Size = New System.Drawing.Size(64, 28)
        Me.btnPreview.TabIndex = 15
        Me.btnPreview.Text = "&Preview"
        '
        'frmSelectInvoicesToPrint
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(496, 446)
        Me.Controls.Add(Me.grpInvoicesToShow)
        Me.Controls.Add(Me.lblEndDate)
        Me.Controls.Add(Me.lblStDate)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.dtpEndDate)
        Me.Controls.Add(Me.dtpStDate)
        Me.Controls.Add(Me.lvInvoices)
        Me.Controls.Add(Me.chkSelectAllInvoices)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnPreview)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.dbgEquipment)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimizeBox = False
        Me.Name = "frmSelectInvoicesToPrint"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Select Invoices to Print"
        Me.TopMost = True
        CType(Me.dbgEquipment, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpInvoicesToShow.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region
#Region " Module Variables "
   Private oDA As New CDataAccess()
   Private m_SelectedInvoice As Integer
   Private mbWait As Boolean
   Private SQL As String
   Private dt As New DataTable()
   Private m_ShowAll As Boolean
   Dim oCG As New CGrid()
   Private Noise As Boolean
   Private miHitRow As Integer
   Public dtInv As New DataTable()


#End Region


#Region " Form & Control Events "
   Private Sub lvInvoices_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvInvoices.SelectedIndexChanged
      FillInvoiceGrid()
   End Sub
 


   Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
      Me.Close()
      DoEvents()
   End Sub

   Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
      If Me.lvInvoices.SelectedItems.Count > 0 Then
            '#If CustomerApp = RELIABLE Then
            '         Dim ocp As New CReliablePrint.CReliablePrint()
            '#Else
            Dim oCP As New CPrintInvoices()
            '#End If

         ocp.Preview = False
         ocp.PrintSelectedInvoices(Me)
      Else
         MsgBox("You have not selected a customer's invoices to print.", MsgBoxStyle.Exclamation)
      End If
   End Sub

   Private Sub dbgEquipment_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dbgEquipment.MouseUp
      Try
         Dim b As Boolean
         If Noise Then Exit Sub
         Noise = True
         miHitRow = oCG.SelectCkBoxRow(dtInv, Me.dbgEquipment, e, "Print", b)
         Noise = False
      Catch ex As System.Exception
      End Try
   End Sub

   Private Sub chkSelectAllInvoices_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSelectAllInvoices.CheckedChanged
      Dim i As Integer
      If Me.chkSelectAllInvoices.Checked Then
         Me.chkSelectAllInvoices.Text = "Unselect All"
      Else
         Me.chkSelectAllInvoices.Text = "Select All Invoices"
      End If
      Try
         For i = 0 To dtInv.Rows.Count - 1
            With dtInv.Rows(i)
               .Item("Print") = IIf(Me.chkSelectAllInvoices.Checked, "true", "false")
            End With
         Next
         If chkSelectAllInvoices.Checked = False Then
            FillInvoiceGrid()
         End If
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

   Private Sub optAllInvoices_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optAllInvoices.CheckedChanged
      FillInvoiceGrid()
   End Sub

   Private Sub btnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPreview.Click
      If Me.lvInvoices.SelectedItems.Count > 0 Then
            '#If CustomerApp = RELIABLE Then
            '         Dim ocp As New CReliablePrint.CReliablePrint()
            '#Else
            Dim oCP As New CPrintInvoices()
            '#End If
         ocp.Preview = True
         ocp.PrintSelectedInvoices(Me)
      Else
         MsgBox("You have not selected a customer's invoices to print.", MsgBoxStyle.Exclamation)
      End If
   End Sub

   Private Sub dbgEquipment_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs) Handles dbgEquipment.Navigate

   End Sub


   Private Sub frmSelectInvoicesToPrint_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
      ' load the invoice listview with Customer Name, CustomerID, InvoiceID
      ' Load the grid with equip items on the selected invoice  
      Dim i As Integer
      Dim iInvID As Integer
      Me.dtpStDate.Value = DateValue(Today)
      Me.dtpEndDate.Value = DateValue(Today)

      SQL &= "select distinct c.companyname,c.customerid  "
      SQL &= "from customers c, invoices i   "
      SQL &= "where c.customerid = i.customerid "
      SQL &= "order by companyname  "

      Dim iRows As Integer = oDA.SendQuery(SQL, dt, ConnectString)
      If iRows > 0 Then
         iInvID = dt.Rows(0).Item("customerid")
         For i = 0 To dt.Rows.Count - 1
            With dt.Rows(i)
               lvInvoices.Items.Add(.Item("companyname"))
               lvInvoices.Items(i).SubItems.Add(.Item("customerid"))
            End With
         Next
      End If
   End Sub

#End Region

#Region " Private Methods "


   Private Sub FillInvoiceGrid()
      Dim id As Integer
      Dim iRows As Integer


      Try
         With lvInvoices
            Try
               id = Val(lvInvoices.SelectedItems(0).SubItems(1).Text)
            Catch
               Exit Sub
            End Try
         End With
         SQL = "select * "
         SQL &= "from invoices "
         SQL &= "where customerid = " & id.ToString & " "
         SQL &= "and invoicedate >= #" & Me.dtpStDate.Value & "# "
         SQL &= "and invoicedate < #" & DateAdd(DateInterval.Day, 1, Me.dtpEndDate.Value) & "# "
         If Me.optOpenItems.Checked Then
            SQL &= "and status = 'OPEN' "
         End If
         'SQL &= "order by invoiceid "
         dtInv = New DataTable("dt")
         oCG.BindDataTableToGrid(dtInv, Me.dbgEquipment)
         'oDA.SendQuery(dtInv, SQL, ConnectString)
         If oDA.SendQuery(SQL, dtInv, ConnectString, "dt") > 0 Then
            oCG.SetTablesStyle("Print", dtInv, Me.dbgEquipment)
            Me.dbgEquipment.SetDataBinding(dtInv, "")
            oCG.DisableAddNew(Me.dbgEquipment, Me)
         End If
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

#End Region

End Class
