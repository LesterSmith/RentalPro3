Imports System.Windows.Forms.Application
Public Class frmViewCustomerAccount
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
   Friend WithEvents btnCheckIn As System.Windows.Forms.Button
   Friend WithEvents CustID As System.Windows.Forms.ColumnHeader
   Friend WithEvents Customer As System.Windows.Forms.ColumnHeader
   Friend WithEvents dbgEquipment As System.Windows.Forms.DataGrid
   Friend WithEvents Label1 As System.Windows.Forms.Label
   Friend WithEvents Label2 As System.Windows.Forms.Label
   Friend WithEvents lvInvoices As System.Windows.Forms.ListView
   Friend WithEvents txtCustomerTotal As System.Windows.Forms.TextBox
   Friend WithEvents lblBalance As System.Windows.Forms.Label
   Friend WithEvents txtNbrInvoices As System.Windows.Forms.TextBox
   Friend WithEvents lblNbrInvoices As System.Windows.Forms.Label
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
      Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmViewCustomerAccount))
      Me.dbgEquipment = New System.Windows.Forms.DataGrid()
      Me.lvInvoices = New System.Windows.Forms.ListView()
      Me.Customer = New System.Windows.Forms.ColumnHeader()
      Me.CustID = New System.Windows.Forms.ColumnHeader()
      Me.Label1 = New System.Windows.Forms.Label()
      Me.Label2 = New System.Windows.Forms.Label()
      Me.btnCancel = New System.Windows.Forms.Button()
      Me.btnCheckIn = New System.Windows.Forms.Button()
      Me.txtCustomerTotal = New System.Windows.Forms.TextBox()
      Me.lblBalance = New System.Windows.Forms.Label()
      Me.txtNbrInvoices = New System.Windows.Forms.TextBox()
      Me.lblNbrInvoices = New System.Windows.Forms.Label()
      CType(Me.dbgEquipment, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.SuspendLayout()
      '
      'dbgEquipment
      '
      Me.dbgEquipment.AllowSorting = False
      Me.dbgEquipment.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                  Or System.Windows.Forms.AnchorStyles.Left) _
                  Or System.Windows.Forms.AnchorStyles.Right)
      Me.dbgEquipment.DataMember = ""
      Me.dbgEquipment.HeaderForeColor = System.Drawing.SystemColors.ControlText
      Me.dbgEquipment.Location = New System.Drawing.Point(8, 136)
      Me.dbgEquipment.Name = "dbgEquipment"
      Me.dbgEquipment.Size = New System.Drawing.Size(487, 152)
      Me.dbgEquipment.TabIndex = 0
      '
      'lvInvoices
      '
      Me.lvInvoices.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                  Or System.Windows.Forms.AnchorStyles.Right)
      Me.lvInvoices.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.Customer, Me.CustID})
      Me.lvInvoices.FullRowSelect = True
      Me.lvInvoices.GridLines = True
      Me.lvInvoices.Location = New System.Drawing.Point(8, 24)
      Me.lvInvoices.MultiSelect = False
      Me.lvInvoices.Name = "lvInvoices"
      Me.lvInvoices.Size = New System.Drawing.Size(487, 88)
      Me.lvInvoices.TabIndex = 1
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
      Me.Label1.Location = New System.Drawing.Point(8, 8)
      Me.Label1.Name = "Label1"
      Me.Label1.Size = New System.Drawing.Size(75, 13)
      Me.Label1.TabIndex = 2
      Me.Label1.Text = "Select Invoice"
      '
      'Label2
      '
      Me.Label2.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                  Or System.Windows.Forms.AnchorStyles.Left) _
                  Or System.Windows.Forms.AnchorStyles.Right)
      Me.Label2.AutoSize = True
      Me.Label2.Location = New System.Drawing.Point(8, 122)
      Me.Label2.Name = "Label2"
      Me.Label2.Size = New System.Drawing.Size(113, 13)
      Me.Label2.TabIndex = 3
      Me.Label2.Text = "Equipment on Invoice"
      '
      'btnCancel
      '
      Me.btnCancel.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
      Me.btnCancel.Location = New System.Drawing.Point(432, 330)
      Me.btnCancel.Name = "btnCancel"
      Me.btnCancel.Size = New System.Drawing.Size(64, 28)
      Me.btnCancel.TabIndex = 4
      Me.btnCancel.Text = "&Cancel"
      '
      'btnCheckIn
      '
      Me.btnCheckIn.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
      Me.btnCheckIn.Location = New System.Drawing.Point(325, 330)
      Me.btnCheckIn.Name = "btnCheckIn"
      Me.btnCheckIn.Size = New System.Drawing.Size(101, 28)
      Me.btnCheckIn.TabIndex = 5
      Me.btnCheckIn.Text = "&Print Statement"
      '
      'txtCustomerTotal
      '
      Me.txtCustomerTotal.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
      Me.txtCustomerTotal.Location = New System.Drawing.Point(376, 294)
      Me.txtCustomerTotal.Name = "txtCustomerTotal"
      Me.txtCustomerTotal.Size = New System.Drawing.Size(104, 20)
      Me.txtCustomerTotal.TabIndex = 6
      Me.txtCustomerTotal.Text = ""
      Me.txtCustomerTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      '
      'lblBalance
      '
      Me.lblBalance.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
      Me.lblBalance.Location = New System.Drawing.Point(316, 290)
      Me.lblBalance.Name = "lblBalance"
      Me.lblBalance.Size = New System.Drawing.Size(54, 26)
      Me.lblBalance.TabIndex = 7
      Me.lblBalance.Text = "Customer Balance"
      '
      'txtNbrInvoices
      '
      Me.txtNbrInvoices.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
      Me.txtNbrInvoices.Location = New System.Drawing.Point(269, 294)
      Me.txtNbrInvoices.Name = "txtNbrInvoices"
      Me.txtNbrInvoices.Size = New System.Drawing.Size(37, 20)
      Me.txtNbrInvoices.TabIndex = 8
      Me.txtNbrInvoices.Text = ""
      Me.txtNbrInvoices.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      '
      'lblNbrInvoices
      '
      Me.lblNbrInvoices.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
      Me.lblNbrInvoices.Location = New System.Drawing.Point(214, 292)
      Me.lblNbrInvoices.Name = "lblNbrInvoices"
      Me.lblNbrInvoices.Size = New System.Drawing.Size(51, 31)
      Me.lblNbrInvoices.TabIndex = 9
      Me.lblNbrInvoices.Text = "Invoice Count"
      '
      'frmViewCustomerAccount
      '
      Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
      Me.ClientSize = New System.Drawing.Size(504, 365)
      Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.lblNbrInvoices, Me.txtNbrInvoices, Me.lblBalance, Me.txtCustomerTotal, Me.btnCheckIn, Me.btnCancel, Me.Label2, Me.Label1, Me.lvInvoices, Me.dbgEquipment})
      Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
      Me.MaximizeBox = False
      Me.MinimizeBox = False
      Me.Name = "frmViewCustomerAccount"
      Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "View Customer Account"
      Me.TopMost = True
      CType(Me.dbgEquipment, System.ComponentModel.ISupportInitialize).EndInit()
      Me.ResumeLayout(False)

   End Sub

#End Region
#Region " Module Variables "
   Private oDA As New CDataAccess()
   Private m_SelectedInvoice As Integer
   Private mbWait As Boolean
   Private SQL As String
   Private dt As New DataTable()
   Private m_ShowAll As Boolean


#End Region


#Region " Form & Control Events "
   Private Sub frmViewCustomerAccount_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
      ' load the invoice listview with Customer Name, CustomerID, InvoiceID
      ' Load the grid with equip items on the selected invoice  
      Dim i As Integer
      Dim iInvID As Integer

      SQL &= "select distinct c.companyname, c.customerid  "
      SQL &= "from customers c, invoices i "
      SQL &= "where i.customerid=c.customerid "
      If Not m_ShowAll Then
         SQL &= "and i.status='OPEN' "
      End If
      SQL &= "order by c.companyname "

      Dim iRows As Integer = oDA.SendQuery(SQL, dt, ConnectString)
      If iRows > 0 Then
         'iInvID = dt.Rows(0).Item("invoiceid")
         For i = 0 To dt.Rows.Count - 1
            With dt.Rows(i)
               lvInvoices.Items.Add(.Item("companyname"))
               lvInvoices.Items(i).SubItems.Add(.Item("customerid"))
               'lvInvoices.Items(i).SubItems.Add(.Item("invoiceid"))
            End With
         Next
      End If
   End Sub

   Private Sub lvInvoices_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvInvoices.SelectedIndexChanged
      Dim id As Integer
      Dim iRows As Integer
      Dim oCG As New CGrid()
      Dim i As Integer
      Dim tot As Decimal = 0

      Try
         With lvInvoices
            Try
               id = Val(lvInvoices.SelectedItems(0).SubItems(1).Text)
            Catch
               Exit Sub
            End Try
         End With
         SQL = "select InvoiceId,Status,InvoiceDate,PONumber,CkCardNumber,PaidOption,BalanceDue "
         SQL &= "from invoices "
         SQL &= "where customerid = " & id.ToString & " "
         If Not m_ShowAll Then
            SQL &= "and status = 'OPEN' "
         End If

         SQL &= "order by Invoicedate "
         dt.Reset()
         dt = New DataTable("dt")

         If oDA.SendQuery(SQL, dt, ConnectString, "dt") > 0 Then
            Dim formats() As String = _
               {"", "60", "T", "R", _
                "", "60", "T", "L", _
                "MM/dd/yyyy", "100", "T", "L", _
                "", "60", "T", "L", _
                "", "60", "T", "L", _
                "", "60", "T", "L", _
                "$#,##0.00", "100", "T", "R"}

            oCG.SetTablesStyle(dt, Me.dbgEquipment, formats)
            Me.dbgEquipment.SetDataBinding(dt, "")
            For i = 0 To dt.Rows.Count - 1
               tot += dt.Rows(i).Item("balancedue")
            Next
            Me.txtNbrInvoices.Text = dt.Rows.Count
            Me.txtCustomerTotal.Text = FormatCurrency(tot)
         End If
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub


   Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
      Me.Close()
      DoEvents()
   End Sub

   ''' <summary>
   '''
   ''' </summary>
   ''' <param name = "sender"></param>
   ''' <param name = "e"></param>
   Private Sub btnCheckIn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCheckIn.Click
      Dim custID As Integer
      If Me.lvInvoices.SelectedItems.Count > 0 Then
         custID = Val(lvInvoices.SelectedItems(0).SubItems(1).Text)
         Dim oPS As New CPrintStatements()
         oPS.CustomerID = custID
         oPS.Preview = True
         oPS.PrintStatements()
      Else
         MsgBox("You have not selected an invoice to check in.", MsgBoxStyle.Exclamation)
      End If
   End Sub


#End Region


#Region " Public Properties "
   Public Property ShowAll() As Boolean
      Get
         Return m_ShowAll
      End Get
      Set(ByVal Value As Boolean)
         m_ShowAll = Value
      End Set
   End Property

#End Region

End Class
