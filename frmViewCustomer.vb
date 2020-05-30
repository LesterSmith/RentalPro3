Public Class frmViewCustomer
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
   Friend WithEvents Panel1 As System.Windows.Forms.Panel
   Friend WithEvents Panel2 As System.Windows.Forms.Panel
   Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
   Friend WithEvents Panel3 As System.Windows.Forms.Panel
   Friend WithEvents dbgDet As System.Windows.Forms.DataGrid
   Friend WithEvents dbgCust As System.Windows.Forms.DataGrid
   Friend WithEvents dbgInv As System.Windows.Forms.DataGrid
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmViewCustomer))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.dbgDet = New System.Windows.Forms.DataGrid()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.dbgCust = New System.Windows.Forms.DataGrid()
        Me.Splitter1 = New System.Windows.Forms.Splitter()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.dbgInv = New System.Windows.Forms.DataGrid()
        Me.Panel1.SuspendLayout()
        CType(Me.dbgDet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        CType(Me.dbgCust, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        CType(Me.dbgInv, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.dbgDet)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 314)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(697, 152)
        Me.Panel1.TabIndex = 0
        '
        'dbgDet
        '
        Me.dbgDet.CaptionText = "Invoice Details"
        Me.dbgDet.DataMember = ""
        Me.dbgDet.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dbgDet.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dbgDet.Location = New System.Drawing.Point(0, 0)
        Me.dbgDet.Name = "dbgDet"
        Me.dbgDet.Size = New System.Drawing.Size(697, 152)
        Me.dbgDet.TabIndex = 0
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.dbgCust)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(697, 112)
        Me.Panel2.TabIndex = 1
        '
        'dbgCust
        '
        Me.dbgCust.CaptionText = "Select Customer"
        Me.dbgCust.DataMember = ""
        Me.dbgCust.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dbgCust.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dbgCust.Location = New System.Drawing.Point(0, 0)
        Me.dbgCust.Name = "dbgCust"
        Me.dbgCust.Size = New System.Drawing.Size(697, 112)
        Me.dbgCust.TabIndex = 0
        '
        'Splitter1
        '
        Me.Splitter1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Splitter1.Location = New System.Drawing.Point(0, 306)
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(697, 8)
        Me.Splitter1.TabIndex = 2
        Me.Splitter1.TabStop = False
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.dbgInv)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel3.Location = New System.Drawing.Point(0, 112)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(697, 194)
        Me.Panel3.TabIndex = 3
        '
        'dbgInv
        '
        Me.dbgInv.CaptionText = "Select Invoice Header"
        Me.dbgInv.DataMember = ""
        Me.dbgInv.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dbgInv.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dbgInv.Location = New System.Drawing.Point(0, 0)
        Me.dbgInv.Name = "dbgInv"
        Me.dbgInv.Size = New System.Drawing.Size(697, 194)
        Me.dbgInv.TabIndex = 0
        '
        'frmViewCustomer
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(697, 466)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Splitter1)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmViewCustomer"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "View Customer Data"
        Me.Panel1.ResumeLayout(False)
        CType(Me.dbgDet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        CType(Me.dbgCust, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        CType(Me.dbgInv, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
#Region " Module Variables "
   Private oDA As New CDataAccess()
   Private dtCust As New DataTable()
   Private dtInv As New DataTable()
   Private dtDet As New DataTable()
   Private SQL As String
   Private oCG As New CGrid()
   Private miInvHitRow As Integer
   Private miCustHitRow As Integer


#End Region



#Region " Private Methods "
   Private Sub LoadInvoiceDetails(ByVal InvID As Integer)
      SQL = "select * from invoice_details where invoiceid = " & InvID & " "
      SQL &= "order by record_type"
      oCG.ClearDataTableForRebinding(dtDet)
      If oDA.SendQuery(SQL, dtDet, ConnectString) > 0 Then
         oCG.BindDataTableToGrid(dtDet, dbgDet)
      End If
   End Sub

#End Region


#Region " Form & Control Events "
   Private Sub frmViewCustomer_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
      'Load customer grid
      SQL = "select CompanyName,ContactName,PhoneNumber,CustomerID "
      SQL &= "from customers order by companyname"
        Dim formats() As String = _
           {"", "150", "T", "L", _
           "", "150", "T", "L", _
           "", "100", "T", "L", _
           "", "60", "T", "L"}

      If oDA.SendQuery(SQL, dtCust, ConnectString, "dt") > 0 Then
         oCG.SetTablesStyle(dtCust, dbgCust, formats)
         oCG.BindDataTableToGrid(dtCust, dbgCust)
      End If
   End Sub

   Private Sub LoadInvoiceHeaders(ByVal CustID As Integer)
      ' load the invoices into header grid
      SQL = "select * from invoices where customerid = " & CustID & " "
      SQL &= "order by invoiceid"
      oCG.ClearDataTableForRebinding(dtInv)
      If oDA.SendQuery(SQL, dtInv, ConnectString) > 0 Then
         oCG.BindDataTableToGrid(dtInv, dbgInv)
      End If
   End Sub

   Private Sub dbgCust_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dbgCust.MouseUp
      Static Noise As Boolean
      Dim custID As Integer
      Try
         Dim b As Boolean
         If Noise Then Exit Sub
         Noise = True
         miCustHitRow = dbgCust.CurrentRowIndex
         custID = dtCust.Rows(miCustHitRow).Item("customerid")
         LoadInvoiceHeaders(custID)

         Noise = False
      Catch
      End Try
   End Sub


   Private Sub dbgInv_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dbgInv.MouseUp
      Static Noise As Boolean
      Dim invID As Integer

      Try
         Dim b As Boolean
         If Noise Then Exit Sub
         Noise = True
         miInvHitRow = dbgInv.CurrentRowIndex
         invID = dtInv.Rows(miInvHitRow).Item("invoiceid")
         LoadInvoiceDetails(invID)
         Noise = False
      Catch
      End Try
   End Sub


#End Region


End Class
