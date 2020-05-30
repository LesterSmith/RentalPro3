Imports System.Windows.Forms.Application
Public Class frmSelectCheckInInvoice
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
   Friend WithEvents InvoiceID As System.Windows.Forms.ColumnHeader
   Friend WithEvents Label1 As System.Windows.Forms.Label
   Friend WithEvents Label2 As System.Windows.Forms.Label
   Friend WithEvents lvInvoices As System.Windows.Forms.ListView
   Friend WithEvents linkPartialCheckin As System.Windows.Forms.LinkLabel
   Friend WithEvents ContactName As System.Windows.Forms.ColumnHeader
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSelectCheckInInvoice))
        Me.dbgEquipment = New System.Windows.Forms.DataGrid()
        Me.lvInvoices = New System.Windows.Forms.ListView()
        Me.Customer = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.CustID = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.InvoiceID = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ContactName = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnCheckIn = New System.Windows.Forms.Button()
        Me.linkPartialCheckin = New System.Windows.Forms.LinkLabel()
        CType(Me.dbgEquipment, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dbgEquipment
        '
        Me.dbgEquipment.AllowSorting = False
        Me.dbgEquipment.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dbgEquipment.DataMember = ""
        Me.dbgEquipment.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dbgEquipment.Location = New System.Drawing.Point(8, 274)
        Me.dbgEquipment.Name = "dbgEquipment"
        Me.dbgEquipment.Size = New System.Drawing.Size(567, 133)
        Me.dbgEquipment.TabIndex = 0
        '
        'lvInvoices
        '
        Me.lvInvoices.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvInvoices.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.Customer, Me.CustID, Me.InvoiceID, Me.ContactName})
        Me.lvInvoices.FullRowSelect = True
        Me.lvInvoices.GridLines = True
        Me.lvInvoices.Location = New System.Drawing.Point(8, 24)
        Me.lvInvoices.MultiSelect = False
        Me.lvInvoices.Name = "lvInvoices"
        Me.lvInvoices.Size = New System.Drawing.Size(567, 227)
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
        'InvoiceID
        '
        Me.InvoiceID.Text = "Invoice Number"
        Me.InvoiceID.Width = 91
        '
        'ContactName
        '
        Me.ContactName.Text = "Contact Name"
        Me.ContactName.Width = 120
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
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 257)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(110, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Equipment on Invoice"
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(501, 415)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(64, 28)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "&Cancel"
        '
        'btnCheckIn
        '
        Me.btnCheckIn.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCheckIn.Location = New System.Drawing.Point(432, 415)
        Me.btnCheckIn.Name = "btnCheckIn"
        Me.btnCheckIn.Size = New System.Drawing.Size(64, 28)
        Me.btnCheckIn.TabIndex = 5
        Me.btnCheckIn.Text = "&OK"
        '
        'linkPartialCheckin
        '
        Me.linkPartialCheckin.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.linkPartialCheckin.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.linkPartialCheckin.Location = New System.Drawing.Point(16, 417)
        Me.linkPartialCheckin.Name = "linkPartialCheckin"
        Me.linkPartialCheckin.Size = New System.Drawing.Size(216, 16)
        Me.linkPartialCheckin.TabIndex = 6
        Me.linkPartialCheckin.TabStop = True
        Me.linkPartialCheckin.Text = "How do I check in part of the invoice?"
        '
        'frmSelectCheckInInvoice
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(584, 447)
        Me.Controls.Add(Me.linkPartialCheckin)
        Me.Controls.Add(Me.btnCheckIn)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lvInvoices)
        Me.Controls.Add(Me.dbgEquipment)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MinimizeBox = False
        Me.Name = "frmSelectCheckInInvoice"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Select Invoice"
        Me.TopMost = True
        CType(Me.dbgEquipment, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region
#Region " Module Variables "
   Private oDA As New CDataAccess()
   Private m_SelectedInvoice As Integer
   Private mbWait As Boolean
   Private SQL As String
   Private miHitRow As Integer
   Private dt As New DataTable()
   Private m_ShowAll As Boolean
   Private oCG As New CGrid()


#End Region


#Region " Public Methods "
   Public Function Display(Optional ByVal ShowAllInvoices As Boolean = False) As Integer
      ' return the selected invoice

      If ShowAllInvoices Then
         m_ShowAll = True
      End If
      mbWait = True
      Me.Show()
      Do While mbWait
         DoEvents()
      Loop
      Display = m_SelectedInvoice
      Me.Close()
      DoEvents()
      Return Display
   End Function

#End Region

#Region " Form & Control Events "
   Private Sub frmSelectCheckInInvoice_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
      ' load the invoice listview with Customer Name, CustomerID, InvoiceID
      ' Load the grid with equip items on the selected invoice  
      Dim i As Integer
      Dim iInvID As Integer

      SQL &= "select distinct c.companyname, c.customerid, i.InvoiceID,i.contactname "
      SQL &= "from customers c, invoices i "
      SQL &= "where i.customerid=c.customerid "
      If Not m_ShowAll Then
         SQL &= "and i.status='CheckedOut' "
      End If
      SQL &= "order by c.companyname,i.invoiceid "

      Dim iRows As Integer = oDA.SendQuery(SQL, dt, ConnectString)
      If iRows > 0 Then
         iInvID = dt.Rows(0).Item("invoiceid")
         For i = 0 To dt.Rows.Count - 1
            With dt.Rows(i)
               lvInvoices.Items.Add(.Item("companyname"))
               lvInvoices.Items(i).SubItems.Add(.Item("customerid"))
               lvInvoices.Items(i).SubItems.Add(.Item("invoiceid"))
               lvInvoices.Items(i).SubItems.Add(.Item("contactname"))
            End With
         Next
      End If
   End Sub

   Private Sub lvInvoices_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvInvoices.SelectedIndexChanged
      Dim id As Integer
      Dim iRows As Integer


      Try
         With lvInvoices
            Try
               id = Val(lvInvoices.SelectedItems(0).SubItems(2).Text)
            Catch
               Exit Sub
            End Try
         End With
         SQL = "select Equip_ID,Equip_Name,InvoiceID,Customer_ID,Rented_Date " 'InvoiceID,InvoiceDate,Status,BalanceDue "
         SQL &= "from invoice_details " 'invoices "
         SQL &= "where invoiceid = " & id.ToString & " "

         SQL &= "and record_type = 15 "
         SQL &= "order by invoiceid "
         dt.Reset()
         dt = New DataTable("dt")
         Dim formats() As String = _
            {"", "60", "T", "L", _
             "", "150", "T", "L", _
             "", "60", "T", "R", _
             "", "60", "T", "R", _
            "MM/dd/yyyy hh:mm tt", "120", "T", "L", _
            "", "60", "F", "L"}
         If oDA.SendQuery(SQL, dt, ConnectString, "dt") > 0 Then
            oCG.SetTablesStyle("CheckIn", dt, Me.dbgEquipment, formats)
            Me.dbgEquipment.SetDataBinding(dt, "")
            oCG.DisableAddNew(Me.dbgEquipment, Me)
            oCG.CheckAllBoxes(Me.dbgEquipment.DataSource, "CheckIn")
         End If
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub


   Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
      m_SelectedInvoice = 0
      mbWait = False
   End Sub

   Private Sub btnCheckIn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCheckIn.Click
      If Me.lvInvoices.SelectedItems.Count > 0 Then
         'm_SelectedInvoice = Val(lvInvoices.SelectedItems(0).SubItems(2).Text)
         Dim oSI As New CInvoiceSplit()
         m_SelectedInvoice = oSI.CheckForSplittingInvoice(Me)
         If m_SelectedInvoice <> 0 Then
            mbWait = False
         Else
            MsgBox("Invoice checkout cancelled, please select another invoice or close the form.", MsgBoxStyle.Information)
            Exit Sub
         End If
      Else
            MsgBox("You have not selected an invoice to check in.", MsgBoxStyle.Exclamation)
         End If
   End Sub

   Private Sub frmSelectCheckInInvoice_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
      m_SelectedInvoice = 0
      mbWait = False
   End Sub

   Private Sub dbgEquipment_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs) Handles dbgEquipment.Navigate

   End Sub

   Private Sub linkPartialCheckin_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles linkPartialCheckin.LinkClicked
      Dim sTxt As String = ""
      sTxt &= "To check in only some of the equipment on an invoice, "
      sTxt &= "uncheck the items that you do not want to check in and "
      sTxt &= "click the OK button.  " & Chr(13) & Chr(10) & vbCrLf
      sTxt &= "At that point, RentalPro will create a new invoice for the "
      sTxt &= "items being checked in now and remove them from the old "
      sTxt &= "invoice.  The items, not being checked in now, will remain "
      sTxt &= "on the original invoice." & Chr(13) & Chr(10)
      Dim f As New frmHelp()
      f.CannedMessage = sTxt
      f.ShowDialog()
   End Sub

   Private Sub dbgEquipment_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dbgEquipment.MouseUp
      Static Noise As Boolean
      Try
         Dim b As Boolean
         If Noise Then Exit Sub
         Noise = True
         miHitRow = oCG.SelectCkBoxRow(Me.dbgEquipment.DataSource, Me.dbgEquipment, e, "CheckIN", b)
         Noise = False
      Catch ex As System.Exception
      End Try
   End Sub

   Private Sub frmSelectCheckInInvoice_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated

   End Sub


#End Region

End Class
