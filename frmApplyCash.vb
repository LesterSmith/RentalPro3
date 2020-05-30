Imports System.Windows.Forms.Application
Public Class frmApplyCash
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
   Friend WithEvents btnClose As System.Windows.Forms.Button
   Friend WithEvents btnUpdate As System.Windows.Forms.Button
   Friend WithEvents dbgCustomers As System.Windows.Forms.DataGrid
   Friend WithEvents dbgInvoices As System.Windows.Forms.DataGrid
   Friend WithEvents lblCkAmt As System.Windows.Forms.Label
   Friend WithEvents lblCkNbr As System.Windows.Forms.Label
   Friend WithEvents lblRemBal As System.Windows.Forms.Label
   Friend WithEvents txtBalanceLeft As System.Windows.Forms.TextBox
   Friend WithEvents txtCheckAmount As System.Windows.Forms.TextBox
   Friend WithEvents txtCheckNumber As System.Windows.Forms.TextBox
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
      Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmApplyCash))
      Me.dbgCustomers = New System.Windows.Forms.DataGrid()
      Me.dbgInvoices = New System.Windows.Forms.DataGrid()
      Me.lblCkAmt = New System.Windows.Forms.Label()
      Me.txtCheckAmount = New System.Windows.Forms.TextBox()
      Me.txtCheckNumber = New System.Windows.Forms.TextBox()
      Me.lblCkNbr = New System.Windows.Forms.Label()
      Me.txtBalanceLeft = New System.Windows.Forms.TextBox()
      Me.lblRemBal = New System.Windows.Forms.Label()
      Me.btnUpdate = New System.Windows.Forms.Button()
      Me.btnClose = New System.Windows.Forms.Button()
      CType(Me.dbgCustomers, System.ComponentModel.ISupportInitialize).BeginInit()
      CType(Me.dbgInvoices, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.SuspendLayout()
      '
      'dbgCustomers
      '
      Me.dbgCustomers.AllowSorting = False
      Me.dbgCustomers.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                  Or System.Windows.Forms.AnchorStyles.Right)
      Me.dbgCustomers.CaptionText = "Select Customer"
      Me.dbgCustomers.DataMember = ""
      Me.dbgCustomers.HeaderForeColor = System.Drawing.SystemColors.ControlText
      Me.dbgCustomers.Location = New System.Drawing.Point(8, 5)
      Me.dbgCustomers.Name = "dbgCustomers"
      Me.dbgCustomers.Size = New System.Drawing.Size(424, 104)
      Me.dbgCustomers.TabIndex = 0
      '
      'dbgInvoices
      '
      Me.dbgInvoices.AllowSorting = False
      Me.dbgInvoices.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                  Or System.Windows.Forms.AnchorStyles.Left) _
                  Or System.Windows.Forms.AnchorStyles.Right)
      Me.dbgInvoices.CaptionText = "Apply Cash to Customer Invoices"
      Me.dbgInvoices.DataMember = ""
      Me.dbgInvoices.HeaderForeColor = System.Drawing.SystemColors.ControlText
      Me.dbgInvoices.Location = New System.Drawing.Point(8, 122)
      Me.dbgInvoices.Name = "dbgInvoices"
      Me.dbgInvoices.Size = New System.Drawing.Size(424, 142)
      Me.dbgInvoices.TabIndex = 1
      '
      'lblCkAmt
      '
      Me.lblCkAmt.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
      Me.lblCkAmt.AutoSize = True
      Me.lblCkAmt.Location = New System.Drawing.Point(440, 8)
      Me.lblCkAmt.Name = "lblCkAmt"
      Me.lblCkAmt.Size = New System.Drawing.Size(78, 13)
      Me.lblCkAmt.TabIndex = 2
      Me.lblCkAmt.Text = "Check Amount"
      '
      'txtCheckAmount
      '
      Me.txtCheckAmount.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
      Me.txtCheckAmount.Location = New System.Drawing.Point(443, 28)
      Me.txtCheckAmount.Name = "txtCheckAmount"
      Me.txtCheckAmount.Size = New System.Drawing.Size(109, 20)
      Me.txtCheckAmount.TabIndex = 3
      Me.txtCheckAmount.Tag = "(No Auto Formatting)"
      Me.txtCheckAmount.Text = ""
      Me.txtCheckAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      '
      'txtCheckNumber
      '
      Me.txtCheckNumber.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
      Me.txtCheckNumber.Location = New System.Drawing.Point(443, 72)
      Me.txtCheckNumber.Name = "txtCheckNumber"
      Me.txtCheckNumber.Size = New System.Drawing.Size(109, 20)
      Me.txtCheckNumber.TabIndex = 5
      Me.txtCheckNumber.Tag = "(No Auto Formatting)"
      Me.txtCheckNumber.Text = ""
      Me.txtCheckNumber.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      '
      'lblCkNbr
      '
      Me.lblCkNbr.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
      Me.lblCkNbr.AutoSize = True
      Me.lblCkNbr.Location = New System.Drawing.Point(441, 56)
      Me.lblCkNbr.Name = "lblCkNbr"
      Me.lblCkNbr.Size = New System.Drawing.Size(80, 13)
      Me.lblCkNbr.TabIndex = 4
      Me.lblCkNbr.Text = "Check Number"
      '
      'txtBalanceLeft
      '
      Me.txtBalanceLeft.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
      Me.txtBalanceLeft.Location = New System.Drawing.Point(443, 120)
      Me.txtBalanceLeft.Name = "txtBalanceLeft"
      Me.txtBalanceLeft.ReadOnly = True
      Me.txtBalanceLeft.Size = New System.Drawing.Size(109, 20)
      Me.txtBalanceLeft.TabIndex = 7
      Me.txtBalanceLeft.Tag = "(No Auto Formatting)"
      Me.txtBalanceLeft.Text = ""
      Me.txtBalanceLeft.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      '
      'lblRemBal
      '
      Me.lblRemBal.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
      Me.lblRemBal.AutoSize = True
      Me.lblRemBal.Location = New System.Drawing.Point(443, 104)
      Me.lblRemBal.Name = "lblRemBal"
      Me.lblRemBal.Size = New System.Drawing.Size(93, 13)
      Me.lblRemBal.TabIndex = 6
      Me.lblRemBal.Text = "Balance To Apply"
      '
      'btnUpdate
      '
      Me.btnUpdate.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
      Me.btnUpdate.Location = New System.Drawing.Point(448, 184)
      Me.btnUpdate.Name = "btnUpdate"
      Me.btnUpdate.Size = New System.Drawing.Size(96, 30)
      Me.btnUpdate.TabIndex = 8
      Me.btnUpdate.Text = "&Update"
      '
      'btnClose
      '
      Me.btnClose.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
      Me.btnClose.Location = New System.Drawing.Point(448, 224)
      Me.btnClose.Name = "btnClose"
      Me.btnClose.Size = New System.Drawing.Size(96, 32)
      Me.btnClose.TabIndex = 9
      Me.btnClose.Text = "&Close"
      '
      'frmApplyCash
      '
      Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
      Me.ClientSize = New System.Drawing.Size(568, 277)
      Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnClose, Me.btnUpdate, Me.txtBalanceLeft, Me.lblRemBal, Me.txtCheckNumber, Me.lblCkNbr, Me.txtCheckAmount, Me.lblCkAmt, Me.dbgInvoices, Me.dbgCustomers})
      Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
      Me.MinimizeBox = False
      Me.Name = "frmApplyCash"
      Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "Receive Payments on Customer Account"
      CType(Me.dbgCustomers, System.ComponentModel.ISupportInitialize).EndInit()
      CType(Me.dbgInvoices, System.ComponentModel.ISupportInitialize).EndInit()
      Me.ResumeLayout(False)

   End Sub

#End Region
   Private SQL As String
   Private oDA As New CDataAccess()
   Private dtCustomers As DataTable
   Private dtInvoices As DataTable
   Private oCG As New CGrid()
   Private iHitRow As Integer
   Private miCustId As Integer

   Structure CashApplied
      Dim InvoiceID As Integer
      Dim Amount As Decimal
      Dim CustomerID As Integer
   End Structure
   Dim CAitem As CashApplied
   Dim CA As New ArrayList()

   Private Sub frmApplyCash_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
      LoadCustomers()
   End Sub
   Private Sub LoadCustomers()
      ' load the customer grid with customer data
      ' for which there are invoices with non zero balances

      SQL = "select distinct c.CompanyName,c.CustomerID "
      SQL &= "from Customers c, Invoices i "
      SQL &= "where i.customerid=c.customerid "
      SQL &= "and i.balancedue > 0 "
      SQL &= "and i.status= 'OPEN' "
      SQL &= "order by c.companyname"
      dtCustomers = New DataTable("dt")

      Dim Formats() As String = {"", "150", "T", "L", _
                                 "", "60", "T", "L"}
      If oDA.SendQuery(SQL, dtCustomers, ConnectString, "dt") > 0 Then
         oCG.SetTablesStyle(dtCustomers, _
            Me.dbgCustomers, _
            Formats)
         Me.dbgCustomers.SetDataBinding(dtCustomers, "")
         oCG.DisableAddNew(dbgCustomers, Me)
      Else
         MsgBox("There are no customers with unpaid balances.", MsgBoxStyle.Information)
      End If
   End Sub

   Private Sub dbgCustomers_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dbgCustomers.MouseUp
      Try
         iHitRow = oCG.GetClickedRow(e, Me.dbgCustomers)
         dbgCustomers.Select(iHitRow)
         miCustId = dtCustomers.Rows(iHitRow).Item("customerid")
         LoadInvoices(iHitRow)
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub
   Private Sub LoadInvoices(ByVal row As Integer)
      ' Load the invoices for the customer specified in the 
      ' selected row of the customer grid
      Dim iCustomerID As Integer = dtCustomers.Rows(row).Item("customerid")
      SQL = "select CustomerID, InvoiceID,InvoiceDate,PONumber,BalanceDue "
      SQL &= "from invoices "
      SQL &= "where customerid= " & iCustomerID & " "
      SQL &= "and balancedue <>0 "
      SQL &= "and status='OPEN' "
      SQL &= "order by invoiceid"
      dtInvoices = New DataTable("dt")
      Me.dbgInvoices.SetDataBinding(dtInvoices, "")
      If oDA.SendQuery(SQL, dtInvoices, ConnectString, "dt") > 0 Then
         Dim Formats() As String = _
            {"", "60", "T", "L", _
            "", "60", "T", "L", _
            "MM/dd/yyyy", "60", "T", "L", _
            "", "60", "T", "L", _
            "$#,##0.00", "100", "T", "R"}
         oCG.SetTablesStyle("Apply", dtInvoices, dbgInvoices, Formats)
         oCG.BindDataTableToGrid(dtInvoices, dbgInvoices)
         oCG.DisableAddNew(dbgInvoices, Me)
      End If
   End Sub
   Private Sub txtBalanceLeft_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtBalanceLeft.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
   End Sub
   Private Sub txtBalanceLeft_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtBalanceLeft.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub txtBalanceLeft_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBalanceLeft.Enter
      txtBalanceLeft.SelectionStart = 0
      txtBalanceLeft.SelectionLength = txtBalanceLeft.Text.Trim.Length
   End Sub
   Private Sub txtCheckAmount_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCheckAmount.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
   End Sub
   Private Sub txtCheckAmount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCheckAmount.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub txtCheckAmount_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCheckAmount.Enter
      txtCheckAmount.SelectionStart = 0
      txtCheckAmount.SelectionLength = txtCheckAmount.Text.Trim.Length
   End Sub
   Private Sub txtCheckAmount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCheckAmount.Leave
      txtCheckAmount.Text = FormatCurrency(txtCheckAmount.Text)
      If txtBalanceLeft.Text.Trim.Length = 0 Then
         txtBalanceLeft.Text = txtCheckAmount.Text
      End If
   End Sub
   Private Sub txtCheckNumber_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCheckNumber.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
   End Sub
   Private Sub txtCheckNumber_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCheckNumber.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub txtCheckNumber_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCheckNumber.Enter
      txtCheckNumber.SelectionStart = 0
      txtCheckNumber.SelectionLength = txtCheckNumber.Text.Trim.Length
   End Sub

   Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
      If CA.Count > 0 Then
         Dim sMsg As String
         Dim iRV As Integer
         sMsg = "You have applied cash to one or more invoices," & Chr(10)
         sMsg &= "but they have not yet been updated.  Do you want" & Chr(10)
         sMsg &= "to close without updating the invoices with the cash" & Chr(10)
         sMsg &= "applied?" & Chr(10)
         sMsg &= "" & Chr(10)
         sMsg &= "Click Yes to close without applying the cash.  Click" & Chr(10)
         sMsg &= "No to leave the form displayed and then click " & Chr(10)
         sMsg &= "the Update button to update the invoices in the" & Chr(10)
         sMsg &= "database." & Chr(10)
         sMsg &= "" & Chr(10)
         iRV = MsgBox(sMsg, CType(292, Microsoft.VisualBasic.MsgBoxStyle), "Confirm Cancel Of Cash Application")

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

   Private Sub dbgInvoices_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dbgInvoices.MouseUp
      Dim bChecked As Boolean
      Dim i As Integer = oCG.SelectCkBoxRow(dtInvoices, Me.dbgInvoices, e, "Apply", bChecked)
      ApplyCash(i, bChecked)
   End Sub
   Private Sub ApplyCash(ByVal iInvRow As Integer, ByVal Checked As Boolean)
      ' 1. ensure that the cash has been entered
      ' 2. subtract the amount of the selected invoice from the
      '    total cash and place the balance into the cash left box
      ' 3. build a collection of cash applied objects containing
      '    the invoiceid, cash amount so we can update the invoice
      '    and insert a cash applied item to the invoice details table
      ' 4. ensure that the cash applied does not exceed the total cash
      ' 5. possibly create an unapplied cash item somewhere if the check
      '    is for more than the total of the invoices.
      Dim dAmt As Decimal
      Dim balleft As Decimal
      Dim i As Integer
      Dim iInv As Integer

      If Me.txtCheckAmount.Text.Trim.Length = 0 Then
         MsgBox("Check amount must be entered before you can apply cash to an invoice.", MsgBoxStyle.Exclamation)
         Exit Sub
      End If

      Select Case Checked
         Case True
            If UnFormat(Me.txtBalanceLeft.Text) > 0 Then
               dAmt = UnFormat(dtInvoices.Rows(iInvRow).Item("balancedue"))
               balleft = UnFormat(Me.txtBalanceLeft.Text)
               If balleft >= dAmt Then
                  balleft -= dAmt
                  With CAitem
                     .InvoiceID = dtInvoices.Rows(iInvRow).Item("invoiceid")
                     .Amount = dAmt
                     .CustomerID = dtCustomers.Rows(Me.dbgCustomers.CurrentRowIndex).Item("customerid")
                     CA.Add(CAitem)
                     Me.txtBalanceLeft.Text = FormatCurrency(balleft)
                  End With
               Else
                  With CAitem
                     .InvoiceID = dtInvoices.Rows(iInvRow).Item("invoiceid")
                     .Amount = balleft
                     CA.Add(CAitem)
                     Me.txtBalanceLeft.Text = FormatCurrency(0)
                  End With
               End If
            Else
               MsgBox("There is no remaining cash amount to apply.", MsgBoxStyle.Exclamation)
            End If
         Case Else
            ' user unchecked the item, remove it from the 
            ' array and update the bal left
            For i = 0 To CA.Count - 1
               CAitem = CA(i)
               With CAitem
                  If .InvoiceID = dtInvoices.Rows(iInvRow).Item("invoiceid") Then
                     balleft = UnFormat(Me.txtBalanceLeft.Text)
                     dAmt = .Amount
                     CA.RemoveAt(i)
                     Me.txtBalanceLeft.Text = FormatCurrency(balleft + dAmt)
                  End If
               End With
            Next
      End Select
   End Sub

   Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
      ' 1. if the arraylist has any items in it, update the invoices tables to 
      '    correct the balance due field.
      ' 2. write an invoice detail credit memo item so that the invoice details will have
      '    a record of the check that paid it off.
      Dim i As Integer
      Dim sErr As String


      Try
         If CA.Count = 0 Then
            MsgBox("You have not applied any cash to invoices.", MsgBoxStyle.Exclamation)
            Exit Sub
         End If
         If Me.txtCheckNumber.Text.Trim.Length = 0 Then
            MsgBox("You must enter a check number for the cash being applied.", MsgBoxStyle.Exclamation)
            Exit Sub
         End If

         ' update the invoices in the tables.
         For i = 0 To CA.Count - 1
            CAitem = CA(i)
            With CAitem
               SQL = "update invoices "
               SQL &= "set balancedue = balancedue - " & .Amount.ToString & ", "
               SQL &= "ckcardnumber = '" & Me.txtCheckNumber.Text & "' "
               SQL &= "where invoiceid = " & .InvoiceID.ToString
               If oDA.SendActionSql(SQL, ConnectString, sErr) < 1 Then
                  Throw New System.Exception("Could not update Invoices: " & .InvoiceID.ToString)
               End If

               ' if the balance was set to zero, close the invoice
               SQL = "update invoices "
               SQL &= "set status = 'CLOSED' "
               SQL &= "where invoiceid = " & .InvoiceID.ToString & "  "
               SQL &= "and balancedue = 0 "
               If oDA.SendActionSql(SQL, ConnectString, sErr) < 0 Then
                  Throw New System.Exception("Could not update Invoices: " & .InvoiceID.ToString)
               End If

               ' insert a new invoice detail credit item
               SQL = "Insert into invoice_details "
               SQL &= "(invoiceid,rented_date,record_type,record_description,amtpaid, customer_id) "
               SQL &= "values("
               SQL &= .InvoiceID.ToString & ", "
               SQL &= "#" & Now.ToString & "#, "
               SQL &= CASH_ON_ACCOUNT_RECORD.ToString & ", 'Cash on Account', "
               SQL &= .Amount.ToString & ", "
               SQL &= miCustId.ToString & ") "
               If oDA.SendActionSql(SQL, ConnectString, sErr) < 1 Then
                  Throw New System.Exception("Could not insert cash payment into invoice details: " & .InvoiceID.ToString)
               End If
            End With
         Next
         If UnFormat(Me.txtBalanceLeft.Text) = 0 Then
            Me.txtCheckAmount.Text = ""
         End If
         For i = CA.Count - 1 To 0 Step -1
            CA.RemoveAt(i)
         Next

         MsgBox("Cash has been applied to customer invoices.", MsgBoxStyle.Information)
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

End Class
