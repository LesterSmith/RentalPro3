Imports System.Windows.Forms.Application
Public Class frmMemo
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
   Friend WithEvents btnApply As System.Windows.Forms.Button
   Friend WithEvents btnCancel As System.Windows.Forms.Button
   Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
   Friend WithEvents grpMemoType As System.Windows.Forms.GroupBox
   Friend WithEvents lblAmount As System.Windows.Forms.Label
   Friend WithEvents lblCust As System.Windows.Forms.Label
   Friend WithEvents lblCustID As System.Windows.Forms.Label
   Friend WithEvents lblInvId As System.Windows.Forms.Label
   Friend WithEvents lblInvoiceID As System.Windows.Forms.Label
   Friend WithEvents lblNumber As System.Windows.Forms.Label
   Friend WithEvents lblReason As System.Windows.Forms.Label
   Friend WithEvents optBillTo As System.Windows.Forms.RadioButton
   Friend WithEvents optCash As System.Windows.Forms.RadioButton
   Friend WithEvents optCreditMemo As System.Windows.Forms.RadioButton
   Friend WithEvents optDebitMemo As System.Windows.Forms.RadioButton
   Friend WithEvents optPaidByCheck As System.Windows.Forms.RadioButton
   Friend WithEvents optPaidByCreditCard As System.Windows.Forms.RadioButton
   Friend WithEvents txtAmount As System.Windows.Forms.TextBox
   Friend WithEvents txtNumber As System.Windows.Forms.TextBox
   Friend WithEvents txtReason As System.Windows.Forms.TextBox
   Friend WithEvents chkAmtTaxable As System.Windows.Forms.CheckBox
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
      Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMemo))
      Me.grpMemoType = New System.Windows.Forms.GroupBox()
      Me.optDebitMemo = New System.Windows.Forms.RadioButton()
      Me.optCreditMemo = New System.Windows.Forms.RadioButton()
      Me.lblAmount = New System.Windows.Forms.Label()
      Me.txtAmount = New System.Windows.Forms.TextBox()
      Me.lblReason = New System.Windows.Forms.Label()
      Me.txtReason = New System.Windows.Forms.TextBox()
      Me.btnCancel = New System.Windows.Forms.Button()
      Me.btnApply = New System.Windows.Forms.Button()
      Me.GroupBox1 = New System.Windows.Forms.GroupBox()
      Me.optCash = New System.Windows.Forms.RadioButton()
      Me.optBillTo = New System.Windows.Forms.RadioButton()
      Me.optPaidByCreditCard = New System.Windows.Forms.RadioButton()
      Me.optPaidByCheck = New System.Windows.Forms.RadioButton()
      Me.lblNumber = New System.Windows.Forms.Label()
      Me.txtNumber = New System.Windows.Forms.TextBox()
      Me.lblInvId = New System.Windows.Forms.Label()
      Me.lblInvoiceID = New System.Windows.Forms.Label()
      Me.lblCust = New System.Windows.Forms.Label()
      Me.lblCustID = New System.Windows.Forms.Label()
      Me.chkAmtTaxable = New System.Windows.Forms.CheckBox()
      Me.grpMemoType.SuspendLayout()
      Me.GroupBox1.SuspendLayout()
      Me.SuspendLayout()
      '
      'grpMemoType
      '
      Me.grpMemoType.Controls.AddRange(New System.Windows.Forms.Control() {Me.optDebitMemo, Me.optCreditMemo})
      Me.grpMemoType.Location = New System.Drawing.Point(8, 8)
      Me.grpMemoType.Name = "grpMemoType"
      Me.grpMemoType.Size = New System.Drawing.Size(176, 72)
      Me.grpMemoType.TabIndex = 0
      Me.grpMemoType.TabStop = False
      Me.grpMemoType.Text = "Type of Memo"
      '
      'optDebitMemo
      '
      Me.optDebitMemo.Location = New System.Drawing.Point(8, 40)
      Me.optDebitMemo.Name = "optDebitMemo"
      Me.optDebitMemo.Size = New System.Drawing.Size(152, 15)
      Me.optDebitMemo.TabIndex = 1
      Me.optDebitMemo.Text = "Debit Customer Account"
      '
      'optCreditMemo
      '
      Me.optCreditMemo.Checked = True
      Me.optCreditMemo.Location = New System.Drawing.Point(8, 21)
      Me.optCreditMemo.Name = "optCreditMemo"
      Me.optCreditMemo.Size = New System.Drawing.Size(152, 15)
      Me.optCreditMemo.TabIndex = 0
      Me.optCreditMemo.TabStop = True
      Me.optCreditMemo.Text = "Credit Customer Account"
      '
      'lblAmount
      '
      Me.lblAmount.AutoSize = True
      Me.lblAmount.Location = New System.Drawing.Point(16, 124)
      Me.lblAmount.Name = "lblAmount"
      Me.lblAmount.Size = New System.Drawing.Size(43, 13)
      Me.lblAmount.TabIndex = 3
      Me.lblAmount.Text = "Amount"
      '
      'txtAmount
      '
      Me.txtAmount.Location = New System.Drawing.Point(64, 120)
      Me.txtAmount.Name = "txtAmount"
      Me.txtAmount.Size = New System.Drawing.Size(88, 20)
      Me.txtAmount.TabIndex = 4
      Me.txtAmount.Tag = "$#,##0.00;($#,##0.00)"
      Me.txtAmount.Text = ""
      Me.txtAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      '
      'lblReason
      '
      Me.lblReason.AutoSize = True
      Me.lblReason.Location = New System.Drawing.Point(16, 144)
      Me.lblReason.Name = "lblReason"
      Me.lblReason.Size = New System.Drawing.Size(43, 13)
      Me.lblReason.TabIndex = 5
      Me.lblReason.Text = "Reason"
      '
      'txtReason
      '
      Me.txtReason.Location = New System.Drawing.Point(64, 144)
      Me.txtReason.MaxLength = 50
      Me.txtReason.Name = "txtReason"
      Me.txtReason.Size = New System.Drawing.Size(385, 20)
      Me.txtReason.TabIndex = 6
      Me.txtReason.Tag = "(No Auto Formatting)"
      Me.txtReason.Text = ""
      '
      'btnCancel
      '
      Me.btnCancel.Location = New System.Drawing.Point(387, 178)
      Me.btnCancel.Name = "btnCancel"
      Me.btnCancel.Size = New System.Drawing.Size(68, 32)
      Me.btnCancel.TabIndex = 7
      Me.btnCancel.Text = "&Cancel"
      '
      'btnApply
      '
      Me.btnApply.Location = New System.Drawing.Point(310, 178)
      Me.btnApply.Name = "btnApply"
      Me.btnApply.Size = New System.Drawing.Size(68, 32)
      Me.btnApply.TabIndex = 8
      Me.btnApply.Text = "&Apply"
      '
      'GroupBox1
      '
      Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.optCash, Me.optBillTo, Me.optPaidByCreditCard, Me.optPaidByCheck})
      Me.GroupBox1.Location = New System.Drawing.Point(192, 8)
      Me.GroupBox1.Name = "GroupBox1"
      Me.GroupBox1.Size = New System.Drawing.Size(272, 71)
      Me.GroupBox1.TabIndex = 56
      Me.GroupBox1.TabStop = False
      Me.GroupBox1.Text = "Type Payment"
      '
      'optCash
      '
      Me.optCash.Location = New System.Drawing.Point(147, 37)
      Me.optCash.Name = "optCash"
      Me.optCash.Size = New System.Drawing.Size(96, 16)
      Me.optCash.TabIndex = 5
      Me.optCash.Text = "Cash Refund"
      '
      'optBillTo
      '
      Me.optBillTo.Location = New System.Drawing.Point(148, 15)
      Me.optBillTo.Name = "optBillTo"
      Me.optBillTo.Size = New System.Drawing.Size(106, 16)
      Me.optBillTo.TabIndex = 3
      Me.optBillTo.Text = "Bill To Customer"
      '
      'optPaidByCreditCard
      '
      Me.optPaidByCreditCard.Location = New System.Drawing.Point(9, 36)
      Me.optPaidByCreditCard.Name = "optPaidByCreditCard"
      Me.optPaidByCreditCard.Size = New System.Drawing.Size(125, 16)
      Me.optPaidByCreditCard.TabIndex = 1
      Me.optPaidByCreditCard.Text = "Apply to Credit Card"
      '
      'optPaidByCheck
      '
      Me.optPaidByCheck.Checked = True
      Me.optPaidByCheck.Location = New System.Drawing.Point(8, 16)
      Me.optPaidByCheck.Name = "optPaidByCheck"
      Me.optPaidByCheck.Size = New System.Drawing.Size(97, 16)
      Me.optPaidByCheck.TabIndex = 0
      Me.optPaidByCheck.TabStop = True
      Me.optPaidByCheck.Text = "Check Refund"
      '
      'lblNumber
      '
      Me.lblNumber.Location = New System.Drawing.Point(168, 123)
      Me.lblNumber.Name = "lblNumber"
      Me.lblNumber.Size = New System.Drawing.Size(72, 16)
      Me.lblNumber.TabIndex = 57
      Me.lblNumber.Text = "Check/Card #"
      '
      'txtNumber
      '
      Me.txtNumber.Location = New System.Drawing.Point(243, 120)
      Me.txtNumber.Name = "txtNumber"
      Me.txtNumber.Size = New System.Drawing.Size(104, 20)
      Me.txtNumber.TabIndex = 58
      Me.txtNumber.Tag = "(No Auto Formatting)"
      Me.txtNumber.Text = ""
      '
      'lblInvId
      '
      Me.lblInvId.AutoSize = True
      Me.lblInvId.Location = New System.Drawing.Point(8, 89)
      Me.lblInvId.Name = "lblInvId"
      Me.lblInvId.Size = New System.Drawing.Size(49, 13)
      Me.lblInvId.TabIndex = 59
      Me.lblInvId.Text = "Invoice #"
      '
      'lblInvoiceID
      '
      Me.lblInvoiceID.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
      Me.lblInvoiceID.Location = New System.Drawing.Point(64, 86)
      Me.lblInvoiceID.Name = "lblInvoiceID"
      Me.lblInvoiceID.Size = New System.Drawing.Size(88, 20)
      Me.lblInvoiceID.TabIndex = 60
      Me.lblInvoiceID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
      '
      'lblCust
      '
      Me.lblCust.AutoSize = True
      Me.lblCust.Location = New System.Drawing.Point(168, 90)
      Me.lblCust.Name = "lblCust"
      Me.lblCust.Size = New System.Drawing.Size(68, 13)
      Me.lblCust.TabIndex = 61
      Me.lblCust.Text = "Customer ID"
      '
      'lblCustID
      '
      Me.lblCustID.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
      Me.lblCustID.Location = New System.Drawing.Point(244, 86)
      Me.lblCustID.Name = "lblCustID"
      Me.lblCustID.Size = New System.Drawing.Size(88, 20)
      Me.lblCustID.TabIndex = 62
      Me.lblCustID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
      '
      'chkAmtTaxable
      '
      Me.chkAmtTaxable.Location = New System.Drawing.Point(64, 176)
      Me.chkAmtTaxable.Name = "chkAmtTaxable"
      Me.chkAmtTaxable.Size = New System.Drawing.Size(112, 16)
      Me.chkAmtTaxable.TabIndex = 63
      Me.chkAmtTaxable.Text = "Amount Taxable"
      '
      'frmMemo
      '
      Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
      Me.ClientSize = New System.Drawing.Size(474, 224)
      Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.chkAmtTaxable, Me.lblCustID, Me.lblCust, Me.lblInvId, Me.lblReason, Me.lblAmount, Me.lblInvoiceID, Me.txtNumber, Me.lblNumber, Me.GroupBox1, Me.btnApply, Me.btnCancel, Me.txtReason, Me.txtAmount, Me.grpMemoType})
      Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
      Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
      Me.MaximizeBox = False
      Me.MinimizeBox = False
      Me.Name = "frmMemo"
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "Credit/Debit Memo"
      Me.grpMemoType.ResumeLayout(False)
      Me.GroupBox1.ResumeLayout(False)
      Me.ResumeLayout(False)

   End Sub

#End Region

#Region " Module Variables "
   Private oDA As New CDataAccess()
   Public CurrentInvoice As Integer


#End Region

#Region " Private Methods "
   Private Sub ApplyMemo()
      Dim CustomerId As String
      Dim SQL As String
      Dim dt As DataTable

      Try
         Dim sErr As String = ""
         ' if the amount of the memo was taxable, insert or update the 
         ' sales tax record for the invoice
         If Not Me.chkAmtTaxable.Checked Then
            GoTo UpdateTables
         End If

         SQL = "select tax_id from customers "
         SQL &= "where customerid = " & Me.lblCustID.Text
         dt = New DataTable()
         If oDA.SendQuery(SQL, dt, ConnectString) > 0 Then
            If Not (IsDBNull(dt.Rows(0).Item(0)) OrElse _
               CType(dt.Rows(0).Item(0), String).Trim.Length = 0) Then
               ' tax id, so we can't update tax
               MsgBox("This customer does not pay tax.", MsgBoxStyle.Exclamation)
               Exit Sub
            End If
         Else
            Throw New System.Exception("Can't read from customer table to get tax id")
         End If

         Dim salesTax As Decimal = _
            UnFormat(Me.txtAmount.Text) * TaxRate


         dt.Reset()
         SQL = "select * from invoice_details "
         SQL &= "where invoiceid = " & Me.lblInvoiceID.Text & " "
         SQL &= "and customer_id = " & Me.lblCustID.Text
         SQL &= "and record_type = 35"
         If oDA.SendQuery(SQL, dt, ConnectString) > 0 Then
            ' have a tax record, update it
            SQL = "update invoice_details "
            SQL &= "set salestax = salestax " & IIf(Me.optCreditMemo.Checked, "- ", "+ ") & salesTax & " "
            SQL &= "where invoiceid = " & Me.lblInvoiceID.Text & " "
            SQL &= "and customer_id = " & Me.lblCustID.Text
            SQL &= " and record_type = 35"
            If oDA.SendActionSql(SQL, ConnectString, sErr) = 0 Then
               Throw New System.Exception("Can't update sales tax in invoice details")
            End If
         Else
            MsgBox("No tax record for this invoice, can't create a taxable memo.", MsgBoxStyle.Exclamation)
            Exit Sub
         End If

UpdateTables:
         ' create an invoice detail record
         Dim amt As Decimal
         SQL = "update invoices "
         If Me.optCreditMemo.Checked Then
            If salesTax <> 0 Then
               amt = UnFormat(Me.txtAmount.Text) - salesTax
            End If
            SQL &= "set balancedue = balancedue - " & amt.ToString & ", "
         Else
            If salesTax <> 0 Then
               amt = UnFormat(Me.txtAmount.Text) + salesTax
            End If
            SQL &= "set balancedue = balancedue + " & amt.ToString & ", "
         End If
         SQL &= "ckcardnumber = '" & Me.txtNumber.Text & "', "
         SQL &= "notes = notes & '" & IIf(Me.optCreditMemo.Checked, "CR ", "DB ") & " Memo: " & _
           Format(Today, "M/d/yyyy") & _
           "  " & Replace(Me.txtReason.Text.Trim, "'", "''") & _
           ", " & FormatCurrency(Me.txtAmount.Text) & vbCrLf & "' "
         SQL &= "where invoiceid = " & Me.lblInvoiceID.Text
         If oDA.SendActionSql(SQL, ConnectString, sErr) < 1 Then
            Throw New System.Exception("Could not update Invoices: " & Me.lblInvoiceID.Text)
         End If

         'if the balance was set to zero, close the invoice
         SQL = "update invoices "
         SQL &= "set status = 'CLOSED' "
         SQL &= "where invoiceid = " & Me.lblInvoiceID.Text & "  "
         SQL &= "and balancedue = 0 "
         If oDA.SendActionSql(SQL, ConnectString, sErr) < 0 Then
            Throw New System.Exception("Could not update Invoices: " & Me.lblInvoiceID.Text)
         End If

         ' insert a new invoice detail credit item
         Dim PaidOption As String
         Select Case True
            Case Me.optBillTo.Checked : PaidOption = "BT"
            Case Me.optCash.Checked : PaidOption = "CA"
            Case Me.optPaidByCheck.Checked : PaidOption = "CK"
            Case Me.optPaidByCreditCard.Checked : PaidOption = "CC"

         End Select
         SQL = "Insert into invoice_details "
         SQL &= "(invoiceid,rented_date,record_type,record_description,amtpaid, customer_id,paidoption,notes) "
         SQL &= "values("
         SQL &= Me.lblInvoiceID.Text & ", "
         SQL &= "#" & Now.ToString & "#, "
         If Me.optCreditMemo.Checked Then
            SQL &= CREDIT_MEMO_RECORD.ToString & ", 'Credit Memo', "
         Else
            SQL &= DEBIT_MEMO_RECORD.ToString & ", 'Debit Memo', "
         End If
         SQL &= UnFormat(Me.txtAmount.Text).ToString & ", "
         SQL &= Me.lblCustID.Text & ","
         SQL &= "'" & PaidOption & "', "
         SQL &= "'" & Replace(Me.txtReason.Text, "'", "''") & "'"
         SQL &= ") "
         If oDA.SendActionSql(SQL, ConnectString, sErr) < 1 Then
            Throw New System.Exception("Could not insert cash payment into invoice details: " & Me.lblInvoiceID.Text)
         End If


CloseTheForm:
         Me.Close()
         DoEvents()

      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

   Private Sub LoadCustomerCombo()
      Dim dt As New DataTable()
      Dim i As Integer
      Dim sql As String = "select customerid from invoices "
      sql &= "where invoiceid = " & Me.lblInvoiceID.Text
      If oDA.SendQuery(sql, dt, ConnectString) > 0 Then
         Me.lblCustID.Text = dt.Rows(0).Item("customerid")
      Else
         MsgBox("Can't load customer id, database error", MsgBoxStyle.Critical)
         Me.Close()
         DoEvents()
      End If

   End Sub

   Private Function CkKeyPressNumeric(ByVal riKeyAscii As Integer, ByVal roTB As TextBox) As Integer
      Dim liKeyReturn As Integer
      ' allow 0-9,., Back, Del,-,Ins, and / if in tag format
      On Error Resume Next
      CkKeyPressNumeric = riKeyAscii
      If riKeyAscii = Keys.Back Or _
         riKeyAscii = Keys.Insert Or _
         riKeyAscii = Keys.Delete Or _
         riKeyAscii = 46 Or _
         (riKeyAscii >= Keys.D0 And riKeyAscii <= Keys.D9) Or _
         riKeyAscii = 45 Or _
         riKeyAscii = 46 Or _
         (InStr(roTB.Tag, "/") > 0 And riKeyAscii = Keys.Divide) _
         Then
         If roTB.SelectionLength = 0 Then
            If InStr(roTB.Text, ".") > 0 Then
               If Len(Mid(roTB.Text, InStr(roTB.Text, ".") + 1)) > 1 Then
                  SendKeys.Send("{TAB}")
                  CkKeyPressNumeric = 0
               End If
            End If
         Else
            roTB.Text = ""
         End If
         Exit Function
      End If
      CkKeyPressNumeric = 0
   End Function

   Public Function UnFmt_T_B(ByVal roTB As TextBox) As Object
      On Error Resume Next
      UnFmt_T_B = Val(Replace(Replace(Replace(Replace(Replace(roTB.Text, "$", ""), ",", ""), ")", ""), "(", ""), "%", ""))
      If InStr(roTB.Text, "%") Then
         UnFmt_T_B = UnFmt_T_B / 100
      End If
      If InStr(roTB.Text, "(") > 0 And InStr(roTB.Text, ")") > 0 Then
         UnFmt_T_B = UnFmt_T_B * -1
      End If
   End Function

   Public Function Fmt_T_B(ByVal roTB As TextBox) As String
      On Error Resume Next
      If InStr(1, roTB.Tag, ";", 1) > 0 Then
         If InStr(roTB.Text, "-") > 0 Or (InStr(roTB.Text, "(") > 0 And InStr(roTB.Text, ")") > 0) Then
            Fmt_T_B = Format$(Math.Abs(Val(roTB.Text)), Mid$(roTB.Tag, InStr(roTB.Tag, ";") + 1))
         Else
            Fmt_T_B = Format$(Math.Abs(Val(roTB.Text)), Microsoft.VisualBasic.Left(roTB.Tag, InStr(roTB.Tag, ";") - 1))
         End If
      ElseIf InStr(1, roTB.Tag, "%", 1) > 0 Then
         Fmt_T_B = Format$(roTB.Text, roTB.Tag)
      Else
         Fmt_T_B = Format$(roTB.Text, roTB.Tag)
      End If
   End Function

   Public Function Fmt_D_F(ByVal rsTxt As Object, ByVal roTB As TextBox) As String
      On Error Resume Next

      If InStr(1, roTB.Tag, ";", 1) > 0 Then
         If InStr(rsTxt, "-") Then

            Fmt_D_F = Format$(Replace(rsTxt, "-", ""), Mid$(roTB.Tag, InStr(roTB.Tag, ";") + 1))
         Else
            Fmt_D_F = Format$(rsTxt, Microsoft.VisualBasic.Left(roTB.Tag, InStr(roTB.Tag, ";") - 1))
         End If
      ElseIf InStr(1, roTB.Tag, "%", 1) > 0 Then
         Fmt_D_F = Format$(rsTxt, roTB.Tag)
      Else
         Fmt_D_F = Format$(rsTxt, roTB.Tag)
      End If
   End Function




#End Region

#Region " Form & Control Events "
   Private Sub frmMemo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
      Me.lblInvoiceID.Text = CurrentInvoice
      LoadCustomerCombo()
   End Sub

   Private Sub txtAmount_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAmount.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
      e.Handled = CkKeyPressNumeric(Asc(Val(e.KeyChar)), txtAmount) = 0
   End Sub
   Private Sub txtAmount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAmount.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub txtAmount_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAmount.Enter
      txtAmount.Text = UnFmt_T_B(txtAmount)
      txtAmount.SelectionStart = 0
      txtAmount.SelectionLength = txtAmount.Text.Trim.Length
   End Sub
   Private Sub txtAmount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAmount.Leave
      txtAmount.Text = Fmt_T_B(txtAmount)
   End Sub


   Private Sub txtNumber_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtNumber.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
   End Sub
   Private Sub txtNumber_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtNumber.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub txtNumber_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNumber.Enter
      txtNumber.SelectionStart = 0
      txtNumber.SelectionLength = txtNumber.Text.Trim.Length
   End Sub
   Private Sub txtReason_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtReason.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
   End Sub
   Private Sub txtReason_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtReason.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub txtReason_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtReason.Enter
      txtReason.SelectionStart = 0
      txtReason.SelectionLength = txtReason.Text.Trim.Length
   End Sub

   Private Sub btnApply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnApply.Click
      ApplyMemo()
   End Sub


   Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
      Me.Close()
      DoEvents()
   End Sub


#End Region

End Class
