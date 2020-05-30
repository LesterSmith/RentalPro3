Imports System.Windows.Forms.Application
Public Class frmCheckinMods
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
   Friend WithEvents Label1 As System.Windows.Forms.Label
   Friend WithEvents Label2 As System.Windows.Forms.Label
   Friend WithEvents lblEquipID As System.Windows.Forms.Label
   Friend WithEvents lblEquipName As System.Windows.Forms.Label
   Friend WithEvents Label3 As System.Windows.Forms.Label
   Friend WithEvents cbPeriod As System.Windows.Forms.ComboBox
   Friend WithEvents buttonCancel As System.Windows.Forms.Button
   Friend WithEvents buttonSave As System.Windows.Forms.Button
   Friend WithEvents cbNbrPeriods As System.Windows.Forms.ComboBox
   Friend WithEvents lblNP As System.Windows.Forms.Label
   Friend WithEvents optUpdateCurrRow As System.Windows.Forms.RadioButton
   Friend WithEvents optAddNewRow As System.Windows.Forms.RadioButton
   Friend WithEvents Label4 As System.Windows.Forms.Label
   Friend WithEvents textPrice As System.Windows.Forms.TextBox
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCheckinMods))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lblEquipID = New System.Windows.Forms.Label()
        Me.lblEquipName = New System.Windows.Forms.Label()
        Me.cbNbrPeriods = New System.Windows.Forms.ComboBox()
        Me.lblNP = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cbPeriod = New System.Windows.Forms.ComboBox()
        Me.buttonCancel = New System.Windows.Forms.Button()
        Me.buttonSave = New System.Windows.Forms.Button()
        Me.optUpdateCurrRow = New System.Windows.Forms.RadioButton()
        Me.optAddNewRow = New System.Windows.Forms.RadioButton()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.textPrice = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 1)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(48, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Equip ID"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 38)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(65, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Equip Name"
        '
        'lblEquipID
        '
        Me.lblEquipID.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblEquipID.Location = New System.Drawing.Point(10, 18)
        Me.lblEquipID.Name = "lblEquipID"
        Me.lblEquipID.Size = New System.Drawing.Size(88, 16)
        Me.lblEquipID.TabIndex = 2
        '
        'lblEquipName
        '
        Me.lblEquipName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblEquipName.Location = New System.Drawing.Point(8, 55)
        Me.lblEquipName.Name = "lblEquipName"
        Me.lblEquipName.Size = New System.Drawing.Size(223, 16)
        Me.lblEquipName.TabIndex = 3
        '
        'cbNbrPeriods
        '
        Me.cbNbrPeriods.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6", "7"})
        Me.cbNbrPeriods.Location = New System.Drawing.Point(8, 97)
        Me.cbNbrPeriods.Name = "cbNbrPeriods"
        Me.cbNbrPeriods.Size = New System.Drawing.Size(86, 21)
        Me.cbNbrPeriods.TabIndex = 0
        '
        'lblNP
        '
        Me.lblNP.AutoSize = True
        Me.lblNP.Location = New System.Drawing.Point(8, 79)
        Me.lblNP.Name = "lblNP"
        Me.lblNP.Size = New System.Drawing.Size(62, 13)
        Me.lblNP.TabIndex = 5
        Me.lblNP.Text = "Nbr Periods"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(8, 121)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(71, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Rental Period"
        '
        'cbPeriod
        '
        Me.cbPeriod.Items.AddRange(New Object() {"Half Day", "Daily", "Weekly", "Monthly", "Week End", "Hourly"})
        Me.cbPeriod.Location = New System.Drawing.Point(8, 138)
        Me.cbPeriod.Name = "cbPeriod"
        Me.cbPeriod.Size = New System.Drawing.Size(88, 21)
        Me.cbPeriod.TabIndex = 1
        '
        'buttonCancel
        '
        Me.buttonCancel.Location = New System.Drawing.Point(154, 178)
        Me.buttonCancel.Name = "buttonCancel"
        Me.buttonCancel.Size = New System.Drawing.Size(72, 24)
        Me.buttonCancel.TabIndex = 4
        Me.buttonCancel.Text = "&Cancel"
        '
        'buttonSave
        '
        Me.buttonSave.Location = New System.Drawing.Point(154, 146)
        Me.buttonSave.Name = "buttonSave"
        Me.buttonSave.Size = New System.Drawing.Size(72, 24)
        Me.buttonSave.TabIndex = 3
        Me.buttonSave.Text = "&Save"
        '
        'optUpdateCurrRow
        '
        Me.optUpdateCurrRow.Checked = True
        Me.optUpdateCurrRow.Location = New System.Drawing.Point(118, 80)
        Me.optUpdateCurrRow.Name = "optUpdateCurrRow"
        Me.optUpdateCurrRow.Size = New System.Drawing.Size(112, 16)
        Me.optUpdateCurrRow.TabIndex = 5
        Me.optUpdateCurrRow.TabStop = True
        Me.optUpdateCurrRow.Text = "Update Curr Row"
        '
        'optAddNewRow
        '
        Me.optAddNewRow.Location = New System.Drawing.Point(118, 104)
        Me.optAddNewRow.Name = "optAddNewRow"
        Me.optAddNewRow.Size = New System.Drawing.Size(104, 16)
        Me.optAddNewRow.TabIndex = 6
        Me.optAddNewRow.Text = "Add New Row"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(8, 163)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(77, 13)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "Override Price "
        '
        'textPrice
        '
        Me.textPrice.Location = New System.Drawing.Point(8, 178)
        Me.textPrice.Name = "textPrice"
        Me.textPrice.Size = New System.Drawing.Size(86, 20)
        Me.textPrice.TabIndex = 2
        Me.textPrice.Tag = "(No Auto Formatting)"
        Me.textPrice.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'frmCheckinMods
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(248, 206)
        Me.Controls.Add(Me.textPrice)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.optAddNewRow)
        Me.Controls.Add(Me.optUpdateCurrRow)
        Me.Controls.Add(Me.buttonSave)
        Me.Controls.Add(Me.buttonCancel)
        Me.Controls.Add(Me.cbPeriod)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.lblNP)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cbNbrPeriods)
        Me.Controls.Add(Me.lblEquipName)
        Me.Controls.Add(Me.lblEquipID)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmCheckinMods"
        Me.Text = "Modify Rental "
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region
   Public frm As frmCheckinNew
   Private formLoading As Boolean

   Private Sub frmCheckinMods_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
      With frm.dtList.Rows(frm.dbgShoppingList.CurrentCell.RowNumber)
         PositionForm(Me)
         Me.lblEquipID.Text = MNS(.Item("equip_id"))
         Me.lblEquipName.Text = MNS(.Item("equip_name"))
         Me.cbNbrPeriods.Text = MNI(.Item("quantity"))
         Me.cbPeriod.Text = MNS(.Item("rental_period"))
         formLoading = True
      End With
   End Sub

   Private Sub buttonCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonCancel.Click
      Me.Close()
      DoEvents()
   End Sub

   Private Sub buttonSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonSave.Click

      Try
         If Me.cbNbrPeriods.Text.Trim.Length = 0 OrElse _
            Me.cbPeriod.Text.Trim.Length = 0 Then
            MsgBox("You must select both a period and number of periods.", MsgBoxStyle.Information)
            Exit Sub
         End If
         Dim dr As DataRow = frm.dtList.Rows(frm.dbgShoppingList.CurrentCell.RowNumber)

         With dr
            If Me.optUpdateCurrRow.Checked Then
               dr("quantity") = Val(Me.cbNbrPeriods.Text)
               dr("rental_period") = Me.cbPeriod.Text
               If UnFormat(Me.textPrice.Text) > 0 Then
                  dr("priceperunit") = UnFormat(Me.textPrice.Text)
               Else
                  Dim o As New CCheckIn() ' just to use the method to get price
                  Dim price As Decimal = o.GetPriceForEquip(dr("equip_id"), Me.cbPeriod.Text)
                  dr("priceperunit") = price
               End If
            Else
               Dim price As Decimal
               If Me.textPrice.Text.Length > 0 Then
                  price = UnFormat(Me.textPrice.Text)
               Else
                  Dim o As New CCheckIn() ' just to use the method to get price
                  If MNB(dr("newprices")) Then
                     price = o.GetPriceFromCheckOutData(dr, Me.cbPeriod.Text)
                  Else
                     price = o.GetPriceForEquip(dr("equip_id"), Me.cbPeriod.Text)
                  End If
                  End If

                  Dim oCG As New CGrid()
                  Dim dRow() As Object = {dr("equip_id"), dr("equip_name"), Val(Me.cbNbrPeriods.Text), Me.cbPeriod.Text, price, Now.ToString, 0, 0}
                  oCG.AddRowToTable(frm.dtList, dRow)
            End If

            Me.Dispose()
            DoEvents()
         End With
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

   Private Sub frmCheckinMods_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
      'If formLoading Then
      '   formLoading = False
      '   Dim s As String = MNS(frm.dtList.Rows(frm.dbgShoppingList.CurrentCell.RowNumber).Item("rental_period"))
      '   DoEvents()
      '   If "Daily_Hourly_Half Day_Weekly_Monthly_Week End".IndexOf(s) = -1 Then
      '      MsgBox("Please select the desired row by click on it with the Left Mouse Button and then click on the row with the Right Mouse Button.", MsgBoxStyle.Information)
      '      Me.Dispose()
      '      DoEvents()
      '   End If
      'End If
   End Sub
   Private Sub textPrice_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles textPrice.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
   End Sub
   Private Sub textPrice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles textPrice.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub textPrice_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles textPrice.Enter
      textPrice.Text = Val(textPrice.Text)
      textPrice.SelectionStart = 0
      textPrice.SelectionLength = textPrice.Text.Trim.Length
   End Sub

   Private Sub textPrice_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles textPrice.Leave
      textPrice.Text = FormatCurrency(textPrice.Text)
   End Sub
End Class
