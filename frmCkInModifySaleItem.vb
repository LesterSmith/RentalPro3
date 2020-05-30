Imports System.Windows.Forms.Application
Public Class frmCkInModifySaleItem
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
   Friend WithEvents Label2 As System.Windows.Forms.Label
   Friend WithEvents Label1 As System.Windows.Forms.Label
   Friend WithEvents lblEquipName As System.Windows.Forms.Label
   Friend WithEvents lblEquipID As System.Windows.Forms.Label
   Friend WithEvents buttonSave As System.Windows.Forms.Button
   Friend WithEvents buttonCancel As System.Windows.Forms.Button
   Friend WithEvents textQty As System.Windows.Forms.TextBox
   Friend WithEvents Label4 As System.Windows.Forms.Label
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
      Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmCkInModifySaleItem))
      Me.Label2 = New System.Windows.Forms.Label()
      Me.Label1 = New System.Windows.Forms.Label()
      Me.lblEquipName = New System.Windows.Forms.Label()
      Me.lblEquipID = New System.Windows.Forms.Label()
      Me.buttonSave = New System.Windows.Forms.Button()
      Me.buttonCancel = New System.Windows.Forms.Button()
      Me.textQty = New System.Windows.Forms.TextBox()
      Me.Label4 = New System.Windows.Forms.Label()
      Me.SuspendLayout()
      '
      'Label2
      '
      Me.Label2.AutoSize = True
      Me.Label2.Location = New System.Drawing.Point(8, 48)
      Me.Label2.Name = "Label2"
      Me.Label2.Size = New System.Drawing.Size(66, 13)
      Me.Label2.TabIndex = 5
      Me.Label2.Text = "Equip Name"
      '
      'Label1
      '
      Me.Label1.AutoSize = True
      Me.Label1.Location = New System.Drawing.Point(8, 8)
      Me.Label1.Name = "Label1"
      Me.Label1.Size = New System.Drawing.Size(48, 13)
      Me.Label1.TabIndex = 3
      Me.Label1.Text = "Equip ID"
      '
      'lblEquipName
      '
      Me.lblEquipName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
      Me.lblEquipName.Location = New System.Drawing.Point(8, 64)
      Me.lblEquipName.Name = "lblEquipName"
      Me.lblEquipName.Size = New System.Drawing.Size(223, 16)
      Me.lblEquipName.TabIndex = 6
      '
      'lblEquipID
      '
      Me.lblEquipID.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
      Me.lblEquipID.Location = New System.Drawing.Point(8, 24)
      Me.lblEquipID.Name = "lblEquipID"
      Me.lblEquipID.Size = New System.Drawing.Size(88, 16)
      Me.lblEquipID.TabIndex = 4
      '
      'buttonSave
      '
      Me.buttonSave.Location = New System.Drawing.Point(144, 104)
      Me.buttonSave.Name = "buttonSave"
      Me.buttonSave.Size = New System.Drawing.Size(72, 24)
      Me.buttonSave.TabIndex = 1
      Me.buttonSave.Text = "&Save"
      '
      'buttonCancel
      '
      Me.buttonCancel.Location = New System.Drawing.Point(144, 136)
      Me.buttonCancel.Name = "buttonCancel"
      Me.buttonCancel.Size = New System.Drawing.Size(72, 24)
      Me.buttonCancel.TabIndex = 2
      Me.buttonCancel.Text = "&Cancel"
      '
      'textQty
      '
      Me.textQty.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(255, Byte))
      Me.textQty.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(0, Byte))
      Me.textQty.Location = New System.Drawing.Point(8, 104)
      Me.textQty.Name = "textQty"
      Me.textQty.Size = New System.Drawing.Size(86, 20)
      Me.textQty.TabIndex = 0
      Me.textQty.Tag = "(No Auto Formatting)"
      Me.textQty.Text = ""
      Me.textQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      '
      'Label4
      '
      Me.Label4.AutoSize = True
      Me.Label4.Location = New System.Drawing.Point(8, 88)
      Me.Label4.Name = "Label4"
      Me.Label4.Size = New System.Drawing.Size(81, 13)
      Me.Label4.TabIndex = 7
      Me.Label4.Text = "Enter Qty Used"
      '
      'frmCkInModifySaleItem
      '
      Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
      Me.ClientSize = New System.Drawing.Size(240, 166)
      Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.textQty, Me.Label4, Me.buttonSave, Me.buttonCancel, Me.Label2, Me.Label1, Me.lblEquipName, Me.lblEquipID})
      Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
      Me.Name = "frmCkInModifySaleItem"
      Me.Text = "Modify Sale Item"
      Me.ResumeLayout(False)

   End Sub

#End Region
   Public frm As frmCheckinNew
   Private origQty As Integer

   Private Sub frmCkInModifySaleItem_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
      With frm.dtList.Rows(frm.dbgShoppingList.CurrentCell.RowNumber)
         PositionForm(Me)
         Me.lblEquipID.Text = MNS(.Item("equip_id"))
         Me.lblEquipName.Text = MNS(.Item("equip_name"))
         Me.textQty.Text = MNI(.Item("quantity"))
         origQty = MNI(.Item("quantity"))
      End With
   End Sub

   Private Sub buttonCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonCancel.Click
      Me.Close()
      DoEvents()
   End Sub
   Private Sub textQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles textQty.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
   End Sub
   Private Sub textQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles textQty.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub textQty_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles textQty.Enter
      textQty.SelectionStart = 0
      textQty.SelectionLength = textQty.Text.Trim.Length
   End Sub

   Private Sub buttonSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonSave.Click
      Try
         Dim qty As Integer = Val(Me.textQty.Text)
         qty -= origQty
         Dim serr As String
         Dim sql As String
         sql = "update products set unitsinstock = unitsinstock - " & qty
         sql &= " where productid = '" & Me.lblEquipID.Text & "' "
         Dim oda As New CDataAccess()
         Dim i As Integer = oda.SendActionSql(sql, ConnectString, serr)
         If i <> 1 Then
            MsgBox("Update of inventory failed.", MsgBoxStyle.Exclamation)
            Exit Sub
         End If
         Dim row As Integer = frm.dbgShoppingList.CurrentCell.RowNumber
         With frm.dtList.Rows(row)
            .Item("quantity") = Val(Me.textQty.Text)
         End With
         Me.Close()
         DoEvents()
      Catch
      End Try
   End Sub
End Class
