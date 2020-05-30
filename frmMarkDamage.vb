Imports System.Windows.Forms.Application
Public Class frmMarkDamage
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
   Friend WithEvents optDamageHold As System.Windows.Forms.RadioButton
   Friend WithEvents optDamageRent As System.Windows.Forms.RadioButton
   Friend WithEvents optRepaired As System.Windows.Forms.RadioButton
   Friend WithEvents textDamageDesc As System.Windows.Forms.TextBox
   Friend WithEvents Label3 As System.Windows.Forms.Label
   Friend WithEvents buttonCancel As System.Windows.Forms.Button
   Friend WithEvents buttonSave As System.Windows.Forms.Button
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMarkDamage))
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblEquipName = New System.Windows.Forms.Label()
        Me.lblEquipID = New System.Windows.Forms.Label()
        Me.optDamageHold = New System.Windows.Forms.RadioButton()
        Me.optDamageRent = New System.Windows.Forms.RadioButton()
        Me.optRepaired = New System.Windows.Forms.RadioButton()
        Me.textDamageDesc = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.buttonCancel = New System.Windows.Forms.Button()
        Me.buttonSave = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 45)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(65, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Equip Name"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 7)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(48, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Equip ID"
        '
        'lblEquipName
        '
        Me.lblEquipName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblEquipName.Location = New System.Drawing.Point(8, 62)
        Me.lblEquipName.Name = "lblEquipName"
        Me.lblEquipName.Size = New System.Drawing.Size(223, 16)
        Me.lblEquipName.TabIndex = 7
        '
        'lblEquipID
        '
        Me.lblEquipID.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblEquipID.Location = New System.Drawing.Point(9, 24)
        Me.lblEquipID.Name = "lblEquipID"
        Me.lblEquipID.Size = New System.Drawing.Size(88, 16)
        Me.lblEquipID.TabIndex = 6
        '
        'optDamageHold
        '
        Me.optDamageHold.Location = New System.Drawing.Point(8, 85)
        Me.optDamageHold.Name = "optDamageHold"
        Me.optDamageHold.Size = New System.Drawing.Size(104, 16)
        Me.optDamageHold.TabIndex = 8
        Me.optDamageHold.Text = "Damage Hold"
        '
        'optDamageRent
        '
        Me.optDamageRent.Location = New System.Drawing.Point(8, 109)
        Me.optDamageRent.Name = "optDamageRent"
        Me.optDamageRent.Size = New System.Drawing.Size(120, 16)
        Me.optDamageRent.TabIndex = 9
        Me.optDamageRent.Text = "Damage Rentable"
        '
        'optRepaired
        '
        Me.optRepaired.Location = New System.Drawing.Point(118, 85)
        Me.optRepaired.Name = "optRepaired"
        Me.optRepaired.Size = New System.Drawing.Size(112, 16)
        Me.optRepaired.TabIndex = 10
        Me.optRepaired.Text = "Repaired"
        '
        'textDamageDesc
        '
        Me.textDamageDesc.Location = New System.Drawing.Point(8, 149)
        Me.textDamageDesc.MaxLength = 50
        Me.textDamageDesc.Name = "textDamageDesc"
        Me.textDamageDesc.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.textDamageDesc.Size = New System.Drawing.Size(224, 20)
        Me.textDamageDesc.TabIndex = 11
        Me.textDamageDesc.Tag = "(No Auto Formatting)"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(8, 132)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(103, 13)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "Damage Description"
        '
        'buttonCancel
        '
        Me.buttonCancel.Location = New System.Drawing.Point(176, 186)
        Me.buttonCancel.Name = "buttonCancel"
        Me.buttonCancel.Size = New System.Drawing.Size(56, 24)
        Me.buttonCancel.TabIndex = 13
        Me.buttonCancel.Text = "&Cancel"
        '
        'buttonSave
        '
        Me.buttonSave.Location = New System.Drawing.Point(112, 186)
        Me.buttonSave.Name = "buttonSave"
        Me.buttonSave.Size = New System.Drawing.Size(56, 24)
        Me.buttonSave.TabIndex = 14
        Me.buttonSave.Text = "&Save"
        '
        'frmMarkDamage
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(240, 217)
        Me.Controls.Add(Me.buttonSave)
        Me.Controls.Add(Me.buttonCancel)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.textDamageDesc)
        Me.Controls.Add(Me.optRepaired)
        Me.Controls.Add(Me.optDamageRent)
        Me.Controls.Add(Me.optDamageHold)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblEquipName)
        Me.Controls.Add(Me.lblEquipID)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmMarkDamage"
        Me.Text = "Mark Damaged Equipment"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region
   Friend frm As frmRental

#Region " Form & Control Events "
   Private Sub frmMarkDamage_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
      With frm.dtEquip.Rows(frm.miHitRow)
         PositionForm(Me)
         Me.lblEquipID.Text = MNS(.Item("equip_id"))
         Me.lblEquipName.Text = MNS(.Item("equip_name"))
         If IsDBNull(.Item("damaged")) OrElse .Item("damaged") = String.Empty Then
            Me.optRepaired.Checked = True
         ElseIf .Item("Damaged") = "H" Then
            Me.optDamageHold.Checked = True
         Else
            Me.optDamageRent.Checked = True
         End If
         Me.textDamageDesc.Text = MNS(.Item("damage_desc"))
      End With
   End Sub

   Private Sub buttonCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles buttonCancel.Click
      Me.Close()
      DoEvents()
   End Sub

   Private Sub buttonSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles buttonSave.Click
      Dim sql As String
      Dim oDA As New CDataAccess()
      Dim serr As String = String.Empty

      sql = "update equipment set damage = "
      With frm.dtEquip.Rows(frm.miHitRow)

         If optRepaired.Checked Then
            sql &= "'', damage_desc = '' "
         ElseIf optDamageHold.Checked Then
            sql &= "'H', damage_desc = '" & MNS(Me.textDamageDesc.Text) & "' "
         Else
            sql &= "'R', damage_desc = '" & MNS(Me.textDamageDesc.Text) & "' "
         End If
         sql &= "where equip_id = '" & .Item("equip_id") & "' "
      End With
      If oDA.SendActionSql(sql, ConnectString, serr) = 0 Then
         MsgBox("Update of database failed on error: " & Chr(10) & serr, MsgBoxStyle.Critical)
         Exit Sub
      End If

      Me.Close()
      DoEvents()
   End Sub
   Private Sub textDamageDesc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles textDamageDesc.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
   End Sub
   Private Sub textDamageDesc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles textDamageDesc.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub textDamageDesc_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles textDamageDesc.Enter
      textDamageDesc.SelectionStart = 0
      textDamageDesc.SelectionLength = textDamageDesc.Text.Trim.Length
   End Sub


#End Region

End Class
