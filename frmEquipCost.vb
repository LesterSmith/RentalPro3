Public Class frmEquipCost
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
   Friend WithEvents cbCostType As System.Windows.Forms.ComboBox
   Public WithEvents cmdAdd As System.Windows.Forms.Button
   Public WithEvents cmdClose As System.Windows.Forms.Button
   Public WithEvents cmdDelete As System.Windows.Forms.Button
   Public WithEvents cmdUpdate As System.Windows.Forms.Button
   Friend WithEvents dgEquip As System.Windows.Forms.DataGrid
   Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
   Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
   Friend WithEvents Label1 As System.Windows.Forms.Label
   Friend WithEvents Label10 As System.Windows.Forms.Label
   Friend WithEvents Label11 As System.Windows.Forms.Label
   Friend WithEvents Label12 As System.Windows.Forms.Label
   Friend WithEvents Label13 As System.Windows.Forms.Label
   Friend WithEvents Label2 As System.Windows.Forms.Label
   Friend WithEvents Label3 As System.Windows.Forms.Label
   Friend WithEvents Label4 As System.Windows.Forms.Label
   Friend WithEvents Label5 As System.Windows.Forms.Label
   Friend WithEvents Label6 As System.Windows.Forms.Label
   Friend WithEvents Label7 As System.Windows.Forms.Label
   Friend WithEvents Label8 As System.Windows.Forms.Label
   Friend WithEvents Label9 As System.Windows.Forms.Label
   Friend WithEvents lblBalanceDue As System.Windows.Forms.Label
   Friend WithEvents lblDamageLoss As System.Windows.Forms.Label
   Friend WithEvents lblLabor As System.Windows.Forms.Label
   Friend WithEvents lblNetGainLoss As System.Windows.Forms.Label
   Friend WithEvents lblPurchasePrice As System.Windows.Forms.Label
   Friend WithEvents lblRentalIncome As System.Windows.Forms.Label
   Friend WithEvents lblSupplies As System.Windows.Forms.Label
   Friend WithEvents lblTotalCost As System.Windows.Forms.Label
   Friend WithEvents lblTotalPayments As System.Windows.Forms.Label
   Friend WithEvents tabAddCost As System.Windows.Forms.TabPage
   Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
   Friend WithEvents tabEditCost As System.Windows.Forms.TabPage
   Friend WithEvents tabTotalCost As System.Windows.Forms.TabPage
   Friend WithEvents txtCost As System.Windows.Forms.TextBox
   Friend WithEvents txtCostDesc As System.Windows.Forms.TextBox
   Friend WithEvents txtMechanic As System.Windows.Forms.TextBox
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
      Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmEquipCost))
      Me.dgEquip = New System.Windows.Forms.DataGrid()
      Me.cbCostType = New System.Windows.Forms.ComboBox()
      Me.GroupBox1 = New System.Windows.Forms.GroupBox()
      Me.txtMechanic = New System.Windows.Forms.TextBox()
      Me.Label4 = New System.Windows.Forms.Label()
      Me.txtCost = New System.Windows.Forms.TextBox()
      Me.Label3 = New System.Windows.Forms.Label()
      Me.Label2 = New System.Windows.Forms.Label()
      Me.txtCostDesc = New System.Windows.Forms.TextBox()
      Me.Label1 = New System.Windows.Forms.Label()
      Me.cmdAdd = New System.Windows.Forms.Button()
      Me.cmdUpdate = New System.Windows.Forms.Button()
      Me.cmdDelete = New System.Windows.Forms.Button()
      Me.cmdClose = New System.Windows.Forms.Button()
      Me.TabControl1 = New System.Windows.Forms.TabControl()
      Me.tabAddCost = New System.Windows.Forms.TabPage()
      Me.tabTotalCost = New System.Windows.Forms.TabPage()
      Me.lblPurchasePrice = New System.Windows.Forms.Label()
      Me.Label13 = New System.Windows.Forms.Label()
      Me.Label12 = New System.Windows.Forms.Label()
      Me.Label11 = New System.Windows.Forms.Label()
      Me.Label10 = New System.Windows.Forms.Label()
      Me.Label9 = New System.Windows.Forms.Label()
      Me.Label8 = New System.Windows.Forms.Label()
      Me.Label7 = New System.Windows.Forms.Label()
      Me.Label6 = New System.Windows.Forms.Label()
      Me.Label5 = New System.Windows.Forms.Label()
      Me.lblTotalPayments = New System.Windows.Forms.Label()
      Me.lblLabor = New System.Windows.Forms.Label()
      Me.lblSupplies = New System.Windows.Forms.Label()
      Me.lblDamageLoss = New System.Windows.Forms.Label()
      Me.lblRentalIncome = New System.Windows.Forms.Label()
      Me.lblBalanceDue = New System.Windows.Forms.Label()
      Me.lblTotalCost = New System.Windows.Forms.Label()
      Me.lblNetGainLoss = New System.Windows.Forms.Label()
      Me.tabEditCost = New System.Windows.Forms.TabPage()
      Me.GroupBox2 = New System.Windows.Forms.GroupBox()
      CType(Me.dgEquip, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.GroupBox1.SuspendLayout()
      Me.TabControl1.SuspendLayout()
      Me.tabAddCost.SuspendLayout()
      Me.tabTotalCost.SuspendLayout()
      Me.tabEditCost.SuspendLayout()
      Me.GroupBox2.SuspendLayout()
      Me.SuspendLayout()
      '
      'dgEquip
      '
      Me.dgEquip.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                  Or System.Windows.Forms.AnchorStyles.Left) _
                  Or System.Windows.Forms.AnchorStyles.Right)
      Me.dgEquip.CaptionText = "Rental Equipment"
      Me.dgEquip.DataMember = ""
      Me.dgEquip.HeaderForeColor = System.Drawing.SystemColors.ControlText
      Me.dgEquip.Location = New System.Drawing.Point(8, 16)
      Me.dgEquip.Name = "dgEquip"
      Me.dgEquip.Size = New System.Drawing.Size(395, 253)
      Me.dgEquip.TabIndex = 0
      '
      'cbCostType
      '
      Me.cbCostType.Items.AddRange(New Object() {"1 - Labor", "2 - Supplies", "3 - Damage Recovered", "4 - Damage Lost", "8 - Loan Payment", "9 - Purchase Price"})
      Me.cbCostType.Location = New System.Drawing.Point(72, 38)
      Me.cbCostType.Name = "cbCostType"
      Me.cbCostType.Size = New System.Drawing.Size(152, 21)
      Me.cbCostType.TabIndex = 1
      '
      'GroupBox1
      '
      Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtMechanic, Me.Label4, Me.txtCost, Me.Label3, Me.Label2, Me.txtCostDesc, Me.Label1, Me.cbCostType})
      Me.GroupBox1.Location = New System.Drawing.Point(6, 6)
      Me.GroupBox1.Name = "GroupBox1"
      Me.GroupBox1.Size = New System.Drawing.Size(296, 112)
      Me.GroupBox1.TabIndex = 2
      Me.GroupBox1.TabStop = False
      Me.GroupBox1.Text = "Enter Cost Data"
      '
      'txtMechanic
      '
      Me.txtMechanic.Location = New System.Drawing.Point(72, 83)
      Me.txtMechanic.MaxLength = 50
      Me.txtMechanic.Name = "txtMechanic"
      Me.txtMechanic.Size = New System.Drawing.Size(168, 20)
      Me.txtMechanic.TabIndex = 6
      Me.txtMechanic.Tag = "(No Auto Formatting)"
      Me.txtMechanic.Text = ""
      '
      'Label4
      '
      Me.Label4.AutoSize = True
      Me.Label4.Location = New System.Drawing.Point(8, 87)
      Me.Label4.Name = "Label4"
      Me.Label4.Size = New System.Drawing.Size(53, 13)
      Me.Label4.TabIndex = 5
      Me.Label4.Text = "Mechanic"
      '
      'txtCost
      '
      Me.txtCost.Location = New System.Drawing.Point(72, 61)
      Me.txtCost.Name = "txtCost"
      Me.txtCost.Size = New System.Drawing.Size(72, 20)
      Me.txtCost.TabIndex = 4
      Me.txtCost.Tag = "$#,##0.00;($#,##0.00)"
      Me.txtCost.Text = ""
      Me.txtCost.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
      '
      'Label3
      '
      Me.Label3.AutoSize = True
      Me.Label3.Location = New System.Drawing.Point(32, 64)
      Me.Label3.Name = "Label3"
      Me.Label3.Size = New System.Drawing.Size(27, 13)
      Me.Label3.TabIndex = 3
      Me.Label3.Text = "Cost"
      '
      'Label2
      '
      Me.Label2.AutoSize = True
      Me.Label2.Location = New System.Drawing.Point(8, 40)
      Me.Label2.Name = "Label2"
      Me.Label2.Size = New System.Drawing.Size(56, 13)
      Me.Label2.TabIndex = 2
      Me.Label2.Text = "Cost Type"
      '
      'txtCostDesc
      '
      Me.txtCostDesc.Location = New System.Drawing.Point(72, 16)
      Me.txtCostDesc.MaxLength = 50
      Me.txtCostDesc.Name = "txtCostDesc"
      Me.txtCostDesc.Size = New System.Drawing.Size(216, 20)
      Me.txtCostDesc.TabIndex = 1
      Me.txtCostDesc.Tag = "(No Auto Formatting)"
      Me.txtCostDesc.Text = ""
      '
      'Label1
      '
      Me.Label1.AutoSize = True
      Me.Label1.Location = New System.Drawing.Point(8, 16)
      Me.Label1.Name = "Label1"
      Me.Label1.Size = New System.Drawing.Size(56, 13)
      Me.Label1.TabIndex = 0
      Me.Label1.Text = "Cost Desc"
      '
      'cmdAdd
      '
      Me.cmdAdd.BackColor = System.Drawing.SystemColors.Control
      Me.cmdAdd.Cursor = System.Windows.Forms.Cursors.Default
      Me.cmdAdd.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.cmdAdd.ForeColor = System.Drawing.SystemColors.ControlText
      Me.cmdAdd.Location = New System.Drawing.Point(310, 14)
      Me.cmdAdd.Name = "cmdAdd"
      Me.cmdAdd.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.cmdAdd.Size = New System.Drawing.Size(77, 26)
      Me.cmdAdd.TabIndex = 15
      Me.cmdAdd.Text = "&Add"
      '
      'cmdUpdate
      '
      Me.cmdUpdate.BackColor = System.Drawing.SystemColors.Control
      Me.cmdUpdate.Cursor = System.Windows.Forms.Cursors.Default
      Me.cmdUpdate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.cmdUpdate.ForeColor = System.Drawing.SystemColors.ControlText
      Me.cmdUpdate.Location = New System.Drawing.Point(304, 72)
      Me.cmdUpdate.Name = "cmdUpdate"
      Me.cmdUpdate.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.cmdUpdate.Size = New System.Drawing.Size(77, 26)
      Me.cmdUpdate.TabIndex = 14
      Me.cmdUpdate.Text = "&Save"
      '
      'cmdDelete
      '
      Me.cmdDelete.BackColor = System.Drawing.SystemColors.Control
      Me.cmdDelete.Cursor = System.Windows.Forms.Cursors.Default
      Me.cmdDelete.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.cmdDelete.ForeColor = System.Drawing.SystemColors.ControlText
      Me.cmdDelete.Location = New System.Drawing.Point(317, 240)
      Me.cmdDelete.Name = "cmdDelete"
      Me.cmdDelete.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.cmdDelete.Size = New System.Drawing.Size(77, 26)
      Me.cmdDelete.TabIndex = 16
      Me.cmdDelete.Text = "&Delete"
      '
      'cmdClose
      '
      Me.cmdClose.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
      Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
      Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
      Me.cmdClose.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
      Me.cmdClose.Location = New System.Drawing.Point(336, 445)
      Me.cmdClose.Name = "cmdClose"
      Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.cmdClose.Size = New System.Drawing.Size(77, 26)
      Me.cmdClose.TabIndex = 17
      Me.cmdClose.Text = "&Close"
      '
      'TabControl1
      '
      Me.TabControl1.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
      Me.TabControl1.Controls.AddRange(New System.Windows.Forms.Control() {Me.tabTotalCost, Me.tabAddCost, Me.tabEditCost})
      Me.TabControl1.Location = New System.Drawing.Point(5, 283)
      Me.TabControl1.Name = "TabControl1"
      Me.TabControl1.SelectedIndex = 0
      Me.TabControl1.Size = New System.Drawing.Size(408, 151)
      Me.TabControl1.TabIndex = 18
      '
      'tabAddCost
      '
      Me.tabAddCost.Controls.AddRange(New System.Windows.Forms.Control() {Me.GroupBox1, Me.cmdAdd})
      Me.tabAddCost.Location = New System.Drawing.Point(4, 22)
      Me.tabAddCost.Name = "tabAddCost"
      Me.tabAddCost.Size = New System.Drawing.Size(400, 125)
      Me.tabAddCost.TabIndex = 0
      Me.tabAddCost.Text = "Add Cost Items"
      '
      'tabTotalCost
      '
      Me.tabTotalCost.Controls.AddRange(New System.Windows.Forms.Control() {Me.lblPurchasePrice, Me.Label13, Me.Label12, Me.Label11, Me.Label10, Me.Label9, Me.Label8, Me.Label7, Me.Label6, Me.Label5, Me.cmdDelete, Me.lblTotalPayments, Me.lblLabor, Me.lblSupplies, Me.lblDamageLoss, Me.lblRentalIncome, Me.lblBalanceDue, Me.lblTotalCost, Me.lblNetGainLoss})
      Me.tabTotalCost.Location = New System.Drawing.Point(4, 22)
      Me.tabTotalCost.Name = "tabTotalCost"
      Me.tabTotalCost.Size = New System.Drawing.Size(400, 125)
      Me.tabTotalCost.TabIndex = 1
      Me.tabTotalCost.Text = "Total Cost of Ownership"
      '
      'lblPurchasePrice
      '
      Me.lblPurchasePrice.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
      Me.lblPurchasePrice.Location = New System.Drawing.Point(96, 8)
      Me.lblPurchasePrice.Name = "lblPurchasePrice"
      Me.lblPurchasePrice.Size = New System.Drawing.Size(90, 16)
      Me.lblPurchasePrice.TabIndex = 26
      Me.lblPurchasePrice.TextAlign = System.Drawing.ContentAlignment.MiddleRight
      '
      'Label13
      '
      Me.Label13.AutoSize = True
      Me.Label13.Location = New System.Drawing.Point(200, 96)
      Me.Label13.Name = "Label13"
      Me.Label13.Size = New System.Drawing.Size(76, 13)
      Me.Label13.TabIndex = 25
      Me.Label13.Text = "Net Gain/Loss"
      '
      'Label12
      '
      Me.Label12.AutoSize = True
      Me.Label12.Location = New System.Drawing.Point(8, 97)
      Me.Label12.Name = "Label12"
      Me.Label12.Size = New System.Drawing.Size(77, 13)
      Me.Label12.TabIndex = 24
      Me.Label12.Text = "Rental Income"
      '
      'Label11
      '
      Me.Label11.AutoSize = True
      Me.Label11.Location = New System.Drawing.Point(220, 78)
      Me.Label11.Name = "Label11"
      Me.Label11.Size = New System.Drawing.Size(56, 13)
      Me.Label11.TabIndex = 23
      Me.Label11.Text = "Total Cost"
      '
      'Label10
      '
      Me.Label10.AutoSize = True
      Me.Label10.Location = New System.Drawing.Point(8, 78)
      Me.Label10.Name = "Label10"
      Me.Label10.Size = New System.Drawing.Size(74, 13)
      Me.Label10.TabIndex = 22
      Me.Label10.Text = "Damage Loss"
      '
      'Label9
      '
      Me.Label9.AutoSize = True
      Me.Label9.Location = New System.Drawing.Point(232, 24)
      Me.Label9.Name = "Label9"
      Me.Label9.Size = New System.Drawing.Size(44, 13)
      Me.Label9.TabIndex = 21
      Me.Label9.Text = "Bal Due"
      '
      'Label8
      '
      Me.Label8.AutoSize = True
      Me.Label8.Location = New System.Drawing.Point(8, 61)
      Me.Label8.Name = "Label8"
      Me.Label8.Size = New System.Drawing.Size(79, 13)
      Me.Label8.TabIndex = 20
      Me.Label8.Text = "Maint Supplies"
      '
      'Label7
      '
      Me.Label7.AutoSize = True
      Me.Label7.Location = New System.Drawing.Point(8, 44)
      Me.Label7.Name = "Label7"
      Me.Label7.Size = New System.Drawing.Size(64, 13)
      Me.Label7.TabIndex = 19
      Me.Label7.Text = "Maint Labor"
      '
      'Label6
      '
      Me.Label6.AutoSize = True
      Me.Label6.Location = New System.Drawing.Point(8, 27)
      Me.Label6.Name = "Label6"
      Me.Label6.Size = New System.Drawing.Size(83, 13)
      Me.Label6.TabIndex = 18
      Me.Label6.Text = "Total Payments"
      '
      'Label5
      '
      Me.Label5.AutoSize = True
      Me.Label5.Location = New System.Drawing.Point(8, 8)
      Me.Label5.Name = "Label5"
      Me.Label5.Size = New System.Drawing.Size(81, 13)
      Me.Label5.TabIndex = 17
      Me.Label5.Text = "Purchase Price"
      '
      'lblTotalPayments
      '
      Me.lblTotalPayments.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
      Me.lblTotalPayments.Location = New System.Drawing.Point(96, 26)
      Me.lblTotalPayments.Name = "lblTotalPayments"
      Me.lblTotalPayments.Size = New System.Drawing.Size(90, 16)
      Me.lblTotalPayments.TabIndex = 27
      Me.lblTotalPayments.TextAlign = System.Drawing.ContentAlignment.MiddleRight
      '
      'lblLabor
      '
      Me.lblLabor.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
      Me.lblLabor.Location = New System.Drawing.Point(96, 44)
      Me.lblLabor.Name = "lblLabor"
      Me.lblLabor.Size = New System.Drawing.Size(90, 16)
      Me.lblLabor.TabIndex = 27
      Me.lblLabor.TextAlign = System.Drawing.ContentAlignment.MiddleRight
      '
      'lblSupplies
      '
      Me.lblSupplies.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
      Me.lblSupplies.Location = New System.Drawing.Point(96, 62)
      Me.lblSupplies.Name = "lblSupplies"
      Me.lblSupplies.Size = New System.Drawing.Size(90, 16)
      Me.lblSupplies.TabIndex = 27
      Me.lblSupplies.TextAlign = System.Drawing.ContentAlignment.MiddleRight
      '
      'lblDamageLoss
      '
      Me.lblDamageLoss.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
      Me.lblDamageLoss.Location = New System.Drawing.Point(96, 80)
      Me.lblDamageLoss.Name = "lblDamageLoss"
      Me.lblDamageLoss.Size = New System.Drawing.Size(90, 16)
      Me.lblDamageLoss.TabIndex = 27
      Me.lblDamageLoss.TextAlign = System.Drawing.ContentAlignment.MiddleRight
      '
      'lblRentalIncome
      '
      Me.lblRentalIncome.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
      Me.lblRentalIncome.Location = New System.Drawing.Point(96, 98)
      Me.lblRentalIncome.Name = "lblRentalIncome"
      Me.lblRentalIncome.Size = New System.Drawing.Size(90, 16)
      Me.lblRentalIncome.TabIndex = 27
      Me.lblRentalIncome.TextAlign = System.Drawing.ContentAlignment.MiddleRight
      '
      'lblBalanceDue
      '
      Me.lblBalanceDue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
      Me.lblBalanceDue.Location = New System.Drawing.Point(283, 25)
      Me.lblBalanceDue.Name = "lblBalanceDue"
      Me.lblBalanceDue.Size = New System.Drawing.Size(86, 16)
      Me.lblBalanceDue.TabIndex = 27
      Me.lblBalanceDue.TextAlign = System.Drawing.ContentAlignment.MiddleRight
      '
      'lblTotalCost
      '
      Me.lblTotalCost.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
      Me.lblTotalCost.Location = New System.Drawing.Point(283, 78)
      Me.lblTotalCost.Name = "lblTotalCost"
      Me.lblTotalCost.Size = New System.Drawing.Size(86, 16)
      Me.lblTotalCost.TabIndex = 27
      Me.lblTotalCost.TextAlign = System.Drawing.ContentAlignment.MiddleRight
      '
      'lblNetGainLoss
      '
      Me.lblNetGainLoss.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
      Me.lblNetGainLoss.Location = New System.Drawing.Point(283, 96)
      Me.lblNetGainLoss.Name = "lblNetGainLoss"
      Me.lblNetGainLoss.Size = New System.Drawing.Size(86, 16)
      Me.lblNetGainLoss.TabIndex = 27
      Me.lblNetGainLoss.TextAlign = System.Drawing.ContentAlignment.MiddleRight
      '
      'tabEditCost
      '
      Me.tabEditCost.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdUpdate})
      Me.tabEditCost.Location = New System.Drawing.Point(4, 22)
      Me.tabEditCost.Name = "tabEditCost"
      Me.tabEditCost.Size = New System.Drawing.Size(400, 125)
      Me.tabEditCost.TabIndex = 2
      Me.tabEditCost.Text = "Edit Cost Items"
      '
      'GroupBox2
      '
      Me.GroupBox2.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                  Or System.Windows.Forms.AnchorStyles.Left) _
                  Or System.Windows.Forms.AnchorStyles.Right)
      Me.GroupBox2.Controls.AddRange(New System.Windows.Forms.Control() {Me.dgEquip})
      Me.GroupBox2.Location = New System.Drawing.Point(1, 2)
      Me.GroupBox2.Name = "GroupBox2"
      Me.GroupBox2.Size = New System.Drawing.Size(411, 277)
      Me.GroupBox2.TabIndex = 19
      Me.GroupBox2.TabStop = False
      Me.GroupBox2.Text = "Select Equipment"
      '
      'frmEquipCost
      '
      Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
      Me.ClientSize = New System.Drawing.Size(421, 478)
      Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.GroupBox2, Me.TabControl1, Me.cmdClose})
      Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
      Me.Name = "frmEquipCost"
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "Maintain Total Cost of Ownership"
      CType(Me.dgEquip, System.ComponentModel.ISupportInitialize).EndInit()
      Me.GroupBox1.ResumeLayout(False)
      Me.TabControl1.ResumeLayout(False)
      Me.tabAddCost.ResumeLayout(False)
      Me.tabTotalCost.ResumeLayout(False)
      Me.tabEditCost.ResumeLayout(False)
      Me.GroupBox2.ResumeLayout(False)
      Me.ResumeLayout(False)

   End Sub

#End Region
#Region " Module Variables "
   Private mbDirty As Boolean
   Private msAddOrEdit As String
   Private oDA As New CDataAccess()
   Private equipHitRow As Integer
   Private costHitRow As Integer
   Private dtEquip As DataTable
   Private oCG As New CGrid()


#End Region

#Region " Private Methods "
   Private Sub LoadEquipGrid()
      Dim SQL As String

      Try
         dtEquip = New DataTable("dt")

         SQL = "select equip_name, equip_id "
         SQL &= "from equipment "
         SQL &= " order by Equip_name"
         Dim Formats() As String = _
            {"", "200", "T", "L", _
            "", "60", "T", "L"}

         If oDA.SendQuery(SQL, dtEquip, modMain.ConnectString, "dt") > 0 Then
            oCG.SetTablesStyle("Select", dtEquip, Me.dgEquip, Formats)
            Me.dgEquip.SetDataBinding(dtEquip, "")
            oCG.DisableAddNew(dgEquip, Me)
         End If

      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
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



   ''' <summary>
   ''' Compute and display Total Cost of Ownership
   ''' </summary>
   Private Sub DisplayTCO()
      Dim sql As String
      Dim dt As New DataTable()
      Dim i As Integer
      Dim tco As Decimal = 0
      Dim maintLabor As Decimal = 0
      Dim maintSupplies As Decimal = 0
      Dim maintLoss As Decimal = 0
      Dim rentalIncome As Decimal = 0
      Dim totalPayments As Decimal = 0

      With Me
         .lblBalanceDue.Text = FormatCurrency(0)
         .lblDamageLoss.Text = FormatCurrency(0)
         .lblLabor.Text = FormatCurrency(0)
         .lblNetGainLoss.Text = FormatCurrency(0)
         .lblPurchasePrice.Text = FormatCurrency(0)
         .lblRentalIncome.Text = FormatCurrency(0)
         .lblSupplies.Text = FormatCurrency(0)
         .lblTotalCost.Text = FormatCurrency(0)
         .lblTotalPayments.Text = FormatCurrency(0)

      End With
      sql = "select * from equip_cost "
      sql &= "where equip_id = '" & dtEquip.Rows(equipHitRow).Item("equip_id") & "' "
      sql &= "order by cost_type"
      If oDA.SendQuery(sql, dt, ConnectString) > 0 Then
         For i = 0 To dt.Rows.Count - 1
            With dt.Rows(i)
               Select Case .Item("cost_type")
                  Case 9
                     lblPurchasePrice.Text = FormatCurrency(MND(.Item("cost_price")))
                  Case 1
                     maintLabor += MND(.Item("cost_price"))
                     tco += MND(.Item("cost_price"))
                  Case 2
                     maintSupplies += MND(.Item("cost_price"))
                     tco += MND(.Item("cost_price"))
                  Case 4
                     maintLoss += MND(.Item("cost_price"))
                     tco += MND(.Item("cost_price"))
                  Case 8
                     totalPayments += MND(.Item("cost_price"))
                     tco += MND(.Item("cost_price"))
               End Select
            End With
         Next
      End If
      Me.lblTotalCost.Text = FormatCurrency(tco)
      Me.lblBalanceDue.Text = FormatCurrency(UnFormat(Me.lblPurchasePrice.Text) - totalPayments)
      Me.lblDamageLoss.Text = FormatCurrency(maintLoss)
      Me.lblSupplies.Text = FormatCurrency(maintSupplies)
      Me.lblLabor.Text = FormatCurrency(maintLabor)
      Me.lblTotalPayments.Text = FormatCurrency(totalPayments)

      ' get rental income
      dt.Reset()
      sql = "select quantity,priceperunit "
      sql &= "from invoice_details "
      sql &= "where equip_id = '" & dtEquip.Rows(equipHitRow).Item("equip_id") & "' "
      If oDA.SendQuery(sql, dt, ConnectString) > 0 Then
         For i = 0 To dt.Rows.Count - 1
            With dt.Rows(0)
               rentalIncome += MND(.Item("quantity")) * _
                  MND(.Item("priceperunit"))
            End With
         Next
         Me.lblRentalIncome.Text = FormatCurrency(rentalIncome)
      End If
      Me.lblNetGainLoss.Text = FormatCurrency(rentalIncome - tco)

   End Sub

   Private Sub UpdateCostTable()
      Dim sql As String
      With Me
         'sql = "update equip_cost "
         'sql &= "set cost_desc = '" & Replace(.txtCostDesc.Text, "'", "''") & "', "
         'sql &= "cost_price = " & UnFormat(.txtCost.Text) & ", "
         'sql &= "mechanic = '" & Replace(.txtMechanic.Text, "'", "''") & ", "
         'sql &= "cost_type = " & sPID & " "
         'sql &= "where unique_id = '" & dtEquip.Rows(miHitRow).Item("unique_id") & "'"
      End With
   End Sub

#End Region

#Region " Form & Control Events "
   Private Sub txtCost_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCost.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
      e.Handled = CkKeyPressNumeric(Asc(Val(e.KeyChar)), txtCost) = 0
   End Sub
   Private Sub txtCost_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCost.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub txtCost_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCost.Enter
      txtCost.Text = UnFmt_T_B(txtCost)
      txtCost.SelectionStart = 0
      txtCost.SelectionLength = txtCost.Text.Trim.Length
   End Sub
   Private Sub txtCost_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCost.Leave
      txtCost.Text = Fmt_T_B(txtCost)
   End Sub
   Private Sub cmdUpdate_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdUpdate.Click
   End Sub


   Private Sub txtCostDesc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCostDesc.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
   End Sub
   Private Sub txtCostDesc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCostDesc.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub txtCostDesc_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCostDesc.Enter
      txtCostDesc.SelectionStart = 0
      txtCostDesc.SelectionLength = txtCostDesc.Text.Trim.Length
   End Sub
   Private Sub txtMechanic_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMechanic.KeyPress
      If e.KeyChar = Chr(13) Then
         e.Handled = True
         Exit Sub
      End If
   End Sub
   Private Sub txtMechanic_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtMechanic.KeyDown
      If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
      If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
      If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
   End Sub
   Private Sub txtMechanic_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMechanic.Enter
      txtMechanic.SelectionStart = 0
      txtMechanic.SelectionLength = txtMechanic.Text.Trim.Length
   End Sub

   Private Sub frmEquipCost_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
      LoadEquipGrid()
   End Sub

   Private Sub dgEquip_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgEquip.MouseUp
      Dim b As Boolean
      oCG.UncheckAllBoxes(dtEquip, "Select")
      equipHitRow = oCG.SelectCkBoxRow(dtEquip, dgEquip, e, "Select", b)
      Select Case Me.TabControl1.SelectedTab.Name
         Case "tabAddCost"
         Case "tabTotalCost"
            ' display the total cost of ownership
            DisplayTCO()
         Case "tabEditCost"
      End Select
   End Sub


   ''' <summary>
   ''' Insert new cost item into cost table.
   ''' </summary>
   ''' <param name = "sender"></param>
   ''' <param name = "e"></param>
   Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
      Dim SQL As String
      Dim sErr As String
      Dim outil As New CUtilities()
      Dim s As String = Me.cbCostType.Text
      Dim sPID As String = outil.GetToken(s, "")
      Dim sPid2 As String = outil.GetToken(s, "")
      Try

         With Me
            SQL = "insert into equip_cost "
            SQL &= "(equip_id, cost_desc, cost_date, cost_price, "
            SQL &= "cost_type, mechanic,cost_type_desc) "
            SQL &= "values("
            SQL &= "'" & dtEquip.Rows(equipHitRow).Item("equip_id") & "', "
            SQL &= "'" & Replace(.txtCostDesc.Text, "'", "''") & "', "
            SQL &= "#" & Now.ToString & "#, "
            SQL &= UnFormat(.txtCost.Text) & ", "
            SQL &= sPID & ", "
            SQL &= "'" & Replace(.txtMechanic.Text, "'", "''") & "', "
            SQL &= "'" & Replace(sPid2, "'", "''") & "'"
            SQL &= ")"

            If oDA.SendActionSql(SQL, ConnectString, sErr) <> 1 Then
               MsgBox("Update of Equip_Cost failed." & Chr(10) & sErr & Chr(10), MsgBoxStyle.Exclamation)
               Exit Sub
            End If
            .txtCost.Text = String.Empty
            .txtCostDesc.Text = String.Empty
            .txtMechanic.Text = String.Empty
            .cbCostType.Text = String.Empty
         End With

         mbDirty = False
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

   Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
      Me.Close()
      System.Windows.Forms.Application.DoEvents()
   End Sub


#End Region

End Class
