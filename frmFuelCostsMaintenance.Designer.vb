<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFuelCostsMaintenance
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmFuelCostsMaintenance))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.txtID = New System.Windows.Forms.TextBox()
        Me.lblID = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.lblProductDesc = New System.Windows.Forms.Label()
        Me.txtPricePerUnit = New System.Windows.Forms.TextBox()
        Me.txtProductDescription = New System.Windows.Forms.TextBox()
        Me.dbgFuelCosts = New System.Windows.Forms.DataGrid()
        Me.cmdAdd = New System.Windows.Forms.Button()
        Me.cmdUpdate = New System.Windows.Forms.Button()
        Me.cmdDelete = New System.Windows.Forms.Button()
        Me.cmdClose = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        CType(Me.dbgFuelCosts, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtID)
        Me.GroupBox1.Controls.Add(Me.lblID)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.lblProductDesc)
        Me.GroupBox1.Controls.Add(Me.txtPricePerUnit)
        Me.GroupBox1.Controls.Add(Me.txtProductDescription)
        Me.GroupBox1.Location = New System.Drawing.Point(4, 142)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(276, 96)
        Me.GroupBox1.TabIndex = 19
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Selected Items"
        '
        'txtID
        '
        Me.txtID.Location = New System.Drawing.Point(121, 69)
        Me.txtID.Name = "txtID"
        Me.txtID.ReadOnly = True
        Me.txtID.Size = New System.Drawing.Size(64, 20)
        Me.txtID.TabIndex = 9
        Me.txtID.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblID
        '
        Me.lblID.AutoSize = True
        Me.lblID.Location = New System.Drawing.Point(13, 72)
        Me.lblID.Name = "lblID"
        Me.lblID.Size = New System.Drawing.Size(18, 13)
        Me.lblID.TabIndex = 8
        Me.lblID.Text = "ID"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(13, 46)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(83, 13)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Price Per Gallon"
        '
        'lblProductDesc
        '
        Me.lblProductDesc.AutoSize = True
        Me.lblProductDesc.Location = New System.Drawing.Point(13, 22)
        Me.lblProductDesc.Name = "lblProductDesc"
        Me.lblProductDesc.Size = New System.Drawing.Size(54, 13)
        Me.lblProductDesc.TabIndex = 6
        Me.lblProductDesc.Text = "Fuel Type"
        '
        'txtPricePerUnit
        '
        Me.txtPricePerUnit.Location = New System.Drawing.Point(122, 44)
        Me.txtPricePerUnit.Name = "txtPricePerUnit"
        Me.txtPricePerUnit.Size = New System.Drawing.Size(100, 20)
        Me.txtPricePerUnit.TabIndex = 3
        Me.txtPricePerUnit.Tag = "(No Auto Formatting)"
        Me.txtPricePerUnit.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtProductDescription
        '
        Me.txtProductDescription.Location = New System.Drawing.Point(122, 19)
        Me.txtProductDescription.MaxLength = 50
        Me.txtProductDescription.Name = "txtProductDescription"
        Me.txtProductDescription.Size = New System.Drawing.Size(135, 20)
        Me.txtProductDescription.TabIndex = 2
        Me.txtProductDescription.Tag = "(No Auto Formatting)"
        '
        'dbgFuelCosts
        '
        Me.dbgFuelCosts.AllowSorting = False
        Me.dbgFuelCosts.DataMember = ""
        Me.dbgFuelCosts.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dbgFuelCosts.Location = New System.Drawing.Point(1, -1)
        Me.dbgFuelCosts.Name = "dbgFuelCosts"
        Me.dbgFuelCosts.Size = New System.Drawing.Size(284, 137)
        Me.dbgFuelCosts.TabIndex = 20
        '
        'cmdAdd
        '
        Me.cmdAdd.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAdd.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAdd.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAdd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAdd.Location = New System.Drawing.Point(112, 244)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAdd.Size = New System.Drawing.Size(77, 26)
        Me.cmdAdd.TabIndex = 24
        Me.cmdAdd.Text = "&Add"
        Me.cmdAdd.UseVisualStyleBackColor = False
        '
        'cmdUpdate
        '
        Me.cmdUpdate.BackColor = System.Drawing.SystemColors.Control
        Me.cmdUpdate.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdUpdate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdUpdate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdUpdate.Location = New System.Drawing.Point(195, 244)
        Me.cmdUpdate.Name = "cmdUpdate"
        Me.cmdUpdate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdUpdate.Size = New System.Drawing.Size(77, 26)
        Me.cmdUpdate.TabIndex = 23
        Me.cmdUpdate.Text = "&Save"
        Me.cmdUpdate.UseVisualStyleBackColor = False
        '
        'cmdDelete
        '
        Me.cmdDelete.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDelete.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdDelete.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDelete.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDelete.Location = New System.Drawing.Point(112, 276)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDelete.Size = New System.Drawing.Size(77, 26)
        Me.cmdDelete.TabIndex = 25
        Me.cmdDelete.Text = "&Delete"
        Me.cmdDelete.UseVisualStyleBackColor = False
        '
        'cmdClose
        '
        Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
        Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdClose.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdClose.Location = New System.Drawing.Point(195, 276)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClose.Size = New System.Drawing.Size(77, 26)
        Me.cmdClose.TabIndex = 26
        Me.cmdClose.Text = "&Close"
        Me.cmdClose.UseVisualStyleBackColor = False
        '
        'frmFuelCostsMaintenance
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 316)
        Me.Controls.Add(Me.cmdAdd)
        Me.Controls.Add(Me.cmdUpdate)
        Me.Controls.Add(Me.cmdDelete)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.dbgFuelCosts)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmFuelCostsMaintenance"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Fuel Costs Maintenance"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.dbgFuelCosts, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtID As System.Windows.Forms.TextBox
    Friend WithEvents lblID As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lblProductDesc As System.Windows.Forms.Label
    Friend WithEvents txtPricePerUnit As System.Windows.Forms.TextBox
    Friend WithEvents txtProductDescription As System.Windows.Forms.TextBox
    Friend WithEvents dbgFuelCosts As System.Windows.Forms.DataGrid
    Public WithEvents cmdAdd As System.Windows.Forms.Button
    Public WithEvents cmdUpdate As System.Windows.Forms.Button
    Public WithEvents cmdDelete As System.Windows.Forms.Button
    Public WithEvents cmdClose As System.Windows.Forms.Button
End Class
