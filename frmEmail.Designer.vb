<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEmail
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmEmail))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtFrom = New System.Windows.Forms.TextBox()
        Me.txtTo = New System.Windows.Forms.TextBox()
        Me.txtCC = New System.Windows.Forms.TextBox()
        Me.txtSubject = New System.Windows.Forms.TextBox()
        Me.txtBody = New System.Windows.Forms.TextBox()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnSend = New System.Windows.Forms.Button()
        Me.SSOleDBCombo1 = New System.Windows.Forms.ComboBox()
        Me._lblFieldLable_0 = New System.Windows.Forms.Label()
        Me.btnAttachFile = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.txtAttachments = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(43, 80)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(23, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "To:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(20, 135)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(46, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Subject:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(42, 107)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(24, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "CC:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(33, 196)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(34, 13)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Body:"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(34, 54)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(33, 13)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "From:"
        '
        'txtFrom
        '
        Me.txtFrom.Location = New System.Drawing.Point(72, 51)
        Me.txtFrom.Name = "txtFrom"
        Me.txtFrom.Size = New System.Drawing.Size(435, 20)
        Me.txtFrom.TabIndex = 5
        '
        'txtTo
        '
        Me.txtTo.Location = New System.Drawing.Point(72, 78)
        Me.txtTo.Name = "txtTo"
        Me.txtTo.Size = New System.Drawing.Size(435, 20)
        Me.txtTo.TabIndex = 6
        '
        'txtCC
        '
        Me.txtCC.Location = New System.Drawing.Point(72, 105)
        Me.txtCC.Name = "txtCC"
        Me.txtCC.Size = New System.Drawing.Size(435, 20)
        Me.txtCC.TabIndex = 7
        '
        'txtSubject
        '
        Me.txtSubject.Location = New System.Drawing.Point(72, 135)
        Me.txtSubject.Name = "txtSubject"
        Me.txtSubject.Size = New System.Drawing.Size(435, 20)
        Me.txtSubject.TabIndex = 8
        '
        'txtBody
        '
        Me.txtBody.Location = New System.Drawing.Point(73, 196)
        Me.txtBody.Multiline = True
        Me.txtBody.Name = "txtBody"
        Me.txtBody.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtBody.Size = New System.Drawing.Size(435, 249)
        Me.txtBody.TabIndex = 9
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(420, 464)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 10
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnSend
        '
        Me.btnSend.Location = New System.Drawing.Point(339, 464)
        Me.btnSend.Name = "btnSend"
        Me.btnSend.Size = New System.Drawing.Size(75, 23)
        Me.btnSend.TabIndex = 11
        Me.btnSend.Text = "&Send"
        Me.btnSend.UseVisualStyleBackColor = True
        '
        'SSOleDBCombo1
        '
        Me.SSOleDBCombo1.Location = New System.Drawing.Point(63, 11)
        Me.SSOleDBCombo1.Name = "SSOleDBCombo1"
        Me.SSOleDBCombo1.Size = New System.Drawing.Size(304, 21)
        Me.SSOleDBCombo1.TabIndex = 16
        '
        '_lblFieldLable_0
        '
        Me._lblFieldLable_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblFieldLable_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblFieldLable_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblFieldLable_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblFieldLable_0.Location = New System.Drawing.Point(3, 9)
        Me._lblFieldLable_0.Name = "_lblFieldLable_0"
        Me._lblFieldLable_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblFieldLable_0.Size = New System.Drawing.Size(53, 30)
        Me._lblFieldLable_0.TabIndex = 17
        Me._lblFieldLable_0.Text = "Select Customer"
        '
        'btnAttachFile
        '
        Me.btnAttachFile.Location = New System.Drawing.Point(420, 9)
        Me.btnAttachFile.Name = "btnAttachFile"
        Me.btnAttachFile.Size = New System.Drawing.Size(75, 23)
        Me.btnAttachFile.TabIndex = 18
        Me.btnAttachFile.Text = "&Attach File"
        Me.btnAttachFile.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'txtAttachments
        '
        Me.txtAttachments.Location = New System.Drawing.Point(72, 161)
        Me.txtAttachments.Name = "txtAttachments"
        Me.txtAttachments.Size = New System.Drawing.Size(435, 20)
        Me.txtAttachments.TabIndex = 20
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(1, 161)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(69, 13)
        Me.Label6.TabIndex = 19
        Me.Label6.Text = "Attachments:"
        '
        'frmEmail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(518, 493)
        Me.Controls.Add(Me.txtAttachments)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.btnAttachFile)
        Me.Controls.Add(Me.SSOleDBCombo1)
        Me.Controls.Add(Me._lblFieldLable_0)
        Me.Controls.Add(Me.btnSend)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.txtBody)
        Me.Controls.Add(Me.txtSubject)
        Me.Controls.Add(Me.txtCC)
        Me.Controls.Add(Me.txtTo)
        Me.Controls.Add(Me.txtFrom)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmEmail"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Email"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtFrom As System.Windows.Forms.TextBox
    Friend WithEvents txtTo As System.Windows.Forms.TextBox
    Friend WithEvents txtCC As System.Windows.Forms.TextBox
    Friend WithEvents txtSubject As System.Windows.Forms.TextBox
    Friend WithEvents txtBody As System.Windows.Forms.TextBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSend As System.Windows.Forms.Button
    Friend WithEvents SSOleDBCombo1 As System.Windows.Forms.ComboBox
    Public WithEvents _lblFieldLable_0 As System.Windows.Forms.Label
    Friend WithEvents btnAttachFile As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents txtAttachments As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
End Class
