Option Strict Off
Option Explicit On 
Friend Class frmWhoHasIt
   Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
   Public Sub New()
      MyBase.New()
      'This call is required by the Windows Form Designer.
      InitializeComponent()
      oDA = New CDataAccess()
   End Sub
   'Form overrides dispose to clean up the component list.
   Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
      If Disposing Then
         If Not components Is Nothing Then
            components.Dispose()
         End If
      End If
      MyBase.Dispose(Disposing)
   End Sub
   'Required by the Windows Form Designer
   Private components As System.ComponentModel.IContainer
   Public ToolTip1 As System.Windows.Forms.ToolTip
   Public WithEvents cmdClose As System.Windows.Forms.Button
   'NOTE: The following procedure is required by the Windows Form Designer
   'It can be modified using the Windows Form Designer.
   'Do not modify it using the code editor.
   Friend WithEvents dbgWhoHasIt As System.Windows.Forms.DataGrid
   Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
   Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
   Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
   Friend WithEvents dgRerents As System.Windows.Forms.DataGrid
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
      Me.components = New System.ComponentModel.Container()
      Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmWhoHasIt))
      Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
      Me.cmdClose = New System.Windows.Forms.Button()
      Me.dbgWhoHasIt = New System.Windows.Forms.DataGrid()
      Me.TabControl1 = New System.Windows.Forms.TabControl()
      Me.TabPage1 = New System.Windows.Forms.TabPage()
      Me.TabPage2 = New System.Windows.Forms.TabPage()
      Me.dgRerents = New System.Windows.Forms.DataGrid()
      CType(Me.dbgWhoHasIt, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.TabControl1.SuspendLayout()
      Me.TabPage1.SuspendLayout()
      Me.TabPage2.SuspendLayout()
      CType(Me.dgRerents, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.SuspendLayout()
      '
      'cmdClose
      '
      Me.cmdClose.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
      Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
      Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
      Me.cmdClose.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
      Me.cmdClose.Location = New System.Drawing.Point(765, 285)
      Me.cmdClose.Name = "cmdClose"
      Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
      Me.cmdClose.Size = New System.Drawing.Size(93, 29)
      Me.cmdClose.TabIndex = 1
      Me.cmdClose.Text = "&Close"
      '
      'dbgWhoHasIt
      '
      Me.dbgWhoHasIt.CaptionVisible = False
      Me.dbgWhoHasIt.DataMember = ""
      Me.dbgWhoHasIt.Dock = System.Windows.Forms.DockStyle.Fill
      Me.dbgWhoHasIt.HeaderForeColor = System.Drawing.SystemColors.ControlText
      Me.dbgWhoHasIt.Name = "dbgWhoHasIt"
      Me.dbgWhoHasIt.Size = New System.Drawing.Size(855, 253)
      Me.dbgWhoHasIt.TabIndex = 2
      '
      'TabControl1
      '
      Me.TabControl1.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                  Or System.Windows.Forms.AnchorStyles.Left) _
                  Or System.Windows.Forms.AnchorStyles.Right)
      Me.TabControl1.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabPage1, Me.TabPage2})
      Me.TabControl1.Name = "TabControl1"
      Me.TabControl1.SelectedIndex = 0
      Me.TabControl1.Size = New System.Drawing.Size(863, 280)
      Me.TabControl1.TabIndex = 3
      '
      'TabPage1
      '
      Me.TabPage1.Controls.AddRange(New System.Windows.Forms.Control() {Me.dbgWhoHasIt})
      Me.TabPage1.Location = New System.Drawing.Point(4, 23)
      Me.TabPage1.Name = "TabPage1"
      Me.TabPage1.Size = New System.Drawing.Size(855, 253)
      Me.TabPage1.TabIndex = 0
      Me.TabPage1.Text = "Our Equipment"
      '
      'TabPage2
      '
      Me.TabPage2.Controls.AddRange(New System.Windows.Forms.Control() {Me.dgRerents})
      Me.TabPage2.Location = New System.Drawing.Point(4, 23)
      Me.TabPage2.Name = "TabPage2"
      Me.TabPage2.Size = New System.Drawing.Size(855, 253)
      Me.TabPage2.TabIndex = 1
      Me.TabPage2.Text = "ReRents"
      '
      'dgRerents
      '
      Me.dgRerents.CaptionVisible = False
      Me.dgRerents.DataMember = ""
      Me.dgRerents.Dock = System.Windows.Forms.DockStyle.Fill
      Me.dgRerents.HeaderForeColor = System.Drawing.SystemColors.ControlText
      Me.dgRerents.Name = "dgRerents"
      Me.dgRerents.Size = New System.Drawing.Size(855, 253)
      Me.dgRerents.TabIndex = 0
      '
      'frmWhoHasIt
      '
      Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
      Me.ClientSize = New System.Drawing.Size(864, 326)
      Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabControl1, Me.cmdClose})
      Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
      Me.Location = New System.Drawing.Point(136, 143)
      Me.MinimizeBox = False
      Me.Name = "frmWhoHasIt"
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "Who Has The Equipment"
      CType(Me.dbgWhoHasIt, System.ComponentModel.ISupportInitialize).EndInit()
      Me.TabControl1.ResumeLayout(False)
      Me.TabPage1.ResumeLayout(False)
      Me.TabPage2.ResumeLayout(False)
      CType(Me.dgRerents, System.ComponentModel.ISupportInitialize).EndInit()
      Me.ResumeLayout(False)

   End Sub
#End Region
#Region " Module Variables "
   Dim oDA As CDataAccess
   Dim oCG As New CGrid()


#End Region


#Region " Form & Control Events "
   Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
      Me.Close()
      System.Windows.Forms.Application.DoEvents()
   End Sub


   Private Sub frmWhoHasIt_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
      LoadTheGrid()
      LoadRerentGrid()
   End Sub


#End Region


#Region " Private Methods "
   Private Sub LoadTheGrid()
      Dim SQL As String
      Dim dt As New DataTable("dt")


      Try
         SQL = ""
         SQL = SQL & "SELECT DISTINCTROW Equipment.Equip_ID, "
         SQL = SQL & "Equipment.Equip_Name, Equipment.Available as Availability, "
         SQL = SQL & "Equipment.Rented_Date, Equipment.Available_Date as [Due Back], "
         SQL = SQL & "Customers.CompanyName, Customers.ContactName, "
         SQL = SQL & "Customers.PhoneNumber "
         SQL = SQL & "FROM Customers RIGHT JOIN Equipment ON "
         SQL = SQL & "Customers.CustomerID = Equipment.Renting_Company_ID "
         SQL = SQL & "where equipment.available<>'YES' "
         SQL &= "order by equipment.equip_name"

         oDA.SendQuery(SQL, dt, ConnectString, "dt")

         If dt.Rows.Count > 0 Then
            Dim Formats() As String = _
               {"", "60", "T", "L", _
               "", "150", "T", "L", _
               "", "60", "T", "L", _
               "MM/dd/yyyy hh:mm tt", "130", "T", "L", _
               "MM/dd/yyyy hh:mm tt", "130", "T", "L", _
               "", "100", "T", "L", _
               "", "100", "T", "L", _
               "", "100", "T", "L"}
            oCG.SetTablesStyle(dt, Me.dbgWhoHasIt, Formats)
            Me.dbgWhoHasIt.SetDataBinding(dt, "")
            oCG.DisableAddNew(dbgWhoHasIt, Me)
         End If
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

   Private Sub LoadRerentGrid()
      Dim sql As String
      Dim dt As New DataTable("dt")


      Try
         sql = ""
         sql &= "select Equip_Name,Customer_Name,Rented_Date, "
         sql &= "Nbr_Periods as Nbr_Per,Period,PO_Number,Customer_Price "
         sql &= "from rerents "
         sql &= "where not isnull(rented_date) "
         sql &= "and isnull(returned_date) "
         If oDA.SendQuery(sql, dt, ConnectString, "dt") > 0 Then
            Dim formats() As String = _
               {"", "150", "T", "L", _
                "", "150", "T", "L", _
                "MM/dd/yy hh:mm tt", "130", "T", "L", _
                "", "60", "T", "R", _
                "", "60", "T", "R", _
                "", "60", "T", "L", _
                "$#,##0", "60", "T", "R"}
            oCG.SetTablesStyle(dt, Me.dgRerents, formats)
            Me.dgRerents.SetDataBinding(dt, "")
            oCG.DisableAddNew(Me.dgRerents, Me)
         End If
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

#End Region




End Class