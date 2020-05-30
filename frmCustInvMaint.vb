Imports System.Data.OleDb
Public Class frmCustInvMaint
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
   Friend WithEvents dbgCust As System.Windows.Forms.DataGrid
   Friend WithEvents dbgInv As System.Windows.Forms.DataGrid
   Friend WithEvents dbgDet As System.Windows.Forms.DataGrid
   Friend WithEvents btnSaveChangesToCust As System.Windows.Forms.Button
   Friend WithEvents btnSaveChangesToInvHdr As System.Windows.Forms.Button
   Friend WithEvents btnSaveChangesToInvDetails As System.Windows.Forms.Button
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
      Me.dbgCust = New System.Windows.Forms.DataGrid()
      Me.dbgInv = New System.Windows.Forms.DataGrid()
      Me.dbgDet = New System.Windows.Forms.DataGrid()
      Me.btnSaveChangesToCust = New System.Windows.Forms.Button()
      Me.btnSaveChangesToInvHdr = New System.Windows.Forms.Button()
      Me.btnSaveChangesToInvDetails = New System.Windows.Forms.Button()
      CType(Me.dbgCust, System.ComponentModel.ISupportInitialize).BeginInit()
      CType(Me.dbgInv, System.ComponentModel.ISupportInitialize).BeginInit()
      CType(Me.dbgDet, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.SuspendLayout()
      '
      'dbgCust
      '
      Me.dbgCust.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                  Or System.Windows.Forms.AnchorStyles.Right)
      Me.dbgCust.CaptionText = "Select Customer"
      Me.dbgCust.DataMember = ""
      Me.dbgCust.HeaderForeColor = System.Drawing.SystemColors.ControlText
      Me.dbgCust.Location = New System.Drawing.Point(6, 4)
      Me.dbgCust.Name = "dbgCust"
      Me.dbgCust.Size = New System.Drawing.Size(692, 112)
      Me.dbgCust.TabIndex = 1
      '
      'dbgInv
      '
      Me.dbgInv.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                  Or System.Windows.Forms.AnchorStyles.Right)
      Me.dbgInv.CaptionText = "Select Invoice Header"
      Me.dbgInv.DataMember = ""
      Me.dbgInv.HeaderForeColor = System.Drawing.SystemColors.ControlText
      Me.dbgInv.Location = New System.Drawing.Point(7, 152)
      Me.dbgInv.Name = "dbgInv"
      Me.dbgInv.Size = New System.Drawing.Size(690, 133)
      Me.dbgInv.TabIndex = 2
      '
      'dbgDet
      '
      Me.dbgDet.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                  Or System.Windows.Forms.AnchorStyles.Right)
      Me.dbgDet.CaptionText = "Invoice Details"
      Me.dbgDet.DataMember = ""
      Me.dbgDet.HeaderForeColor = System.Drawing.SystemColors.ControlText
      Me.dbgDet.Location = New System.Drawing.Point(6, 315)
      Me.dbgDet.Name = "dbgDet"
      Me.dbgDet.Size = New System.Drawing.Size(693, 200)
      Me.dbgDet.TabIndex = 3
      '
      'btnSaveChangesToCust
      '
      Me.btnSaveChangesToCust.Location = New System.Drawing.Point(590, 122)
      Me.btnSaveChangesToCust.Name = "btnSaveChangesToCust"
      Me.btnSaveChangesToCust.Size = New System.Drawing.Size(96, 24)
      Me.btnSaveChangesToCust.TabIndex = 4
      Me.btnSaveChangesToCust.Text = "&Save Changes"
      '
      'btnSaveChangesToInvHdr
      '
      Me.btnSaveChangesToInvHdr.Location = New System.Drawing.Point(592, 288)
      Me.btnSaveChangesToInvHdr.Name = "btnSaveChangesToInvHdr"
      Me.btnSaveChangesToInvHdr.Size = New System.Drawing.Size(96, 24)
      Me.btnSaveChangesToInvHdr.TabIndex = 5
      Me.btnSaveChangesToInvHdr.Text = "&Save Changes"
      '
      'btnSaveChangesToInvDetails
      '
      Me.btnSaveChangesToInvDetails.Location = New System.Drawing.Point(592, 524)
      Me.btnSaveChangesToInvDetails.Name = "btnSaveChangesToInvDetails"
      Me.btnSaveChangesToInvDetails.Size = New System.Drawing.Size(96, 24)
      Me.btnSaveChangesToInvDetails.TabIndex = 6
      Me.btnSaveChangesToInvDetails.Text = "&Save Changes"
      '
      'frmCustInvMaint
      '
      Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
      Me.ClientSize = New System.Drawing.Size(704, 558)
      Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnSaveChangesToInvDetails, Me.btnSaveChangesToInvHdr, Me.btnSaveChangesToCust, Me.dbgDet, Me.dbgInv, Me.dbgCust})
      Me.Name = "frmCustInvMaint"
      Me.Text = "frmCustInvMaint"
      CType(Me.dbgCust, System.ComponentModel.ISupportInitialize).EndInit()
      CType(Me.dbgInv, System.ComponentModel.ISupportInitialize).EndInit()
      CType(Me.dbgDet, System.ComponentModel.ISupportInitialize).EndInit()
      Me.ResumeLayout(False)

   End Sub

#End Region

#Region " Module Variables "
   Private oDA As New CDataAccess()
   Private dtCust As New DataTable()
   Private dtInv As New DataTable()
   Private dtDet As New DataTable()
   Private SQL As String
   Private oCG As New CGrid()
   Private miInvHitRow As Integer
   Private miCustHitRow As Integer
   Private dsCust As New DataSet()
   Private dsInv As New DataSet()
   Private dsDet As New DataSet()
   Private daCust As OleDbDataAdapter
   Private daInv As OleDbDataAdapter
   Private daDet As OleDbDataAdapter
   Private cn As OleDbConnection
   Private sqlCust As String
   Private sqlInv As String
   Private sqlDet As String
   Private scbCust As OleDbCommandBuilder
   Private scbInv As OleDbCommandBuilder
   Private scbDet As OleDbCommandBuilder
   Private cmdCust As OleDbCommand
   Private cmdInv As OleDbCommand
   Private cmdDet As OleDbCommand


#End Region

#Region " Private Methods "
   Private Sub LoadInvoiceDetails(ByVal InvID As Integer)

      Try
         sqlDet = "select * from invoice_details where invoiceid = " & InvID & " "
         sqlDet &= "order by record_type"
         cmdDet = New OleDbCommand(sqlDet, cn)
         daDet = New OleDbDataAdapter(cmdDet)
         scbDet = New OleDbCommandBuilder(daDet)
         dsDet = New DataSet()
         daDet.Fill(dsDet, "Det")
         Me.dbgDet.DataSource = dsDet.Tables(0)
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

   Private Sub LoadInvoiceHeaders(ByVal CustID As Integer)
      ' load the invoices into header grid

      Try
         sqlInv = "select * from invoices where customerid = " & CustID & " "
         sqlInv &= "order by invoiceid"
         cmdInv = New OleDbCommand(sqlInv, cn)
         daInv = New OleDbDataAdapter(cmdInv)
         scbInv = New OleDbCommandBuilder(daInv)
         dsInv = New DataSet()
         daInv.Fill(dsInv, "Inv")
         Me.dbgInv.DataSource = dsInv.Tables(0)
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

#End Region

#Region " Form & Control Events "
   Private Sub frmCustInvMaint_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
      'Load customer grid

      Try
         sqlCust = "select * from customers order by companyname"
         cn = New OleDbConnection(ConnectString)
         cmdCust = New OleDbCommand(sqlCust, cn)
         daCust = New OleDbDataAdapter(cmdCust)
         scbCust = New OleDbCommandBuilder(daCust)
         daCust.Fill(dsCust, "Cust")
         Me.dbgCust.DataSource = dsCust.Tables(0)
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

   Private Sub dbgCust_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dbgCust.MouseUp

      Try
         Static Noise As Boolean
         Dim custID As Integer
         Try
            Dim b As Boolean
            If Noise Then Exit Sub
            Noise = True
            miCustHitRow = dbgCust.CurrentRowIndex
            custID = dsCust.Tables(0).Rows(miCustHitRow).Item("customerid")
            LoadInvoiceHeaders(custID)

            Noise = False
         Catch
         End Try
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

   Private Sub dbgInv_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dbgInv.MouseUp

      Try
         Static Noise As Boolean
         Dim invID As Integer

         Try
            Dim b As Boolean
            If Noise Then Exit Sub
            Noise = True
            miInvHitRow = dbgInv.CurrentRowIndex
            invID = dsInv.Tables(0).Rows(miInvHitRow).Item("invoiceid")
            LoadInvoiceDetails(invID)
            Noise = False
         Catch
         End Try
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub


   Private Sub dbgCust_Navigate(ByVal sender As Object, ByVal ne As System.Windows.Forms.NavigateEventArgs) Handles dbgCust.Navigate
      'Dim dt As New DataTable()
      'dt = dsCust.Tables(0).GetChanges
      'Debug.Assert(dt.Rows.Count > 0)

      Try
         daCust.Update(dsCust.Tables(0))
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub


   Private Sub dbgCust_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dbgCust.CurrentCellChanged
      Dim dt As New DataTable()

      Try
         'dt = dsCust.Tables(0).GetChanges
         'Debug.Assert(dt.Rows.Count > 0)
         daCust.Update(dsCust.Tables(0))
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub


   Private Sub dbgInv_Navigate(ByVal sender As Object, ByVal ne As System.Windows.Forms.NavigateEventArgs) Handles dbgInv.Navigate

      Try
         daInv.Update(dsInv.Tables(0))
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub


   Private Sub dbgInv_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dbgInv.CurrentCellChanged
      'Dim dt As New DataTable()
      'dt = dsCust.Tables(0).GetChanges
      'Debug.Assert(dt.Rows.Count > 0)

      Try
         daInv.Update(dsInv.Tables(0))
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub


   Private Sub dbgDet_Navigate(ByVal sender As Object, ByVal ne As System.Windows.Forms.NavigateEventArgs) Handles dbgDet.Navigate

      Try
         daDet.Update(dsDet.Tables(0))
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub


   Private Sub dbgDet_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dbgDet.CurrentCellChanged
      'Dim dt As New DataTable()
      'dt = dsCust.Tables(0).GetChanges
      'Debug.Assert(dt.Rows.Count > 0)

      Try
         daDet.Update(dsDet.Tables(0))
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub



#End Region

End Class
