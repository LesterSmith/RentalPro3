Imports System.IO

Public Class frmEmail
    Private oDA As CDataAccess
    Private _emailServer As String
    Private _email As HHISoftware.EmailWrapperLib

    Public Sub New()
        MyBase.New()
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        oDA = New CDataAccess()
        _emailServer = modMain.EmailServer
    End Sub
    Private Sub frmEmail_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            txtBody.Text = modMain.EmailBody
            txtFrom.Text = modMain.EMail
            txtSubject.Text = modMain.EmailSubject
            If Not Directory.Exists(CutePDFFilePath) Then
                Directory.CreateDirectory("C:\CutePDFFiles")
                CutePDFFilePath = "C:\CutePDFFiles"
            End If

            LoadCustomerCombo()

            If Not String.IsNullOrEmpty(modMain.CustomerEmail) Then
                SSOleDBCombo1.Text = modMain.CustomerEmail
            End If
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try

    End Sub

    Private Sub GetCustomerMail(ByVal customerName As String)

        Dim sql As String = String.Empty

        sql = "select EmailAddress from Customers "
        sql = sql & "where companyname = '" & Replace(SSOleDBCombo1.Text, "'", "''") & "'"

        Dim dt As DataTable
        dt = New DataTable
        If oDA.SendQuery(sql, dt, ConnectString) > 0 Then
            With dt.Rows(0)
                txtTo.Text = .Item("EmailAddress")
            End With
        End If
    End Sub

    Private Sub LoadCustomerCombo()
        Dim dt As New DataTable()
        Dim i As Integer
        Dim sql As String = "select companyname from customers "
        sql &= "order by companyname"
        oDA.SendQuery(sql, dt, ConnectString)
        Me.SSOleDBCombo1.Items.Clear()

        For i = 0 To dt.Rows.Count - 1
            With dt.Rows(i)
                Me.SSOleDBCombo1.Items.Add(.Item("companyname"))
            End With
        Next
    End Sub

    Private Sub SSOleDBCombo1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SSOleDBCombo1.SelectedIndexChanged
        Try
            GetCustomerMail(SSOleDBCombo1.Text)
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

    Private Sub btnAttachFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAttachFile.Click
        OpenFileDialog1.FileName = "*.pdf"
        OpenFileDialog1.InitialDirectory = CutePDFFilePath
        OpenFileDialog1.ShowDialog()
        If Not File.Exists(OpenFileDialog1.FileName) Then
            MsgBox("You did not select a file to attach.")
            Exit Sub
        End If
        If Not String.IsNullOrEmpty(txtAttachments.Text) Then
            txtAttachments.Text += "; " + Path.GetFileName(OpenFileDialog1.FileName)
        Else
            txtAttachments.Text += Path.GetFileName(OpenFileDialog1.FileName)
        End If
    End Sub

    Private Sub btnSend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSend.Click
        Try
            _email = New HHISoftware.EmailWrapperLib(_emailServer, txtFrom.Text, "")
            _email.SendEmail(txtBody.Text, txtSubject.Text, txtTo.Text, attach:=txtAttachments.Text)
        Catch ex As Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub
End Class