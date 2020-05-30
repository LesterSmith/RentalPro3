Imports System.Drawing.Point
Public Class frmEmployees
    Inherits System.Windows.Forms.Form
#Region " Module Variables "
    Dim msAddorEdit As String
    Dim mbDirty As Boolean
    Dim mbFormLoading As Boolean
    Private oDA As CDataAccess
    Private iHitRow As Integer
    Private dtSI As New DataTable("dt")
    Private oCG As CGrid
#End Region

#Region " Form & Control Events "
    Private Sub frmEmployees_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        oDA = New CDataAccess()
        oCG = New CGrid()
        Me.LoadTheGrid()
        Me.iHitRow = 0
        Me.LoadRates()
    End Sub

    Private Sub dbgEmployees_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dbgEmployees.MouseUp
        Try
            Dim pt As Point = New Point(e.X, e.Y)
            Dim hti As DataGrid.HitTestInfo = Me.dbgEmployees.HitTest(pt)
            Me.dbgEmployees.Select(hti.Row)
            iHitRow = hti.Row
            Me.LoadRates()
            mbDirty = False
            msAddorEdit = "E"
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub
#End Region

#Region " Private Methods "
    Private Sub UpdateComplete()
        MsgBox("Employee change has been made.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "Update Made")
    End Sub
    Private Sub LoadRates()
        Try

            With Me.dtSI.Rows(Me.iHitRow)
                txtEmpNbr.Text = dtSI.Rows(iHitRow).Item("Emp Nbr") & ""
                Me.txtName.Text = IIf(IsDBNull(.Item("Name")), "", .Item("Name"))
                Me.txtID.Text = .Item("id")
                txtInitials.Text = .Item("Initials") & ""
                txtPassword.Text = .Item("Password") & ""
                txtPrivilege.Text = .Item("Privilege") & ""
                mbDirty = False
            End With

        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try

    End Sub
    Private Sub LoadTheGrid()
        Dim SQL As String


        Try
            oCG.InitializeDatatableForStyles(dtSI)
            oCG.BindDataTableToGrid(dtSI, Me.dbgEmployees)

            SQL = "select ID, "
            SQL = SQL & "Employee_Initials as Initials, "
            SQL = SQL & "Employee_Name as Name, Employee_Number as [Emp Nbr], Password, Privilege "
            SQL = SQL & "from Employee "
            SQL = SQL & " order by Employee_Name"
            oDA.SendQuery(SQL, dtSI, ConnectString, "dt")
            If dtSI.Rows.Count > 0 Then
                Dim Formats() As String = _
                   {"", "40", "T", "L", _
                   "", "40", "T", "L", _
                   "", "100", "T", "L", _
                   "", "60", "T", "L", _
                   "", "60", "T", "L", _
                   "", "40", "T", "L"}
                oCG.SetTablesStyle(dtSI, Me.dbgEmployees, Formats)
                oCG.BindDataTableToGrid(dtSI, Me.dbgEmployees)
                oCG.DisableAddNew(Me.dbgEmployees, Me)
            End If
            mbDirty = False
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub ClearTextBoxes()
        With Me
            .txtInitials.Text = String.Empty
            .txtName.Text = String.Empty
            .txtID.Text = String.Empty
            .txtPassword.Text = String.Empty
            .txtPrivilege.Text = String.Empty
            .txtEmpNbr.Text = String.Empty
        End With
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        If mbDirty Then
            If MsgBox("You have unsaved changes; do you want to close without saving your changes?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Exit Sub
            End If
        End If

        Me.Close()
        System.Windows.Forms.Application.DoEvents()
    End Sub

    Private Sub cmdUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdate.Click
        Dim SQL As String
        Dim sErr As String


        Try
            If String.IsNullOrEmpty(txtName.Text) OrElse String.IsNullOrEmpty(txtInitials.Text) Then
                MsgBox("You must enter at least Name and Initials.", MsgBoxStyle.Exclamation)
                Exit Sub
            End If

            With Me
                If msAddorEdit = "A" Then
                    SQL = "insert into  Employee "
                    SQL &= "(Employee_Name,Employee_Initials, Employee_Number, Password, Privilige) "
                    SQL &= "values ("
                    SQL &= "'" & Replace(.txtName.Text, "'", "''") & "', "
                    SQL &= "'" & Replace(.txtInitials.Text, "'", "''") & "', "
                    SQL &= "'" & Replace(.txtEmpNbr.Text, "'", "") & "', "
                    SQL &= "'" & .txtPassword.Text & "', "
                    SQL &= "'" & .txtPrivilege.Text & "'"
                    SQL &= ")"
                Else
                    SQL = "update Employee "
                    SQL &= "set Employee_Name = '" & Replace(.txtName.Text, "'", "''") & "', "
                    SQL &= "Employee_Initials = '" & Replace(.txtInitials.Text, "'", "") & "', "
                    SQL &= "Employee_Number = '" & Replace(txtEmpNbr.Text, "'", "") & "', "
                    SQL &= "Password = '" & txtPassword.Text & "', "
                    SQL &= "Privilege = '" & txtPrivilege.Text & "' "
                    SQL = SQL & "where id = " & .txtID.Text & " "
                End If
            End With
            If oDA.SendActionSql(SQL, ConnectString, sErr) < 1 Then
                MsgBox("Update failed: " & Chr(10) & sErr, MsgBoxStyle.Critical)
            End If
            UpdateComplete()
            msAddorEdit = "E"
            LoadTheGrid()
            iHitRow = 0
            LoadRates()
            mbDirty = False
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        If mbDirty Then
            If MsgBox("You have unsaved changes; do you want to add without saving your changes?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Exit Sub
            End If
        End If


        ClearTextBoxes()
        msAddorEdit = "A"
    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
        Dim SQL As String
        Dim dt As New DataTable()
        Dim oDA As New CDataAccess()
        Dim iRows As Integer

        Try
            Dim sErr As String = ""

            If MsgBox("Are you sure you want to delete the selected row?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Exit Sub
            End If
            SQL = "delete from Employee "
            SQL &= "where id= " & Me.dtSI.Rows(Me.iHitRow).Item("id") & " "
            iRows = oDA.SendActionSql(SQL, ConnectString, sErr)
            If iRows = 0 Then
                MsgBox("Delete of Employee failed.  " & Chr(10) & sErr, MsgBoxStyle.Critical)
                Exit Sub
            End If
            UpdateComplete()
            Me.LoadTheGrid()
            iHitRow = 0
            LoadRates()
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub
#End Region


End Class