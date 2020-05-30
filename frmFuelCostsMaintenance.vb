Public Class frmFuelCostsMaintenance
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

#Region " ..ctor "

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        oDA = New CDataAccess()
        oCG = New CGrid()
    End Sub
#End Region

#Region " Form & Control Events "
    Private Sub dbgSALEitems_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dbgFuelCosts.MouseUp
        Try
            Dim pt = New Point(e.X, e.Y)
            Dim hti As DataGrid.HitTestInfo = Me.dbgFuelCosts.HitTest(pt)
            Me.dbgFuelCosts.Select(hti.Row)
            iHitRow = hti.Row
            Me.LoadRates()
            mbDirty = False
            msAddorEdit = "E"
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

    Private Sub cmdClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        If mbDirty Then
            If MsgBox("You have unsaved changes; do you want to close without saving your changes?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Exit Sub
            End If
        End If

        Me.Close()
        System.Windows.Forms.Application.DoEvents()
    End Sub
    Private Sub txtPricePerUnit_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPricePerUnit.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtPricePerUnit_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPricePerUnit.KeyDown
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
        If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
        If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
    End Sub
    Private Sub txtPricePerUnit_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPricePerUnit.Enter
        txtPricePerUnit.SelectionStart = 0
        txtPricePerUnit.SelectionLength = txtPricePerUnit.Text.Trim.Length
    End Sub
    Private Sub txtProductDescription_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtProductDescription.KeyPress
        If e.KeyChar = Chr(13) Then
            e.Handled = True
            Exit Sub
        End If
    End Sub
    Private Sub txtProductDescription_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtProductDescription.KeyDown
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
        If e.KeyCode = Keys.Up Then SendKeys.Send("+{TAB}")
        If e.KeyCode = Keys.Down Then SendKeys.Send("{TAB}")
    End Sub
    Private Sub txtProductDescription_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProductDescription.Enter
        txtProductDescription.SelectionStart = 0
        txtProductDescription.SelectionLength = txtProductDescription.Text.Trim.Length
    End Sub

    Private Sub frmLaborItemsMaint_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.LoadTheGrid()
        Me.iHitRow = 0
        Me.LoadRates()
    End Sub


    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Dim SQL As String
        Dim dt As New DataTable()

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
            SQL = "delete from fuel_price "
            SQL &= "where unique_id= " & Me.dtSI.Rows(Me.iHitRow).Item("unique_id") & " "
            iRows = oDA.SendActionSql(SQL, ConnectString, sErr)
            If iRows = 0 Then
                MsgBox("Delete of labor item failed.  " & Chr(10) & sErr, MsgBoxStyle.Critical)
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


    Private Sub cmdUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdate.Click
        Dim SQL As String
        Dim sErr As String


        Try
            With Me
                If msAddorEdit = "A" Then
                    SQL = "insert into  fuel_price"
                    SQL &= "(fuel_type,price) "
                    SQL &= "values("
                    SQL &= "'" & Replace(.txtProductDescription.Text, "'", "''") & "', "
                    SQL &= UnFormat(.txtPricePerUnit.Text) & " "
                    SQL &= ")"
                Else
                    SQL = "update fuel_price "
                    SQL &= "set fuel_type = '" & Replace(.txtProductDescription.Text, "'", "''") & "', "
                    SQL &= "price = " & UnFormat(.txtPricePerUnit.Text) & " "
                    SQL = SQL & "where unique_id = " & .txtID.Text & " "
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

#End Region


#Region " Private Methods "
    Private Sub UpdateComplete()
        MsgBox("Fuel Price change has been made.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "Update Made")
    End Sub
    Private Sub LoadRates()
        Try

            With Me.dtSI.Rows(Me.iHitRow)
                If Not IsDBNull(.Item("Price")) Then
                    Me.txtPricePerUnit.Text = FormatCurrency(.Item("Price"))
                Else
                    Me.txtPricePerUnit.Text = "$0.00"
                End If

                Me.txtProductDescription.Text = IIf(IsDBNull(.Item("fuel_type")), "", .Item("fuel_type"))
                Me.txtID.Text = .Item("unique_id")
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
            oCG.BindDataTableToGrid(dtSI, Me.dbgFuelCosts)

            SQL = "select Fuel_Type, "
            SQL = SQL & "Price, "
            SQL = SQL & "Unique_Id "
            SQL = SQL & "from Fuel_Price "
            SQL = SQL & " order by Fuel_type"
            oDA.SendQuery(SQL, dtSI, ConnectString, "dt")
            If dtSI.Rows.Count > 0 Then
                Dim Formats() As String = _
                   {"", "140", "T", "L", _
                   "$#,##0.00", "60", "T", "R", _
                   "", "60", "T", "R"}
                oCG.SetTablesStyle(dtSI, Me.dbgFuelCosts, Formats)
                oCG.BindDataTableToGrid(dtSI, Me.dbgFuelCosts)
                oCG.DisableAddNew(Me.dbgFuelCosts, Me)
            End If
            mbDirty = False
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Private Sub ClearTextBoxes()
        With Me
            .txtPricePerUnit.Text = ""
            .txtProductDescription.Text = ""
            .txtID.Text = ""
        End With
    End Sub

#End Region


End Class