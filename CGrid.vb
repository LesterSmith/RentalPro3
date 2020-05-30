'* Purpose: Automatic formatting of datagrid columns
'* including optional check box columns.
'* Author:  Les Smith
'* Date Created: 05/15/2003 at 09:09:57
'* CopyRight:  HHI Software
'*
Imports System.Windows.Forms
Imports System.Drawing
Public Class CGrid
#Region " Public Methods "
    ''' <summary>
    ''' Loops thru the captions of the passed grid and
    ''' removes "_" replacing with " " and changing the
    ''' Names to Proper Names
    ''' </summary>
    Public Function FixGridColumnCaption(ByVal _
       GridColumnCaption As String) As String
        Return ProperName(Replace(GridColumnCaption, "_", " ")) & "  "
    End Function

    ''' <summary>
    ''' Uncheck all checkboxes in passed datatable
    ''' </summary>
    ''' <param name = "dt"></param>
    ''' <param name = "CBName"></param>
    Public Sub UncheckAllBoxes(ByRef dt As DataTable, ByVal CBName As String)
        Dim i As Integer
        For i = 0 To dt.Rows.Count - 1
            dt.Rows(i).Item(CBName) = "false"
        Next
    End Sub
    ''' <summary>
    ''' Check all checkboxes in passed datatable
    ''' </summary>
    ''' <param name = "dt"></param>
    ''' <param name = "CBName"></param>
    Public Sub CheckAllBoxes(ByRef dt As DataTable, ByVal CBName As String)
        Dim i As Integer
        For i = 0 To dt.Rows.Count - 1
            dt.Rows(i).Item(CBName) = "true"
        Next
    End Sub


    Public Sub BindDataTableToGrid(ByRef dt As DataTable, _
       ByRef dbg As DataGrid)
        dbg.SetDataBinding(dt, "")
    End Sub

    Public Sub InitializeDatatableForStyles(ByRef dt As DataTable)
        ' must name the datatable to set table styles
        dt = New DataTable("dt")
    End Sub

    Public Sub ClearDataTableForRebinding(ByRef dt As DataTable)
        dt = New DataTable("dt")
    End Sub

    Public Sub DisableAddNew(ByRef dg As DataGrid, ByRef Frm As Form)
        ' Disable addnew capability on the grid.
        ' Note that AllowEdit and AllowDelete can be disabled
        ' by adding or changing the "AllowNew" property to 
        ' AllowDelete or AllowEdit.
        Dim cm As CurrencyManager = _
           CType(Frm.BindingContext(dg.DataSource, dg.DataMember),  _
           CurrencyManager)
        CType(cm.List, DataView).AllowNew = False
    End Sub

    Public Overloads Sub SetTablesStyle(ByRef dt As DataTable, _
       ByRef dg As DataGrid)
        ' All columns are unformatted text boxes, with no check box
        ' All columns will be readonly
        Dim i As Integer
        Dim ts As New DataGridTableStyle()
        Dim myDataCol As New DataGridBoolColumn() ' checkbox column

        Try

            dg.TableStyles.Clear()
            ts.GridColumnStyles.Clear()

            ' map the table style to the dt
            ts.MappingName = dt.TableName ' "dt"

            ' add the column style for the ckbox col
            ts.GridColumnStyles.Add(myDataCol)

            ' for the rest of the rows, make text box columns
            For i = 0 To dt.Columns.Count - 1
                Dim dgtbc As New DataGridTextBoxColumn()
                dgtbc.HeaderText = dt.Columns.Item(i).ColumnName
                dgtbc.MappingName = dt.Columns.Item(i).ColumnName
                dgtbc.ReadOnly = True
                dgtbc.NullText = String.Empty
                ts.GridColumnStyles.Add(dgtbc)
                ts.AllowSorting = False
            Next

            ' add the table style to the grid
            dg.TableStyles.Add(ts)

        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

    Public Overloads Sub SetTablesStyle(ByVal AddCkBox As String, _
       ByRef dt As DataTable, _
       ByRef dg As DataGrid)
        ' All columns are unformatted text boxes, with a check box
        ' All columns will be readonly
        Dim i As Integer
        Dim ts As New DataGridTableStyle()
        Dim myDataCol As New DataGridBoolColumn() ' checkbox column

        Try
            If AddCkBox.Length > 0 Then
                AddCheckBoxColumn(dt, AddCkBox)
            End If

            dg.TableStyles.Clear()
            ts.GridColumnStyles.Clear()

            ' map the table style to the dt
            ts.MappingName = dt.TableName '"dt"

            ' set the header and mapping for the ckbox column
            If AddCkBox.Length > 0 Then
                With myDataCol
                    .HeaderText = dt.Columns.Item(AddCkBox).ColumnName
                    .MappingName = dt.Columns.Item(AddCkBox).ColumnName
                    .FalseValue = "false"
                    .TrueValue = "true"
                    .AllowNull = False
                    .ReadOnly = True
                End With
                ' add the column style for the ckbox col
                ts.GridColumnStyles.Add(myDataCol)
            End If

            ' for the rest of the rows, make text box columns
            For i = 0 To dt.Columns.Count - 1
                If Not dt.Columns(i).ColumnName = AddCkBox Then
                    Dim dgtbc As New DataGridTextBoxColumn()
                    dgtbc.HeaderText = dt.Columns.Item(i).ColumnName
                    dgtbc.MappingName = dt.Columns.Item(i).ColumnName
                    dgtbc.ReadOnly = True
                    dgtbc.NullText = String.Empty
                    ts.GridColumnStyles.Add(dgtbc)
                    ts.AllowSorting = False
                End If
            Next

            ' add the table style to the grid
            dg.TableStyles.Add(ts)

        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub


    Public Overloads Sub SetTablesStyle(ByVal AddCkBox As String, _
       ByRef dt As DataTable, _
       ByRef dg As DataGrid, _
       ByVal ParamArray Formats() As String)
        ' This method allows a checkbox and optional formatting
        ' param array.  If AddCkBox.Length >0 it will add 
        ' a new first column as a check box
        ' if paramarray is supplied it is an array of 
        ' paired items of "format,width,ReadOnly,format,width,..."
        ' where format="$#,##0.00"|"MM/dd/yyyy"|etc
        '       width = "90", readonly =T|F
        ' with every column having at least a pair of placeholders,
        ' including optional checkbox column
        Dim i As Integer
        Dim ts As New DataGridTableStyle()
        Dim myDataCol As New DataGridBoolColumn() ' checkbox column

        Try
            If AddCkBox.Length > 0 Then
                AddCheckBoxColumn(dt, AddCkBox)
            End If

            dg.TableStyles.Clear()
            ts.GridColumnStyles.Clear()

            ' map the table style to the dt
            ts.MappingName = dt.TableName '"dt"

            ' set the header and mapping for the ckbox column
            If AddCkBox.Length > 0 Then
                With myDataCol
                    .HeaderText = dt.Columns.Item(AddCkBox).ColumnName
                    .MappingName = dt.Columns.Item(AddCkBox).ColumnName
                    .FalseValue = "false"
                    .TrueValue = "true"
                    .AllowNull = False
                    .ReadOnly = True
                End With
                ' add the column style for the ckbox col
                ts.GridColumnStyles.Add(myDataCol)
            End If


            ' for the rest of the rows, make text box columns
            Dim j As Integer '= IIf(AddCkBox.Length > 0, 1, 0)

            For i = 0 To dt.Columns.Count - 1
                If Not dt.Columns(i).ColumnName = AddCkBox Then
                    Dim dgtbc As New DataGridTextBoxColumn()
                    dgtbc.HeaderText = _
                       FixGridColumnCaption(dt.Columns.Item(i).ColumnName)
                    dgtbc.MappingName = dt.Columns.Item(i).ColumnName
                    If UBound(Formats) > 0 Then
                        Try
                            j = i * 4
                            dgtbc.Format = Formats(j)
                            dgtbc.Width = Val(Formats(j + 1))
                            dgtbc.ReadOnly = IIf(Formats(j + 2) = "T", True, False)
                            If Formats(j + 3) = "C" Then
                                dgtbc.Alignment = HorizontalAlignment.Center
                            ElseIf Formats(j + 3) = "L" Then
                                dgtbc.Alignment = HorizontalAlignment.Left
                            ElseIf Formats(j + 3) = "R" Then
                                dgtbc.Alignment = HorizontalAlignment.Right
                            End If
                        Catch
                        End Try
                        dgtbc.NullText = String.Empty
                        ts.GridColumnStyles.Add(dgtbc)
                        ts.AllowSorting = False
                    End If
                End If
            Next

            ' add the table style to the grid
            dg.TableStyles.Add(ts)

        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

    Public Overloads Sub SetTablesStyle(ByRef dt As DataTable, _
       ByRef dg As DataGrid, _
       ByVal ParamArray Formats() As String)
        ' This method is used to pass a param array
        ' if paramarray is supplied it is an array of 
        ' sets of 3 items of "format,width,ReadOnly,format,width,..."
        ' where format="$#,##0.00"|"MM/dd/yyyy"|etc
        '       width = "90", readonly =T|F
        ' with every column having at least a pair of placeholders,
        ' including optional checkbox column
        Dim i As Integer
        Dim ts As New DataGridTableStyle()
        Dim myDataCol As New DataGridBoolColumn() ' checkbox column

        Try

            dg.TableStyles.Clear()
            ts.GridColumnStyles.Clear()

            ' map the table style to the dt
            ts.MappingName = dt.TableName '"dt"

            ' for the rest of the rows, make text box columns
            Dim j As Integer '= IIf(AddCkBox.Length > 0, 1, 0)

            For i = 0 To dt.Columns.Count - 1
                Dim dgtbc As New DataGridTextBoxColumn()
                dgtbc.HeaderText = _
                   FixGridColumnCaption(dt.Columns.Item(i).ColumnName)
                dgtbc.MappingName = dt.Columns.Item(i).ColumnName
                If UBound(Formats) > 0 Then
                    Try
                        j = i * 4
                        dgtbc.Format = Formats(j)
                        dgtbc.Width = Val(Formats(j + 1))
                        dgtbc.ReadOnly = IIf(Formats(j + 2) = "T", True, False)
                        If Formats(j + 3) = "C" Then
                            dgtbc.Alignment = HorizontalAlignment.Center
                        ElseIf Formats(j + 3) = "L" Then
                            dgtbc.Alignment = HorizontalAlignment.Left
                        ElseIf Formats(j + 3) = "R" Then
                            dgtbc.Alignment = HorizontalAlignment.Right
                        End If
                    Catch ex As Exception
                        'StructuredErrorHandler(ex)
                    End Try
                End If
                dgtbc.NullText = String.Empty
                ts.GridColumnStyles.Add(dgtbc)
                ts.AllowSorting = False
            Next

            ' add the table style to the grid
            dg.TableStyles.Add(ts)

        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

    Public Function AddRowToTable(ByRef dt As DataTable, ByVal ParamArray DRows() As Object) As Boolean
        Dim i As Short
        Try
            Dim newRow As DataRow = dt.NewRow
            For i = 0 To UBound(DRows)
                ' add a row to the passed dtList
                newRow(i) = DRows(i)
            Next
            dt.Rows.Add(newRow)
            Return True
        Catch ex As System.Exception
            MsgBox(ex.ToString)
            Return False
        End Try
    End Function

    Public Sub AddCheckBoxColumn(ByRef dt As DataTable, _
       ByVal CName As String)
        ' adds the column for the checkbox to the dt and set to false
        ' the column will be placed at the end of the datatable,
        ' but it will appear in the grid in the first column
        Try
            dt.Columns.Add(CName)
            Dim i As Integer
            For i = 0 To dt.Rows.Count - 1
                dt.Rows(i).Item(CName) = "false"
            Next
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

    Public Function GetClickedRow(ByVal e As  _
       System.Windows.Forms.MouseEventArgs, _
        ByRef dbg As DataGrid) As Integer
        ' this method is called from the mouseup event of a datagrid
        ' returns the clicked row number (zero based)
        Try
            Dim pt = New Point(e.X, e.Y)
            Dim hti As DataGrid.HitTestInfo = dbg.HitTest(pt)
            Return hti.Row
        Catch ex As System.Exception
            ' ignore the error if the user clicked outside the grid rows,
            ' e.g., in the header...
        End Try
    End Function

    Public Sub GetClickedCell(ByRef iRow As Integer, _
        ByRef iCol As Integer, _
        ByRef dbg As DataGrid, _
        ByVal e As System.Windows.Forms.MouseEventArgs)
        Try
            Dim pt = New Point(e.X, e.Y)
            Dim hti As DataGrid.HitTestInfo = dbg.HitTest(pt)

        Catch ex As System.Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Public Overloads Function SelectCkBoxRow(ByRef dt As DataTable, _
       ByRef dbg As DataGrid, _
       ByVal e As System.Windows.Forms.MouseEventArgs, _
       ByVal CBName As String, _
       ByRef CkUnCk As Boolean) As Integer
        Dim retval As Integer

        Try
            Dim pt = New Point(e.X, e.Y)
            Dim hti As DataGrid.HitTestInfo = dbg.HitTest(pt)

            If hti.Type = DataGrid.HitTestType.Cell Then
                dbg.CurrentCell = New DataGridCell(hti.Row, hti.Column)

                If dt.Rows(dbg.CurrentRowIndex).Item(CBName) = "true" Then
                    dt.Rows(dbg.CurrentRowIndex).Item(CBName) = "false"
                    CkUnCk = False
                Else
                    dt.Rows(dbg.CurrentRowIndex).Item(CBName) = "true"
                    CkUnCk = True
                End If
                dbg.Select(hti.Row)
            End If
            Return hti.Row
        Catch ex As System.Exception
        End Try
    End Function

    ' selects any row clicked and returns the hit row
    Public Overloads Function SelectCkBoxRow(ByRef dg As DataGrid) As Integer
        Return SelectCkBoxRow(dg, dg.CurrentCell.RowNumber)
    End Function

    ' Select the row indicated by x,y position of the mouse.
    Public Overloads Function SelectCkBoxRow(ByRef dbg As DataGrid, _
       ByVal e As System.Windows.Forms.MouseEventArgs) As Integer

        Try
            Dim pt = New Point(e.X, e.Y)
            Dim hti As DataGrid.HitTestInfo = dbg.HitTest(pt)
            Return SelectCkBoxRow(dbg, hti.Row)
        Catch
            Return 0
        End Try
    End Function

    ' Select row designated by the row parameter.
    Public Overloads Function SelectCkBoxRow(ByRef dg As DataGrid, _
       ByVal row As Integer) As Integer
        Try
            dg.Select(row)
        Catch
        End Try
        Return row
    End Function

    ' This overloaded method only sets the chkbox and selects
    ' the row if the user clicks in the checkbox cell.
    Public Overloads Function SelectCkBoxRow(ByRef dt As DataTable, _
       ByRef dbg As DataGrid, _
       ByVal e As System.Windows.Forms.MouseEventArgs, _
       ByVal CBName As String, _
       ByRef CkUnCk As Boolean, _
       ByVal colNum As Integer) As Integer
        Dim retval As Integer

        Try
            Dim pt = New Point(e.X, e.Y)
            Dim hti As DataGrid.HitTestInfo = dbg.HitTest(pt)

            If hti.Type = DataGrid.HitTestType.Cell Then
                dbg.CurrentCell = New DataGridCell(hti.Row, hti.Column)

                If colNum = hti.Column Then
                    If dt.Rows(dbg.CurrentRowIndex).Item(CBName) = "true" Then
                        dt.Rows(dbg.CurrentRowIndex).Item(CBName) = "false"
                        CkUnCk = False
                    Else
                        dt.Rows(dbg.CurrentRowIndex).Item(CBName) = "true"
                        CkUnCk = True
                    End If
                    dbg.Select(hti.Row)
                Else
                    CkUnCk = False
                End If
                Return hti.Row
            End If
        Catch ex As System.Exception
        End Try
    End Function

    ' This overloaded method simply select the current grid row
    Public Overloads Function SelectCkBoxRow(ByRef dt As DataTable, _
       ByRef dbg As DataGrid, _
       ByVal CBName As String) As Integer
        Dim retval As Integer
        dt.Rows(dbg.CurrentRowIndex).Item(CBName) = "true"
        dbg.Select(dbg.CurrentRowIndex)
    End Function

    Function ProperName(ByVal psIn As String) As String
        Dim sTemp As String ' original string
        Dim sTemp2 As String ' next word
        Dim sTemp3 As String ' staging string for return
        Dim i As Integer

        ' convert the passed string to lower case
        If psIn.IndexOf(" ") > -1 Then
            sTemp = psIn.ToLower
        Else
            sTemp = psIn
        End If
        sTemp2 = ""

        Do While Trim$(sTemp) <> ""
            GetNextPNWord(sTemp3, sTemp)
            If UCase(Left(sTemp3, 2)) = "MC" Then
                sTemp2 &= "Mc" & UCase$(Mid$(sTemp3, 3, 1)) & Mid$(sTemp3, 4)
            Else
                sTemp2 &= UCase$(Left$(sTemp3, 1)) & Mid$(sTemp3, 2) & " "
            End If
        Loop

        Return sTemp2
        Exit Function
    End Function


#End Region


#Region " Private Methods "
    Private Sub GetNextPNWord(ByRef sTemp3 As String, ByRef sTemp As String)
        ' return next space delimited word in stemp3 until
        ' stemp2 is exhausted
        Dim i As Integer = InStr(sTemp, " ")
        If i > 0 Then
            sTemp3 = Left$(sTemp, i - 1)
            sTemp = Trim$(Mid$(sTemp, i))
        Else
            sTemp3 = sTemp
            sTemp = ""
        End If
    End Sub

#End Region





End Class
