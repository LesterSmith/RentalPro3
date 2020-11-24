'****************************************
'* Purpose: Handles action queries with code
'* for handling locks and unique constraint violations.
'*
'* Author:  Les Smith
'* Date Created: 09/17/2002 at 02:15:16
'****************************************
Imports System.Data.OleDb
Imports System.Windows.Forms.Application
Public Class CDataAccess
    Implements IDisposable
#Region " Class Level Variables "
    Dim dbCmdOle As New OleDbCommand()
    Dim daOle As New OleDbDataAdapter()
    Dim ConnOle As New OleDbConnection()
    Dim mytrans As OleDbTransaction


#End Region


#Region " Public Methods "
    Public Overloads Function SendQuery(ByVal Sql As String, ByRef dt As DataTable, ByRef Connection As OleDbConnection) As Integer
        ' overloaded function that accepts a connection
        ' and does not connnect and disconnect 
        ' used when program is going to do numerous hits
        ' to database in a loop
        ' Returns number of rows affected
        ' If error, ErrNum and ErrMsg will have the
        ' respective values
        ' dt will be filled
        ' Returns -1 if connection can't be made
        ' returns -2 on any other error
        Dim localDT As New DataTable()

        Try
            dbCmdOle.CommandText = Sql
            dbCmdOle.Connection = Connection
            daOle = New OleDbDataAdapter(dbCmdOle)
            daOle.Fill(localDT)
            dt = localDT
            Return localDT.Rows.Count

        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Function

    Public Overloads Function OpenConnection(ByVal ConnStr As String, ByRef Conn As OleDbConnection) As Boolean
        ' opens a connection to the database and return true if successful
        Dim i As Integer
        On Error Resume Next
        Do
            Conn.ConnectionString = ConnStr
            Conn.Open()
            If Err.Number <> 0 Then
                Dim start As Double = Microsoft.VisualBasic.Timer
                Do While Microsoft.VisualBasic.Timer - start < 1
                    DoEvents()
                Loop

                i += 1
                If i > 5 Then
                    MsgBox("Can't open conntection to " & ConnStr & vbCrLf & " Err: " & Err.Description, MsgBoxStyle.Critical)
                    Return False
                End If
                Err.Clear()
            Else
                Exit Do
            End If
        Loop
        Return True
    End Function

    Public Overloads Function SendActionSql(ByVal Sql As String,
       ByRef conn As OleDbConnection,
       ByRef ErrMsg As String) _
       As Integer
        '' Executes the passed action sql.  Returns the following:
        '' iRowsAffected if successful
        '' -1 = unique constraint violation
        '' -2 = locked and retries exceeded
        '' -3 = other error (ErrMsg will contain err.description)
        Dim iStart As Integer
        Dim iRowsAffected As Integer
        Dim dbCmd As New OleDbCommand()
        Dim errCnt As Integer
        ErrMsg = ""
        On Error Resume Next
        dbCmd.CommandText = Sql
        dbCmd.Connection = conn
TryAgain:
        iRowsAffected = dbCmd.ExecuteNonQuery

        ' Check results
        If Err.Number <> 0 Then
            ' we had an error , ck it out
            If InStr(1, Err.Description, "locked", 1) > 0 Then
                ' we are locked out of the database
                Err.Clear()

                errCnt += 1
                If errCnt > 10 Then
                    Return -2
                End If
                iStart = Microsoft.VisualBasic.Timer
                Do While Microsoft.VisualBasic.Timer - iStart < 1
                    DoEvents()
                Loop
                GoTo TryAgain
            ElseIf InStr(1, Err.Description, "duplicate", 1) > 0 Then
                Return -1
            Else
                ErrMsg = Err.Description
                Return -3
            End If
        Else
            Return iRowsAffected
        End If
    End Function

    Public Overloads Function SendActionSql(ByVal dbCmd As OleDb.OleDbCommand, ByRef errmsg As String) As Integer
        Try
            Dim rowsAffected As Integer = dbCmd.ExecuteNonQuery
            Return rowsAffected
        Catch ex As System.Exception
            Return -1
        End Try
    End Function

    Public Overloads Function SendActionSql(ByVal Sql As String,
       ByVal ConnectString As String,
       ByRef ErrMsg As String,
       Optional ByVal SS_OLE_Flag As String = "O") _
       As Integer
        '' Executes the passed action sql.  Returns the following:
        '' iRowsAffected if successful
        '' -1 = unique constraint violation
        '' -2 = locked and retries exceeded
        '' -3 = other error (ErrMsg will contain err.description)
        Dim iStart As Integer
        Dim iRowsAffected As Integer
        Dim dbCmdOle As New OleDbCommand()
        Dim ConnOle As New OleDbConnection()
        Dim mytrans As OleDbTransaction
        'Dim dbCmdSS As New SqlCommand()
        'Dim ConnSS As New SqlConnection()
        Dim errCnt As Integer
        ErrMsg = ""
        On Error Resume Next
        If Not Me.OpenConnection(ConnectString, ConnOle) Then
            ErrMsg = "Unable to connect to " & ConnectString
            Return -3
        End If
        mytrans = ConnOle.BeginTransaction()
        dbCmdOle.CommandText = Sql
        dbCmdOle.Connection = ConnOle
        dbCmdOle.Transaction = mytrans

TryAgain:
        iRowsAffected = dbCmdOle.ExecuteNonQuery
        mytrans.Commit()

        ' Check results
        If Err.Number <> 0 Then
            ' we had an error , ck it out
            If InStr(1, Err.Description, "locked", 1) > 0 Then
                ' we are locked out of the database
                Err.Clear()

                errCnt += 1
                If errCnt > 100 Then
                    Return -2
                End If
                iStart = Microsoft.VisualBasic.Timer
                Do While Microsoft.VisualBasic.Timer - iStart < 1
                    DoEvents()
                Loop
                GoTo TryAgain
            ElseIf InStr(1, Err.Description, "duplicate", 1) > 0 Then
                Return -1
            Else
                ErrMsg = Err.Description
                Return -3
            End If
        Else
            Return iRowsAffected
        End If
    End Function

    Public Overloads Function SendQuery(ByVal Sql As String,
       ByRef dt As DataTable,
       ByVal ConnectString As String,
       ByVal dtStr As String) As Integer

        ' Returns number of rows affected
        ' If error, ErrNum and ErrMsg will have the
        ' respective values
        ' dt will be filled
        ' Returns -1 if connection can't be made
        ' returns -2 on any other error
        ' SSOLE = "S" for SqlSrvr or "O" for Ole
        ' This function supports OLE
        Dim localDT As New DataTable(dtStr)

        Try
            If Not Me.OpenConnection(ConnectString, ConnOle) Then
                Return -1 ' can't connect to db
            End If
            Me.dbCmdOle.CommandText = Sql
            dbCmdOle.Connection = ConnOle
            daOle = New OleDbDataAdapter(dbCmdOle)
            daOle.Fill(localDT)
            ConnOle.Close()
            dt = localDT
            Return localDT.Rows.Count

        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Function

    Public Overloads Function SendQuery(ByVal Sql As String,
       ByRef dt As DataTable,
       ByVal ConnectString As String) As Integer

        ' Returns number of rows affected
        ' If error, ErrNum and ErrMsg will have the
        ' respective values
        ' dt will be filled
        ' Returns -1 if connection can't be made
        ' returns -2 on any other error
        ' SSOLE = "S" for SqlSrvr or "O" for Ole
        ' This function supports OLE
        Dim localDT As New DataTable()

        Try
            If Not Me.OpenConnection(ConnectString, ConnOle) Then
                Return -1 ' can't connect to db
            End If
            Me.dbCmdOle.CommandText = Sql
            dbCmdOle.Connection = ConnOle
            daOle = New OleDbDataAdapter(dbCmdOle)
            daOle.Fill(localDT)
            ConnOle.Close()
            dt = localDT
            Return localDT.Rows.Count

        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Function

#End Region

    Protected Overridable Sub Dispose(disposing As Boolean)
        If disposing Then
            ' TODO: dispose managed state (managed objects).
            If ConnOle.State = ConnectionState.Open Then
                ConnOle.Close()
            End If
            ConnOle.Dispose()
            dbCmdOle.Dispose()
            daOle.Dispose()

            'If your class holds references to other .NET objects 
            'that themselves implement IDisposable then you should implement IDisposable 
            'and call their Dispose method in yours

        End If

        ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.

        'If you're holding any OS resources, e.g. file or image handles, 
        'then you should release them. That will be a rare thing for most people and can be pretty much ignored

        ' TODO: set large fields to null.
        'If any of your fields may refer to objects that occupy 
        'a large amount of memory then those fields should be set to Nothing

    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub











End Class
