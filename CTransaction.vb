Option Strict Off
Option Explicit On
Friend Class CTransaction
#Region " Class Level Variables "
    Dim SQL As String
    Dim oDA As CDataAccess


#End Region

#Region " Public Methods "
    Public Function ReserveTemp(ByRef rsEquipID As String) As Boolean
        ' this method will mark the equipment as temporarily
        ' reserved so we can't rent it again
        Dim sErr As String
        Dim dt As New DataTable()

        Try
            SQL = "select available from equipment "
            SQL = SQL & "where equip_id = '" & rsEquipID & "' "
            SQL = SQL & "and ucase(available) = 'YES' "

            If oDA.SendQuery(SQL, dt, ConnectString) = 0 Then
                MsgBox("This equipment is not available for rent.", MsgBoxStyle.Information)
                ReserveTemp = False
                Exit Function
            End If

            ' it's available, reserve it
            SQL = "update equipment set available = 'ON HOLD' "
            SQL &= ",user_id ='" & UserName & "' "
            SQL = SQL & "where equip_id = '" & rsEquipID & "'"
            oDA.SendActionSql(SQL, ConnectString, sErr)

            ReserveTemp = True

            Exit Function

        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Function

    ''' <summary>
    ''' See if we have any rerents and attempt to restore
    ''' them in the rerents table.
    ''' </summary>
    Private Sub RestoreRerents()
        Dim i As Integer
        Dim sql As String
        Dim dt As New DataTable("dt")
        Dim j As Integer
        Dim dr As DataRow
        Dim serr As String

        Try
            sql = "select itemid,rerent_id from tempitems "
            sql &= "where user_id = '" & modMain.UserName & "' "
            sql &= "and itemid = '" & RERENT & "' "

            If oDA.SendQuery(sql, dt, ConnectString, "dt") > 0 Then
                For i = 0 To dt.Rows.Count - 1
                    dr = dt.Rows(i)
                    ' get new recordset of rows matching the name, period,nbrperiods
                    sql = "update rerents set rented_date = null where unique_id = " & dr("rerent_id")
                    If oDA.SendActionSql(sql, ConnectString, serr) <> 1 Then
                        Throw New System.Exception("Update failed to reset rerent: " & dr("rerent_id") & Chr(10) & serr)
                    End If
                Next
            End If
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

    Public Overloads Sub RemoveTempReservation(ByVal All As Boolean)
        ' This method removes the temp hold on the item
        Dim lRowsAffected As Integer
        Dim sErr As String

        Try
            SQL = "update equipment set available = 'YES',user_id = Null "
            If All Then
                SQL = SQL & "where available = 'ON HOLD'  or available = 'SOLD HOLD' "
            Else
                SQL = SQL & "where available = 'ON HOLD' "
            End If
            SQL &= "and user_id = '" & UserName & "' "
            oDA.SendActionSql(SQL, ConnectString, sErr)

            ' restore rerents if any
            Me.RestoreRerents()

            ' now delete temporary items table
            SQL = "delete from tempitems where user_id = '" & UserName & "'"
            oDA.SendActionSql(SQL, ConnectString, sErr)
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

    ''' <summary>
    ''' This method removes the temp hold on the item
    ''' </summary>
    ''' <param name = "rsEquipID"></param>
    Public Overloads Sub RemoveTempReservation(ByRef rsEquipID As String)
        Dim lRowsAffected As Integer
        Dim sErr As String

        Try
            SQL = "update equipment set available = 'YES' "
            If rsEquipID = String.Empty Then
                SQL &= ",user_id=Null "
                SQL = SQL & "where available = 'ON HOLD'   "
                SQL &= "and user_id = '" & UserName & "' "
            Else
                SQL = SQL & "where equip_id = '" & rsEquipID & "'"
            End If
            oDA.SendActionSql(SQL, ConnectString, sErr) 'ignoring return code here

            ' restore rerents if any
            Me.RestoreRerents() ' we get error here if we can't make connection

            ' now delete temporary items table
            SQL = "delete from tempitems where user_id = '" & UserName & "'"
            oDA.SendActionSql(SQL, ConnectString, sErr) 'ignoring return code here
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

    Private Function PlaceOnRent(ByRef rsEquipID As String) As Boolean
        ' this method will mark the equipment as temporarily
        ' reserved so we can't rent it again
        Dim sErr As String
        Dim dt As New DataTable()


        Try
            SQL = "select available from equipment "
            SQL = SQL & "where equip_id = '" & rsEquipID & "' "
            SQL = SQL & "and ucase(available) = 'ON HOLD' "

            If oDA.SendQuery(SQL, dt, ConnectString) = 0 Then
                MsgBox("This equipment is not available for rent.", MsgBoxStyle.Information)
                PlaceOnRent = False
                Exit Function
            End If

            ' it's available, reserve it
            SQL = "update equipment set available = 'ON RENT' "

            SQL = SQL & "where equip_id = '" & rsEquipID & "'"
            oDA.SendActionSql(SQL, ConnectString, sErr)
            PlaceOnRent = True
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Function

    Public Function PlaceShoppingCartListOnRent() As Object
        Dim i As Short
        Dim oItem As CItems
        Dim b As Boolean
        Dim dt As New DataTable()

        Try
            SQL = "select * from tempitems"
            If oDA.SendQuery(SQL, dt, ConnectString) > 0 Then
                For i = 0 To dt.Rows.Count - 1
                    With dt.Rows(i)
                        If dt.Rows(i).Item("rentorsale") = "Rent" And
                           dt.Rows(i).Item("itemid") <> RERENT Then
                            b = PlaceOnRent(.Item("ItemID"))
                        End If
                    End With
                Next i
            Else
                Throw New System.Exception("Found no items reserved in temp items table.")
            End If
        Catch ex As System.Exception
            'StructuredErrorHandler(ex)
        End Try
    End Function


#End Region

#Region " Constructors "
    Public Sub New()
        oDA = New CDataAccess()
    End Sub

#End Region

End Class