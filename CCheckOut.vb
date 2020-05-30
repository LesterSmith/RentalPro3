''' Handles miscellaneous functions for the check out
''' function.
Public Class CCheckOut
#Region " Class Level Variables "
   Private SQL As String
   Private oDA As New CDataAccess()
   Private dt As New DataTable()


#End Region


#Region " Private Methods "
   ''' <summary>
   ''' Returns true if equipment is available.
   ''' </summary>
   ''' <param name = "equipID"></param>
   ''' <returns>boolean</returns>
   Public Function IsEquipAvailable(ByVal equipID As String, Optional ByVal Msg As Boolean = False) As Boolean
      SQL = "select equip_id from equipment "
      SQL &= "where equip_id ='" & equipID & "' "
      SQL &= "and available = 'YES' "
      SQL &= "and (isnull(damage) or damage = '' or damage = 'R') "
      If oDA.SendQuery(SQL, dt, ConnectString) > 0 Then
         Return True
      Else
         If Msg Then
            MsgBox("Another customer currently has the equipment on hold or rented or equipment is damaged", MsgBoxStyle.Exclamation)
         End If
         Return False
      End If
   End Function

#End Region


End Class
