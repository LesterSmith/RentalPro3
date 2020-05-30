'****************************************
'* Purpose:
'*
'* Author:  Les Smith
'* Date Created: 05/15/2003 at 09:09:39
'* CopyRight:  HHI Software
'****************************************
'*
Option Strict Off
Option Explicit On 
Friend Class CCustomer
#Region " Module Variables "
   Dim SQL As String
   Dim dt As New DataTable()
   Dim oDA As New CDataAccess()


#End Region


#Region " Public Methods "
   Public Function GetNewCustID() As Integer
      Try
         SQL = "select max(customerid) from customers "
         dt = New DataTable()
         oDA.SendQuery(SQL, dt, ConnectString)

         If dt.Rows.Count = 0 Then
            Return 1001
         End If

         Return dt.Rows(0).Item(0) + 1
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
         Return 0
      End Try
   End Function

#End Region




#Region " Constructor "
   Public Sub New()
      oDA = New CDataAccess()
   End Sub

#End Region

End Class