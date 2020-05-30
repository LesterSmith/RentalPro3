''' This class contains often called functions so
''' that we don't have to instantiate an object.
''' 
Public Class CShare
   ''' <summary>
   ''' returns time due back
   ''' </summary>
   ''' <param name = "period"></param>
   ''' <param name = "Qty"></param>
   ''' <param name = "nowTime"></param>
   ''' <returns>DateTime</returns>
   Public Shared Function GetDueBackTime(ByVal period As String, ByVal Qty As Integer, ByVal nowTime As DateTime) As DateTime

      Dim dtDueBack As DateTime

      Try
         Select Case period
            Case DAILY : dtDueBack = DateAdd(DateInterval.Day, Qty, nowTime)
            Case MONTHLY : dtDueBack = DateAdd(DateInterval.Month, Qty, nowTime)
            Case HALF_DAY : dtDueBack = DateAdd(DateInterval.Hour, 4, nowTime)
            Case WEEK_END : dtDueBack = DateAdd(DateInterval.Day, 2, nowTime)
            Case WEEKLY : dtDueBack = DateAdd(DateInterval.Day, 5, nowTime)
         End Select
         Return dtDueBack
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Function

End Class
