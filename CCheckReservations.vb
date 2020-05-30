' Validate reservations and rentals against reservations.
Public Class CCheckReservations
   Private oDA As CDataAccess

#Region " Public Methods "
   Public Sub New()
      oDA = New CDataAccess()
   End Sub

   Public Function IsEquipAvailable(ByVal equipClass As Integer, ByVal reqRentDate As DateTime, ByVal reqRentEndDate As DateTime) As Boolean
      Dim Sql As String
      Dim dt As New DataTable()
      Dim i As Integer
      Dim resCount As Short
      Dim reserveStartDate As DateTime
      Dim availableDate As DateTime
      Dim available As String
      Dim nowPlusOneDay As DateTime = DateAdd(DateInterval.Day, 1, Now())
      Dim sNowPlusOneDay As String = nowPlusOneDay.ToString

      Try
         ' get recordset of union of equipment and reservations tables
         Sql = "select price_id,rented_date,iif(available='ON RENT',#"
         Sql &= sNowPlusOneDay & "#,available_date) as avail_date,Available,user_id "
         Sql &= "from equipment "
         Sql &= "where  price_id = " & equipClass & " "
         Sql &= "union "
         Sql &= "select equip_class as price_id, "
         Sql &= "res_date as rented_date, "
         Sql &= "res_end_date as avail_date,'RES' as Available,'RES' as user_id "
         Sql &= "from reservations "
         Sql &= "where equip_class = " & equipClass & " "

         If oDA.SendQuery(Sql, dt, ConnectString) > 0 Then
            Dim dr As DataRow
            For i = 0 To dt.Rows.Count - 1
               ' now loop thru the dt 
               ' count and discounting the possibilities
               ' if rescount > 0 when done, we can reserve
               dr = dt.Rows(i)
               available = dr("available")
               ' if available="YES" this is an equipment record
               ' and we can add to the availability count
               If available = "YES" Then
                  resCount += 1
                  GoTo GetNextRow
               ElseIf available = "ON RENT" Then
                  ' if it's on rent, it is not available right now
                  ' and since this is only called by isrentable,
                  ' that's all that matters
                  GoTo GetNextRow
               ElseIf available = "ON HOLD" Then
                  ' on hold takes a piece of equipment from contention
                  ' only if computername is not me, does nothing to counter
                  ' if user_id <> RES and <> ME then it's on hold from another computer
                  ' so it's probably being rented...
                  If dr("user_id") = modMain.UserName Then
                     resCount += 1
                     GoTo GetNextRow
                  End If
               End If

               ' equip is reserved
               If IsDBNull(dr("rented_date")) Or IsDBNull(dr("avail_date")) Then
                  GoTo GetNextRow
               End If

               reserveStartDate = dr("rented_date")

               availableDate = dr("avail_date")
               ' here is where it gets hairy
               ' compute any possible overlap


               ' if the rent start date is w/i the reserved period
               If reqRentDate >= reserveStartDate AndAlso _
                  reqRentDate <= availableDate Then
                  ' rentstart is w/i res period
                  resCount -= 1
               ElseIf reqRentEndDate >= reserveStartDate AndAlso _
                  reqRentEndDate <= availableDate Then
                  ' rentenddate is in the res period
                  resCount -= 1
               ElseIf reqRentDate <= reserveStartDate And _
                  reqRentEndDate >= availableDate Then
                  ' rent period spans the res period
                  resCount -= 1
               Else
                  ' a reservation can't add to available equipment
                  ' because it can never represent availability,
                  ' only unavailability...
                  If available <> "RES" Then
                     resCount += 1
                  End If
               End If
GetNextRow:
            Next i
            Return (resCount > 0)
         Else
            Return False
         End If
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Function

   ''' <summary>
   ''' Here we have a request to rent and we need to search the 
   ''' reservations table to see if we have a piece of equipment
   ''' that is rentable.  To  do that we need to get a recordset
   ''' of equip from the equip table for the requested equip class
   ''' (price_id) 
   ''' </summary>
   ''' <param name = "equipClass"></param>
   ''' <param name = "resDate"></param>
   ''' <param name = "resPeriod"></param>
   ''' <param name = "nbrPeriods"></param>
   ''' <returns>Boolean</returns>
   Public Function IsReservable(ByVal equipClass As Integer, ByVal reqResDate As DateTime, ByVal reqResEndDate As DateTime) As Boolean
      Dim Sql As String
      Dim dt As New DataTable()
      Dim i As Integer
      Dim resCount As Short
      Dim rentedDate As DateTime
      Dim availableDate As DateTime
      Dim available As String
      Dim nowPlusOneDay As DateTime = DateAdd(DateInterval.Day, 1, Now())
      Dim sNowPlusOneDay As String = nowPlusOneDay.ToString
      Dim table As String

      Try
         ' get recordset of union of equipment and reservations tables
         Sql = "select price_id,rented_date,iif(available='ON RENT',#"
         Sql &= sNowPlusOneDay & "#,available_date) as avail_date,Available,'EQU' as tbl "
         Sql &= "from equipment "
         Sql &= "where  price_id = " & equipClass & " "
         Sql &= "union "
         Sql &= "select equip_class as price_id, "
         Sql &= "res_date as rented_date, "
         Sql &= "res_end_date as avail_date,'RES' as Available,'RES' as tbl "
         Sql &= "from reservations "
         Sql &= "where equip_class = " & equipClass & " "

         If oDA.SendQuery(Sql, dt, ConnectString) > 0 Then
            Dim dr As DataRow
            For i = 0 To dt.Rows.Count - 1
               ' now loop thru the dt 
               ' count and discounting the possibilities
               ' if rescount > 0 when done, we can reserve
               dr = dt.Rows(i)
               available = dr("available") : table = dr("tbl")
               If available = "YES" Then
                  resCount += 1
                  GoTo GetNextRow
               ElseIf available = "ON RENT" Then
                  resCount += 1
               ElseIf available = "ON HOLD" Then
                  MsgBox("You have equipment of requested type on Hold for pending check out, please finish or cancel checkout and try making reservation again.", MsgBoxStyle.Exclamation)
                  Return False
               End If

               rentedDate = dr("rented_date")
               availableDate = dr("avail_date")
               ' here is where it gets hairy
               ' compute any possible overlap
               If reqResDate >= rentedDate And _
                  reqResDate <= availableDate Then
                  resCount -= 1
               ElseIf reqResEndDate >= rentedDate And _
                  reqResEndDate <= availableDate Then
                  resCount -= 1
               Else
                  ' a reservation can't add to available equipment
                  ' because it can never represent availability,
                  ' only unavailability...
                  If available <> "RES" Then
                     resCount += 1
                  End If
               End If
GetNextRow:
            Next i
            Return (resCount > 0)
         Else
            Return False
         End If
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Function

   ''' <summary>
   ''' Returns true if requested equipment rental does not conflict
   ''' with existing reservations.
   ''' </summary>
   ''' <param name = "equipClass"></param>
   ''' <param name = "resDate"></param>
   ''' <param name = "resPeriod"></param>
   ''' <param name = "nbrPeriods"></param>
   ''' <returns>Boolean</returns>
   Public Function IsRentable(ByVal equipClass As Integer, _
      ByVal rentDate As Date, _
      ByVal rentPeriod As String, _
      ByVal nbrPeriods As Integer) As Boolean
      Dim Sql As String
      Dim dt As New DataTable()
      Dim i As Integer
      Dim resCount As Short
      Dim reservedDate As DateTime
      Dim availableDate As DateTime
      Dim available As String
      Dim rentEndDate As DateTime

      Try
         ' the equipment table has price_id, herein known as equipment class
         ' we must determine if a requested rental or reservation spans into any
         ' area
         ' first we get a recordset of available equipment of the requested class

         Select Case rentPeriod
            Case HOURLY
               rentEndDate = DateAdd(DateInterval.Hour, CType(nbrPeriods, Double), rentDate)
            Case DAILY
               rentEndDate = DateAdd(DateInterval.Day, CType(nbrPeriods, Double), rentDate)
            Case HALF_DAY
               rentEndDate = DateAdd(DateInterval.Hour, 5, rentDate)
            Case WEEKLY
               rentEndDate = DateAdd(DateInterval.Day, 7 * CType(nbrPeriods, Double), rentDate)
            Case MONTHLY
               rentEndDate = DateAdd(DateInterval.Month, CType(nbrPeriods, Double), rentDate)
            Case WEEK_END
               rentEndDate = DateAdd(DateInterval.Day, CType(nbrPeriods, Double), rentDate)
         End Select

         If Not IsEquipAvailable(equipClass, rentDate, rentEndDate) Then
            Dim sMsg As String
            Dim iRV As Integer
            sMsg = "This equipment is reserved for a customer.  You can" & Chr(10)
            sMsg &= "click the Reserve Button to see who has this equipment" & Chr(10)
            sMsg &= "reserved and if it is for the current customer, you can " & Chr(10)
            sMsg &= "delete their reservation so that you can rent the equipment." & Chr(10)
            sMsg &= "" & Chr(10)
            sMsg &= "You can Click Ok to ignore the reservation, " & Chr(10)
            sMsg &= "or click Cancel to view the Reservation List." & Chr(10)
            sMsg &= "" & Chr(10)
            iRV = MsgBox(sMsg, CType(49, Microsoft.VisualBasic.MsgBoxStyle), "Reservation Conflict")

            If iRV = 1 Then
               ' Ok Code goes here
               Return True
            Else
               ' Cancel code goes here
               Return False
            End If
         End If
         Return True
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Function

#End Region
End Class
