''' Purpose: Handles Checkin and PrintInvoice 
''' for the checkin process.  Incorporates the new
''' logic for autocalc and meter calculations.
''' Author:  Les Smith
''' Date Created: 09/19/2003 at 12:37:16
''' CopyRight:  HHI Software
'''
Public Class CCheckIn

#Region "Private Variables"
   Private frm As frmCheckinNew
   Private oDA As CDataAccess
   Private oCG As New CGrid()
   Private elapsedDays As Integer
   Private elapsedHours As Integer
   Private allowedHoursRun As Single
   Private actualHoursRun As Single
   Private sql As String
   Dim origRowCount As Integer
   Private CheckIn As DateTime = Now
#End Region
#Region "Class Constructor"
   ''' <summary>
   ''' Overloaded constructor, called by frmCheckinNew
   ''' </summary>
   ''' <param name = "o"></param>
   Public Sub New(ByVal o As frmCheckinNew)
      frm = o
      oDA = New CDataAccess()
      oCG = New CGrid()
   End Sub
   ''' <summary>
   ''' Overloaded constructor, Called by PrintInvoices
   ''' </summary>
   Public Sub New()
      oDA = New CDataAccess()
   End Sub
#End Region
#Region "Auto Checkin Driver"
   ''' <summary>
   ''' Auto calculate the price of each piece of equipment.
   ''' Base pricing upon elapsed time and meter usage.
   ''' First, compute the elapsed time between rented time
   ''' and now, allowing for grace time at checkin.  Next,
   ''' see if any equipment was overused, and if so add
   ''' The methodology is to make the current record
   ''' be the greatest period as possible
   ''' and then work down the remaing elapsed hours by
   ''' adding new rows with the period as large as possible
   ''' </summary>
   Public Function AutoCalcCheckIn(Optional ByVal chkInDate As DateTime = #1/1/2001#) As DateTime
      Dim i As Integer
      Dim usedCurrRow As Boolean = False
      Dim price As Single
      Dim totPrice As Single = 0
      Dim allowedUseHours As Integer
      Dim bNewPrices As Boolean
      Try
         origRowCount = frm.dtList.Rows.Count
         ' only look at the original items, else we calc on added items
         ' and get in recursion loop...
         For i = 0 To origRowCount - 1
            Dim dr As DataRow = frm.dtList.Rows(i)

            ' if newprices field is extant we use new get pricing method
            Try
               bNewPrices = dr("newprices")
            Catch
               bNewPrices = False
            End Try
            If dr("equip_id") = RERENT And Not MNB(dr("newprices")) Then
               MsgBox("We have no price table for ReRents, so they have to be calculated manually.", MsgBoxStyle.Information)
               GoTo CalcNextRow
            End If

            If dr("rental_period") <> SALE And dr("rental_period") <> "N/A" Then
               totPrice = 0
               Dim rt As New RentalTime()
               rt.Days = 0
               rt.Weeks = 0
               rt.HalfDays = 0
               rt.Months = 0
               Dim CheckOut As DateTime = dr("rented_date") 'dtpCheckOut.Value.Date
               If chkInDate = #1/1/2001# Then
                  CheckIn = Now  'dtpCheckIn.Value.Date
               Else
                  CheckIn = chkInDate
               End If

               ' if equipment was checked out in the future
               ' we must prompt user for future checkin date

               If CheckOut > CheckIn Then
                  Dim ciDate As String
                  Do
                     ciDate = InputBox("Checkout date: " & CheckOut.ToString & _
                        " was in the future.  The checkin date must be beyond that date" & _
                        " or Autocalc can't compute the invoice items. " & _
                        "Enter the new date and time as 'mm/dd/yyyy hh:mm AM/PM'. " & _
                        "If you enter nothing, you can manually calculate the invoice.", _
                        "Enter Future Checkin Date", "")
                     If ciDate.Length > 0 Then
                        Try
                           CheckIn = CType(ciDate, DateTime)
                           If CheckIn > CheckOut Then
                              Exit Do
                           End If
                        Catch
                           MsgBox("The date and time must be in the exact format as specified by the Input prompt.", MsgBoxStyle.Exclamation)
                        End Try
                     Else
                        Return CheckOut
                     End If
                     ciDate = String.Empty
                  Loop
               End If

               modAutoCalc.AutoCalc(CheckOut, CheckIn, rt)

               Dim sAutoCalcExplain As String = "("
               With rt
                  If .Months > 0 Then
                     sAutoCalcExplain &= .Months.ToString & "Mn"
                     If bNewPrices Then
                        price = Me.GetPriceFromCheckOutData(dr, MONTHLY)
                     Else
                        price = Me.GetPriceForEquip(dr("equip_id"), MONTHLY)
                     End If
                     totPrice += price * .Months
                  End If
                  If .Weeks > 0 Then
                     If sAutoCalcExplain.Length > 1 Then sAutoCalcExplain &= ","
                     sAutoCalcExplain &= .Weeks.ToString & "Wk"
                     If bNewPrices Then
                        price = Me.GetPriceFromCheckOutData(dr, WEEKLY)
                     Else
                        price = Me.GetPriceForEquip(dr("equip_id"), WEEKLY)
                     End If
                     totPrice += price * .Weeks
                  End If
                  If .Days > 0 Then
                     If sAutoCalcExplain.Length > 1 Then sAutoCalcExplain &= ","
                     sAutoCalcExplain &= .Days.ToString & "Day"
                     If bNewPrices Then
                        price = Me.GetPriceFromCheckOutData(dr, DAILY)
                     Else
                        price = Me.GetPriceForEquip(dr("equip_id"), DAILY)
                     End If
                     totPrice += price * .Days
                  End If
                  If .HalfDays > 0 Then
                     If sAutoCalcExplain.Length > 1 Then sAutoCalcExplain &= ","
                     sAutoCalcExplain &= .HalfDays.ToString & "HD"
                     If bNewPrices Then
                        price = Me.GetPriceFromCheckOutData(dr, HALF_DAY)
                     Else
                        price = Me.GetPriceForEquip(dr("equip_id"), HALF_DAY)
                     End If
                     totPrice += price * .HalfDays
                  End If

                  If sAutoCalcExplain.Length > 1 Then
                     sAutoCalcExplain &= ")"
                  Else
                     sAutoCalcExplain = String.Empty
                  End If
               End With
               If totPrice > 0 Then
                  dr("priceperunit") = totPrice
               End If
CheckMeterOverUse:
               ' now check for meter over use
               allowedUseHours = _
                  (rt.Months * HoursPerMonth) + _
                  (rt.Weeks * 40) + _
                  (rt.Days * 8) + _
                  rt.HalfDays * 4
               modAutoCalc.ElapsedHours = allowedUseHours
               modAutoCalc.ElapsedTime = sAutoCalcExplain & _
                  vbCrLf & allowedUseHours.ToString & " Allowed Meter Hrs"

               If allowedUseHours > 0 Then
                  Me.CheckForMeterOverage(frm.dtList, i, allowedUseHours)
               End If
            End If
CalcNextRow:
         Next
         Return CheckIn

      Catch ex As System.Exception
         StructuredErrorHandler(ex)
         Return Now
      End Try
   End Function

#End Region
#Region "Add Rows Method"
   ''' <summary>
   ''' This method will attempt to use the current grid row.
   ''' if it is the same rental period as the one we need to add.
   ''' if it cannot use it, it will see if it can use the last
   ''' grid row, which may have just been added with the same 
   ''' rental period as we need.  Otherwise, add a new row.
   ''' </summary>
   ''' <param name = "rentalPeriod"></param>
   ''' <param name = "cRow"></param>
   ''' <param name = "qty"></param>
   ''' <param name = "price"></param>
   Private Sub AddOrUseExistingRow(ByVal rentalPeriod As String, _
      ByVal cRow As Integer, _
      ByVal qty As Integer, _
      ByVal price As Decimal, _
      ByVal rowUsed As Boolean)
      Dim dr As DataRow = frm.dtList.Rows(cRow)
      If dr("rental_period") = rentalPeriod Then
         If rowUsed Then
            dr("quantity") += qty
         Else
            dr("quantity") = qty ' was using +=
         End If
      ElseIf Not rowUsed Then
         dr("quantity") = qty
         dr("rental_period") = rentalPeriod
         dr("priceperunit") = price
      Else
         Dim dr2 As DataRow = frm.dtList.Rows(frm.dtList.Rows.Count - 1)
         If dr2("rental_period") = rentalPeriod And _
            dr("equip_id") = dr2("equip_id") AndAlso _
            frm.dtList.Rows.Count > origRowCount Then
            dr2("quantity") += qty
         Else
            Dim dRow() As Object = {dr("equip_id"), dr("equip_name"), qty, rentalPeriod, price, dr("rented_date"), 0, 0}
            oCG.AddRowToTable(frm.dtList, dRow)
         End If
      End If
   End Sub
#End Region
#Region "Meter Reading Methods"
   ''' <summary>
   ''' Loop thru the datatable of checkin items looking for 
   ''' Items that need meter readings and prompt the user
   ''' accordingly.  Update the dt meter_in reading column.
   ''' </summary>
   Public Sub GetMeterHours()
      Dim i As Short
      Dim meterHoursIn As Single

      Try
         For i = 0 To frm.dtList.Rows.Count - 1
            With frm.dtList.Rows(i)
               If .Item("meterout") > 0 Then
                  meterHoursIn = 0
                  Do
                     meterHoursIn = Val(InputBox("Please enter the meter hours for" & Chr(10) & _
                     "Equipment: " & .Item("equip_id") & ", " & .Item("equip_name") & Chr(10) & _
                     "Meter Out: " & Format(.Item("meterout"), "####0.00"), "Enter Meter In Reading", "N.NN"))
                     If Val(meterHoursIn) < 1 Or _
                        Val(meterHoursIn) < .Item("meterout") Then
                        Dim sMsg As String
                        Dim iRV As Integer
                        sMsg = "If you do not enter the meter hours properly, any" & Chr(10)
                        sMsg &= "possible over use charges cannot be billed." & Chr(10)
                        sMsg &= "" & Chr(10)
                        sMsg &= "Click Yes to proceed without the meter reading" & Chr(10)
                        sMsg &= "or No to enter the meter reading. " & Chr(10)
                        sMsg &= "" & Chr(10)
                        iRV = MsgBox(sMsg, CType(308, Microsoft.VisualBasic.MsgBoxStyle), "Meter Hours Needed")

                        If iRV = 6 Then
                           ' Yes Code goes here
                           .Item("meterin") = 0
                           Exit Do
                        Else
                           ' No code goes here
                        End If
                     Else
                        .Item("meterin") = Val(meterHoursIn)
                        Exit Do
                     End If
                  Loop
               End If
            End With
         Next
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub



   ''' <summary>
   ''' Add item for meter overage if required.
   ''' </summary>
   ''' <param name = "dt"></param>
   ''' <param name = "i"></param>
   ''' <param name = "allowedUseHours"></param>
   Private Sub CheckForMeterOverage(ByRef dt As DataTable, ByVal i As Integer, ByVal allowedUseHours As Integer)

      Try
         If MNSng(dt.Rows(i).Item("meterout")) > 0 Then
            Dim meterUsage = MNSng(dt.Rows(i).Item("meterin")) - MNSng(dt.Rows(i).Item("meterout"))
            If meterUsage > allowedUseHours Then
               Dim rate As Decimal
               rate = Me.GetPriceForEquip(dt.Rows(i).Item("equip_id"), HOURLY)
               If rate = 0 Then
                  rate = Me.GetPriceForEquip(dt.Rows(i).Item("equip_id"), DAILY) / 8
               End If
               Dim overUsePrice As Decimal = (meterUsage - allowedUseHours) * rate
               AddMeterOverCharge(dt, i, meterUsage - allowedUseHours, rate)
            End If
         End If
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

   ''' <summary>
   ''' If the user ran the equipment more than 8 hours a day,
   ''' create a charge for the number of hours of overuse.
   ''' </summary>
   ''' <param name = "dt"></param>
   ''' <param name = "i"></param>
   ''' <param name = "charge"></param>
   Private Sub AddMeterOverCharge(ByRef dt As DataTable, _
      ByVal i As Integer, _
      ByVal overHours As Single, _
      ByVal rate As Decimal)

      Try
         Dim dRow() As Object = {dt.Rows(i).Item("equip_id"), "Meter Overuse (" & Format(overHours, "0.00") & ")", 1, "N/A", overHours * rate, Now.ToString, 0, 0}
         oCG.AddRowToTable(dt, dRow)
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

#End Region
#Region "Pricing Methods"
   ''' <summary>
   ''' retrieves price from checkout detail.
   ''' </summary>
   ''' <param name = "dt"></param>
   ''' <param name = "iRow"></param>
   ''' <returns>Decimal</returns>
   Public Function GetPriceFromCheckOutData(ByRef dr As DataRow, ByVal PriceType As String, Optional ByVal ReqPriceOnly As Boolean = False) As Decimal

      Try
         Select Case PriceType
            Case HALF_DAY
               If Not ReqPriceOnly Then
                  If MNSng(dr("halfday")) <> 0 Then
                     Return MNSng(dr("halfday"))
                  Else
                     Return MNSng(dr("daily"))
                  End If
               Else
                  Return MNSng(dr("halfday"))
               End If
            Case DAILY : Return MNSng(dr("daily"))
            Case WEEKLY : Return MNSng(dr("weekly"))
            Case MONTHLY : Return MNSng(dr("monthly"))
            Case WEEK_END : Return MNSng(dr("weekend"))
            Case HOURLY
               Dim hourRate As Decimal
               If modMain.UseHourlyRates Then
                  hourRate = MNSng(dr("hourrate"))
               Else
                  hourRate = MNSng(dr("daily")) / 8
               End If
               Return hourRate
         End Select
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Function



   ''' <summary>
   ''' Return the price_id for the passed equip_id.  It will
   ''' be used to retrieve a price for the equip
   ''' </summary>
   ''' <param name = "equipID"></param>
   ''' <param name = "PriceType"></param>
   ''' <returns>decimal</returns>
   Public Function GetPriceForEquip(ByVal equipID As String, ByVal PriceType As String, Optional ByVal ReqPriceOnly As Boolean = False) As Decimal
      Dim dt As New DataTable()

      ' if ReqPriceOnly is true return the requested
      ' price even if 0, i.e., half_day returns 0 if
      ' no halfday rate, same for hourly

      Try
         Try
            sql = "select price_id from equipment where equip_id = '" & equipID & "'"
            If oDA.SendQuery(sql, dt, ConnectString) = 0 Then
               Return 0
            End If
            sql = "select * from rental_rates where price_id = " & dt.Rows(0).Item(0)
            dt.Reset()
            If oDA.SendQuery(sql, dt, ConnectString) = 0 Then Return 0
            Dim dr As DataRow = dt.Rows(0)

            Select Case PriceType
               Case HALF_DAY
                  If Not ReqPriceOnly Then
                     If MNSng(dr("halfday")) <> 0 Then
                        Return MNSng(dr("halfday"))
                     Else
                        Return MNSng(dr("daily"))
                     End If
                  Else
                     Return MNSng(dr("halfday"))
                  End If
               Case DAILY : Return MNSng(dr("daily"))
               Case WEEKLY : Return MNSng(dr("weekly"))
               Case MONTHLY : Return MNSng(dr("monthly"))
               Case WEEK_END : Return MNSng(dr("weekend"))
               Case HOURLY
                  Dim hourRate As Decimal
                  If modMain.UseHourlyRates Then
                     hourRate = MNSng(dr("hourrate"))
                  Else
                     hourRate = MNSng(dr("daily")) / 8
                  End If
                  Return hourRate
            End Select
         Catch ex As System.Exception
            StructuredErrorHandler(ex)
         End Try
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Function

#End Region
#Region "Obsolute Code"
   ''' <summary>
   ''' Calc the cost for a monthly rntal for current grid row.
   ''' </summary>
   'Private Sub CalcMonthly(ByRef dt As DataTable, ByVal i As Integer)
   '   Dim days As Integer = elapsedHours \ 24

   'End Sub

   '''' <summary>
   '''' Calc the cost for the weekend.  If the user is checking in
   '''' past 8am, we must charge him.
   '''' </summary>
   'Private Sub CalcWeekEnd(ByRef dt As DataTable, ByVal i As Integer)

   '   Try
   '      Dim dr As DataRow = dt.Rows(i)

   '      If Today.DayOfWeek <> DayOfWeek.Monday Then
   '         MsgBox("Equipment: " & dr("equip_id") & ", " & dr("equip_name") & Chr(10) & _
   '                  " was rented for a weekend, but is not being checked in on Monday." & Chr(10) & _
   '                  "Rental cost is too complex to compute and must be computed " & Chr(10) & _
   '                  "manually.", MsgBoxStyle.Exclamation)
   '         Exit Sub
   '      End If
   '      If Now.Hour > 9 Then
   '         Dim elapsedhours As Integer
   '         Dim eightAMToday As DateTime = _
   '            CType(Today & " 08:00:00 AM", DateTime)
   '         elapsedhours = DateDiff(DateInterval.Hour, eightAMToday, Now)
   '         Dim price As Decimal = Me.GetPriceForEquip(dt.Rows(i).Item("equip_id"), HOURLY)
   '         AddOrUseExistingRow(HOURLY, i, elapsedhours, price, True)
   '      End If

   '   Catch ex As System.Exception
   '      StructuredErrorHandler(ex)
   '   End Try
   'End Sub



   ''' <summary>
   '''
   ''' </summary>
   'Private Sub CalcHalfDay(ByRef dt As DataTable, ByVal i As Integer)

   '   Try
   '      If elapsedHours > modMain.GraceHoursForHalfDayRent + 4 Then
   '         With dt.Rows(i)
   '            .Item("quantity") = 1
   '            .Item("rental_period") = DAILY
   '            .Item("priceperunit") = Me.GetPriceForEquip(.Item("equip_id"), DAILY)
   '         End With
   '      End If
   '   Catch ex As System.Exception
   '      StructuredErrorHandler(ex)
   '   End Try
   'End Sub



   ''' <summary>
   ''' Add a row to the grid for a half day rate.  Go to db to get the
   ''' half day rate for this equip, if not supplied with call.
   ''' </summary>
   'Private Sub AddOverPeriodItem(ByRef dt As DataTable, ByRef dr As DataRow, Optional ByVal price As Single = 0)
   '   Dim hdRate As Single

   '   Try
   '      Dim desc As String = "Half Day Over Use"
   '      If price = 0 Then
   '         hdRate = GetPriceForEquip(dr.Item("equip_id"), HALF_DAY)
   '      Else
   '         hdRate = price
   '      End If
   '      Dim dRow() As Object = {dr.Item("equip_id"), desc, 1, HALF_DAY, hdRate, Now.ToString, 0, 0}
   '      oCG.AddRowToTable(dt, dRow)
   '   Catch ex As System.Exception
   '      StructuredErrorHandler(ex)
   '   End Try
   'End Sub
   ''' <summary>
   ''' Calc the cost for a weekly rental for current grid row.
   ''' </summary>

   'Private Sub CalcWeekly(ByRef dt As DataTable, ByVal i As Integer)
   '   ' compute elapsed days and any overage hours beyond grace

   '   Try
   '      Dim days As Integer = elapsedHours \ 24
   '      Dim weeks As Integer = days \ 7
   '      Dim overDays As Integer = days - (weeks * 7)
   '      Dim overHours As Integer = elapsedHours - (days * 24)
   '      Dim extraHourCharge As Single = 0
   '      Dim dr As DataRow = dt.Rows(i)

   '      ' update week quantity with actual elapsed weeks
   '      dr.Item("quantity") = weeks

   '      ' charge for extra hours, compute this first as
   '      ' it may turn into an extra day
   '      If overHours > GraceHoursForDayRent Then
   '         If UseHalfDays Then
   '            If overHours <= 4 Then
   '               ' add halfday rate item
   '               Dim hdRate As Single = Me.GetPriceForEquip(dr("equip_id"), HALF_DAY)
   '               Dim dRow() As Object = {dr("equip_id"), "Half Day Overuse", 1, HALF_DAY, hdRate, Now.ToString, 0, 0}
   '               oCG.AddRowToTable(dt, dRow)
   '            Else
   '               ' add extra day
   '               overDays += 1
   '            End If
   '         Else
   '            ' using hourly rates on over time
   '            If overHours <= 3 Then
   '               ' add item at overhours * day rate/8
   '               Dim overRate As Single = dt.Rows(i).Item("priceperunit") / 8 * overHours
   '               Dim dRow() As Object = {dr("equip_id"), "Overuse Hours", overHours, HOURLY, overRate, Now.ToString, 0, 0}
   '               oCG.AddRowToTable(dt, dRow)
   '            Else
   '               ' add extra day 
   '               overDays += 1
   '            End If
   '         End If
   '      End If

   '      ' charge for extra days over a week
   '      If overDays > 0 Then
   '         Dim dayRate As Single = Me.GetPriceForEquip(dr.Item("equip_id"), DAILY)
   '         Dim dRow() As Object = {dr.Item("equip_id"), "Extra Days Usage", overDays, DAILY, dayRate, Now.ToString, 0, 0}
   '         oCG.AddRowToTable(dt, dRow)
   '      End If

   '      ' now check for meter over use
   '      Dim allowedUseHours As Integer = days * 8 + overHours
   '      Me.CheckForMeterOverage(dt, i, allowedUseHours)
   '   Catch ex As System.Exception
   '      StructuredErrorHandler(ex)
   '   End Try
   'End Sub
   ''' <summary>
   ''' Ensure that all parameters are set up for checking in
   ''' equipment.
   ''' </summary>
   ''' <returns>Boolean</returns>
   'Private Function VerifySettingsForCkOut() As Boolean
   '   ' ck to see the correct options are turned on
   '   If frm.optLeftBlankCheck.Checked Or frm.optLeftCardNumber.Checked Then
   '      MsgBox("Can't leave check or card at check in.", MsgBoxStyle.Exclamation)
   '      Return False
   '   End If

   '   If UnFormat(frm.txtDeposit.Text) > 0 Then
   '      Dim sMsg As String
   '      Dim iRV As Integer
   '      sMsg = "Deposit amount is greater than 0.  Are you going" & Chr(10)
   '      sMsg &= "to keep the deposit?" & Chr(10)
   '      sMsg &= "" & Chr(10)
   '      sMsg &= "Click Yes to continue.  Click No to cancel and then" & Chr(10)
   '      sMsg &= "clear the deposit amount to 0 and press Manual " & Chr(10)
   '      sMsg &= "Recalculate button before printing invoice." & Chr(10)
   '      sMsg &= "" & Chr(10)
   '      iRV = MsgBox(sMsg, CType(292, Microsoft.VisualBasic.MsgBoxStyle), "Confirm Keeping Deposit")

   '      If iRV = 6 Then
   '         ' Yes Code goes here
   '         Return True
   '      Else
   '         ' No code goes here
   '         Return False
   '      End If
   '   End If
   '   Return True
   'End Function
   ''' <summary>
   ''' Calc the cost for a daily rental for current grid row.
   ''' </summary>
   'Private Sub CalcDaily(ByRef dt As DataTable, ByVal i As Integer)
   '   ' compute elapsed days and any overage hours beyond grace

   '   Try
   '      Dim days As Integer = elapsedHours \ 24
   '      Dim overHours As Integer = elapsedHours - (days * 24)
   '      Dim extraHourCharge As Single = 0
   '      Dim dr As DataRow = dt.Rows(i)

   '      If days > 0 Then
   '         dt.Rows(i).Item("quantity") = days
   '      End If

   '      If overHours > modMain.GraceHoursForDayRent Then
   '         If UseHalfDays Then
   '            If overHours <= 4 Then
   '               ' add halfday rate item
   '               Dim hdRate As Single = Me.GetPriceForEquip(dr("equip_id"), HALF_DAY)
   '               Dim dRow() As Object = {dr("equip_id"), "Half Day Overuse", 1, HALF_DAY, hdRate, Now.ToString, 0, 0}
   '               oCG.AddRowToTable(dt, dRow)
   '            Else
   '               ' add extra day
   '               dt.Rows(i).Item("quantity") += 1
   '            End If
   '         Else
   '            ' using hourly rates on over time
   '            If overHours <= 3 Then
   '               ' add item at overhours * day rate/8
   '               Dim overRate As Single = dt.Rows(i).Item("priceperunit") / 8 * overHours
   '               Dim dRow() As Object = {dr("equip_id"), "Overuse Hours", overHours, HOURLY, overRate, Now.ToString, 0, 0}
   '               oCG.AddRowToTable(dt, dRow)
   '            Else
   '               ' add extra day 
   '               dt.Rows(i).Item("quantity") += 1
   '            End If
   '         End If
   '      End If

   '      ' now check for meter over use
   '      Dim allowedUseHours As Integer = days * 8 + overHours
   '      Me.CheckForMeterOverage(dt, i, allowedUseHours)
   '   Catch ex As System.Exception
   '      StructuredErrorHandler(ex)
   '   End Try
   'End Sub
#End Region
End Class
