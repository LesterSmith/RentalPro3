Module modAutoCalc
   ' these need to change to configurations settings
   Public OPEN_HOUR As Integer = StoreOpenHour               '8am
   Public HALF_DAY_START_HOUR As Integer = AfterNoonRentBeginHour    '1pm - anything checked out later 
   '                                                         than this but returned by 
   '                                                         opening the next day is 
   '                                                         billed 1/2 day
   Public WEEKEND_START_HOUR As Integer = FridayWeekEndStartHour     '4pm - this is the hour on Friday
   '                                                         which marks the start of
   '                                                         a weekend.  Anything checked 
   '                                                         out after this time but
   '                                                         returned by opening Mon.
   '                                                         is billed 1 day

   ' following not used if boolean is false
   Public CALCULATE_BEST_RATE As Boolean = CalcBestRate
   Public Const HALF_DAY_RATE As Integer = 50
   Public Const DAY_RATE As Integer = 100
   Public Const WEEK_RATE As Integer = 300
   Public Const MONTH_RATE As Integer = 900

   ' used when calc best rate is false
   Public MINIMUM_DAYS_PER_WEEK As Integer = WeeklyBreakDays  '3
   Public MINIMUM_DAYS_PER_MONTH As Integer = MonthlyBreakDays  '17

   ' used if best rate = true
   Public MinimumDaysForWeek As Integer = GetMinimumPeriod(DAY_RATE, WEEK_RATE)
   Public MinimumWeeksForMonth As Integer = GetMinimumPeriod(WEEK_RATE, MONTH_RATE)

   ' these are already config, but must be equated to real modmain constants
   Public DAY_GRACE_HOURS As Integer = GraceHoursForDayRent  '2
   Public HALF_DAY_GRACE_HOURS As Integer = GraceHoursForHalfDayRent '1

   Public Const WEEK_DAYS As Integer = 7
   Public Const MONTH_DAYS As Integer = 28

   Public Const HALF_DAY_HOURS As Integer = 4

   ' AutoCalc Structure
   Public Structure RentalTime
      Dim Months As Integer
      Dim Weeks As Integer
      Dim Days As Integer
      Dim HalfDays As Integer
   End Structure

   Public ElapsedTime As String
   Public ElapsedHours As Integer

   ' Find ratio between two rates
   ' Used only when calculate best rate is true
   Private Function GetMinimumPeriod(ByVal Rate As Integer, ByVal NextRate As Integer) As Integer

      Dim i As Integer = NextRate \ Rate

      If NextRate Mod Rate > 0 Then
         Return i + 1
      Else
         Return i
      End If

   End Function

   ' Adds number of days to a date, ignoring Sundays
   Private Function AddWorkDays(ByVal Start As DateTime, ByVal WorkDays As Integer) As DateTime
      Dim EndDate As DateTime = Start
      Dim i As Integer

      For i = 1 To WorkDays
         Select Case EndDate.DayOfWeek
            Case DayOfWeek.Saturday
               EndDate = EndDate.AddDays(2)
            Case Else
               EndDate = EndDate.AddDays(1)
         End Select
      Next

      Return EndDate

   End Function
   ' Recursive Method until all time is accounted for.
   Public Sub AutoCalc(ByVal CheckOut As DateTime, ByVal CheckIn As DateTime, ByRef TotalTime As RentalTime)

      Dim MinHalfDay As DateTime
      Dim MinDay As DateTime
      Dim MinWeek As DateTime
      Dim MinMonth As DateTime

      Dim DueHalfDay As DateTime
      Dim DueDay As DateTime
      Dim DueWeek As DateTime
      Dim DueMonth As DateTime


      'Calculate the minimum and due dates for Half Day,  Day, Week, and Month

      '*** CALCULATE HALF DAY DATES *************
      Select Case CheckOut.DayOfWeek
         Case DayOfWeek.Monday, _
              DayOfWeek.Tuesday, _
              DayOfWeek.Wednesday, _
              DayOfWeek.Thursday, _
              DayOfWeek.Friday, _
              DayOfWeek.Sunday

            MinHalfDay = CheckOut.AddMilliseconds(1)
            If CheckOut.Hour >= HALF_DAY_START_HOUR Then
               DueHalfDay = CheckOut.Date.AddDays(1).AddHours(OPEN_HOUR + HALF_DAY_GRACE_HOURS)
            Else
               DueHalfDay = CheckOut.AddHours(HALF_DAY_HOURS + HALF_DAY_GRACE_HOURS)
            End If

         Case DayOfWeek.Saturday

            MinHalfDay = CheckOut.AddMilliseconds(1)
            If CheckOut.Hour >= HALF_DAY_START_HOUR Then
               DueHalfDay = CheckOut.AddHours(HALF_DAY_HOURS + HALF_DAY_GRACE_HOURS)
            Else
               DueHalfDay = CheckOut.AddHours(HALF_DAY_HOURS + HALF_DAY_GRACE_HOURS)
            End If

      End Select
      '******************************************

      '*** CALCULATE DAY DATES ******************
      Select Case CheckOut.DayOfWeek
         Case DayOfWeek.Monday, _
              DayOfWeek.Tuesday, _
              DayOfWeek.Wednesday, _
              DayOfWeek.Thursday, _
              DayOfWeek.Sunday

            MinDay = DueHalfDay.AddMilliseconds(1)
            DueDay = CheckOut.AddDays(1).AddHours(DAY_GRACE_HOURS)

         Case DayOfWeek.Friday

            MinDay = DueHalfDay.AddMilliseconds(1)
            If CheckOut.TimeOfDay.Hours >= WEEKEND_START_HOUR Then
               'Count entire weekend as 1 day until Monday morning
               DueDay = CheckOut.Date.AddDays(3).AddHours(OPEN_HOUR + DAY_GRACE_HOURS)
               'Consider the checkout to be Saturday morning
               CheckOut = CheckOut.Date.AddDays(1).AddHours(OPEN_HOUR).AddMilliseconds(-1)
            Else
               DueDay = CheckOut.AddDays(1).AddHours(DAY_GRACE_HOURS)
            End If

         Case DayOfWeek.Saturday

            MinDay = DueHalfDay.AddMilliseconds(1)
            'One day goes to Open Hour on Monday
            DueDay = CheckOut.Date.AddDays(2).AddHours(OPEN_HOUR).AddHours(DAY_GRACE_HOURS)

      End Select
      '******************************************


      '*** CALCULATE WEEK DATES *****************
      MinWeek = AddWorkDays(CheckOut, MinimumDaysForWeek)
      If MinWeek.DayOfWeek = DayOfWeek.Sunday Then
         MinWeek = MinWeek.Date.AddDays(1).AddHours(OPEN_HOUR)
      End If
      DueWeek = CheckOut.AddDays(WEEK_DAYS).AddHours(DAY_GRACE_HOURS)
      If DueWeek.DayOfWeek = DayOfWeek.Sunday Then
         DueWeek = DueWeek.Date.AddDays(1).AddHours(OPEN_HOUR).AddHours(DAY_GRACE_HOURS)
      End If
      '******************************************


      '*** CALCULATE MONTH DATES ****************
      If CALCULATE_BEST_RATE Then
         'Use the calculated best value
         MinMonth = CheckOut.AddDays(MinimumWeeksForMonth * WEEK_DAYS)
      Else
         'Use the specified conversion
         MinMonth = CheckOut.AddDays(MINIMUM_DAYS_PER_MONTH)
      End If

      'If min date evaluates to a Sunday, bump to Monday morning
      If MinMonth.DayOfWeek = DayOfWeek.Sunday Then
         MinMonth = MinMonth.Date.AddDays(1).AddHours(OPEN_HOUR)
      End If

      DueMonth = CheckOut.AddDays(MONTH_DAYS).AddHours(DAY_GRACE_HOURS)

      'If due date evaluates to a Sunday, bump to Monday morning
      If DueMonth.DayOfWeek = DayOfWeek.Sunday Then
         DueMonth = DueMonth.Date.AddDays(1).AddHours(OPEN_HOUR).AddHours(DAY_GRACE_HOURS)
      End If
      '******************************************



      'Compare the CheckIn date against the min/due dates to determine period
      Select Case True
         Case DueMonth.CompareTo(CheckIn) < 0
            'Greater than 1 month - Add and then calculate the leftovers
            TotalTime.Months += 1
            AutoCalc(CheckOut.AddDays(MONTH_DAYS), CheckIn, TotalTime)

         Case MinMonth.CompareTo(CheckIn) <= 0
            'Exactly 1 month
            TotalTime.Months += 1

         Case DueWeek.CompareTo(CheckIn) < 0
            'Greater than 1 week - Add and then calculate the leftovers
            TotalTime.Weeks += 1
            AutoCalc(CheckOut.AddDays(WEEK_DAYS), CheckIn, TotalTime)

         Case MinWeek.CompareTo(CheckIn) <= 0
            'Exactly 1 week
            TotalTime.Weeks += 1

         Case DueDay.CompareTo(CheckIn) < 0
            'Greater than 1 day - Add and then calculate the leftovers
            TotalTime.Days += 1

            'If this is on Saturday, then bump the day ahead for the weekend
            If CheckOut.DayOfWeek = DayOfWeek.Saturday Then
               CheckOut = CheckOut.AddDays(1)
            End If

            AutoCalc(CheckOut.AddDays(1), CheckIn, TotalTime)

         Case MinDay.CompareTo(CheckIn) <= 0
            'Exactly 1 day
            TotalTime.Days += 1

         Case CheckOut.CompareTo(CheckIn) <> 0
            'Charge a half day if there is any difference between the two dates
            TotalTime.HalfDays = 1

      End Select
   End Sub

End Module
