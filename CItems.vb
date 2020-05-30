'****************************************
'* Purpose:
'*
'* Author:  Les Smith
'* Date Created: 05/15/2003 at 09:10:11
'* CopyRight:  HHI Software
'****************************************
'*
Option Strict Off
Option Explicit On 
Public Class CItems

#Region " Class Level Variables "
   Private FItemID As String 'local copy
   Private FItemName As String 'local copy
   Private FRentalPeriod As String ' 
   Private FRentalTime As String 'local copy
   Private FPrice As Decimal 'local copy
   Private FItemTotal As Decimal 'local copy
   Private FDeposit As Decimal 'local copy
   Private FDelivery As Decimal 'local copy
   Private FRentOrSell As String 'local copy
   ' Private FExpectedReturn As Date 'local copy
   'Private FHourly As Decimal 'local copy
   Private FHalfDay As Decimal 'local copy
   Private FDaily As Decimal 'local copy
   Private FWeekly As Decimal 'local copy
   Private FMonthly As Decimal 'local copy
   Private FWeekEnd As Decimal
   Private FMinimum As Short 'local copy


#End Region


#Region " Constructor "
   Public Sub New(ByVal ItemID As String, _
      ByVal ItemName As String, _
      ByVal RentalPeriod As String, _
      ByVal Price As Decimal, _
      ByVal ItemTotal As Decimal, _
      ByVal Deposit As Decimal, _
      ByVal Delivery As Decimal, _
      ByVal Rentorsell As String, _
      ByVal HalfDay As Decimal, _
      ByVal Daily As Decimal, _
      ByVal Weekly As Decimal, _
      ByVal Monthly As Decimal, _
      ByVal WeekEnd As Decimal, _
      ByVal Minimum As Short)
      FItemID = ItemID
      FItemName = ItemName
      FRentalPeriod = RentalPeriod
      FPrice = Price
      FItemTotal = ItemTotal
      FDeposit = Deposit
      FDelivery = Delivery
      FRentOrSell = Rentorsell
      FHalfDay = HalfDay
      FDaily = Daily
      FWeekly = Weekly
      FMonthly = Monthly
      FWeekEnd = WeekEnd
      FMinimum = Minimum
   End Sub

#End Region

#Region " Public Properties "
   Public Property ItemID() As String
      Get
         ItemID = FItemID
      End Get
      Set(ByVal Value As String)
         FItemID = Value
      End Set
   End Property
   Public Property ItemName() As String
      Get
         ItemName = FItemName
      End Get
      Set(ByVal Value As String)
         FItemName = Value
      End Set
   End Property
   Public Property RentalPeriod() As String
      Get
         RentalPeriod = FRentalPeriod
      End Get
      Set(ByVal Value As String)
         FRentalPeriod = Value
      End Set
   End Property
   Public Property Price() As Decimal
      Get
         Price = FPrice
      End Get
      Set(ByVal Value As Decimal)
         FPrice = Value
      End Set
   End Property
   Public Property ItemTotal() As Decimal
      Get
         ItemTotal = FItemTotal
      End Get
      Set(ByVal Value As Decimal)
         FItemTotal = Value
      End Set
   End Property
   Public Property Deposit() As Decimal
      Get
         Deposit = FDeposit
      End Get
      Set(ByVal Value As Decimal)
         FDeposit = Value
      End Set
   End Property
   Public Property Delivery() As Decimal
      Get
         Delivery = FDelivery
      End Get
      Set(ByVal Value As Decimal)
         FDelivery = Value
      End Set
   End Property
   Public Property RentOrSell() As String
      Get
         RentOrSell = FRentOrSell
      End Get
      Set(ByVal Value As String)
         FRentOrSell = Value
      End Set
   End Property
   Public Property RentalTime() As String
      Get
         RentalTime = FRentalTime
      End Get
      Set(ByVal Value As String)
         FRentalTime = Value
      End Set
   End Property
   Public Property HalfDay() As Decimal
      Get
         HalfDay = FHalfDay
      End Get
      Set(ByVal Value As Decimal)
         FHalfDay = Value
      End Set
   End Property
   Public Property Daily() As Decimal
      Get
         Daily = FDaily
      End Get
      Set(ByVal Value As Decimal)
         FDaily = Value
      End Set
   End Property
   Public Property Weekly() As Decimal
      Get
         Weekly = FWeekly
      End Get
      Set(ByVal Value As Decimal)
         FWeekly = Value
      End Set
   End Property
   Public Property Monthly() As Decimal
      Get
         Monthly = FMonthly
      End Get
      Set(ByVal Value As Decimal)
         FMonthly = Value
      End Set
   End Property



   Public Property WeekEnd() As Decimal
      Get
         WeekEnd = FWeekEnd
      End Get
      Set(ByVal Value As Decimal)
         FWeekEnd = Value
      End Set
   End Property
   Public Property Minimum() As Short
      Get
         Minimum = FMinimum
      End Get
      Set(ByVal Value As Short)
         FMinimum = Value
      End Set
   End Property

   'Public Property Hourly() As Decimal
   '   Get
   '      Hourly = FHourly
   '   End Get
   '   Set(ByVal Value As Decimal)
   '      FHourly = Value
   '   End Set
   'End Property
   'Public Property ExpectedReturn() As Date
   '   Get
   '      ExpectedReturn = FExpectedReturn
   '   End Get
   '   Set(ByVal Value As Date)
   '      FExpectedReturn = Value
   '   End Set
   'End Property


#End Region



End Class