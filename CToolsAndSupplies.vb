Public Class CToolsAndSupplies
   Private m_ItemCount As Integer
   Private m_ItemId As String
   Private m_ItemPrice As Decimal
   Private m_ItemExtendedPrice As Decimal
   Private m_RentOrSale As String = "Sale"
   Private m_Deposit As Decimal = 0
   Private m_ItemName As String

    Public Property ItemCount As Integer
        Get
            Return m_ItemCount
        End Get
        Set(ByVal Value As Integer)
           m_ItemCount = Value
        End Set
    End Property

    Public Property ItemId As String
        Get
            Return m_ItemId
        End Get
        Set(ByVal Value As String)
           m_ItemId = Value
        End Set
    End Property

    Public Property ItemPrice As Decimal
        Get
            Return m_ItemPrice
        End Get
        Set(ByVal Value As Decimal)
           m_ItemPrice = Value
        End Set
    End Property

    Public Property ItemExtendedPrice As Decimal
        Get
            Return m_ItemExtendedPrice
        End Get
        Set(ByVal Value As Decimal)
           m_ItemExtendedPrice = Value
        End Set
    End Property

    Public Property RentOrSale As String
        Get
            Return m_RentOrSale
        End Get
        Set(ByVal Value As String)
           m_RentOrSale = Value
        End Set
    End Property

    Public Property Deposit As Decimal
        Get
            Return m_Deposit
        End Get
        Set(ByVal Value As Decimal)
           m_Deposit = Value
        End Set
    End Property

    Public Property ItemName As String
        Get
            Return m_ItemName
        End Get
        Set(ByVal Value As String)
           m_ItemName = Value
        End Set
    End Property
End Class
