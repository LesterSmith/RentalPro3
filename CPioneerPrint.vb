'****************************************
'* Purpose: This class supports printing
'* for a program.  It accepts a string
'* and prints it to a report with headings
'* and page numbers etc.  It will word wrap
'* a memo string.
'*
'* Author:  Les Smith
'* Date Created: 06/06/2002 at 11:15:08
'* CopyRight:  HHI Software, Inc.
'****************************************
Imports System
Imports System.Drawing
Imports System.Drawing.Text
Imports System.IO
Imports System.Text
Imports System.Windows.Forms
Public Class CPioneerPrint
#Region " Class Level Variables "
   '` The following methods are exposed by this
   '' object.
   '` 1) InitializeReport - Sets up font for detail
   '' printing.
   '` 2) PrintString - Call to print a string object of
   '' multiple lines.
   '` Internally, a PrintDocument object will be
   '` instantiated, along with a PrintPage Event.
   '` Calling the .Print method of the PrintDocument
   '` instance will automatically fire the PrintPage
   '` event to get the next page to print.  In VB6,
   '` printing was done linearly.  In VB.NET, it is
   '` done through a callback methodology.  The user
   '` initiates the printing object and the object
   '' calls the PrintPage event to get a page to print.
   ''
   Private ColorCoded As Boolean = True
   Private WithEvents PrintDoc As Printing.PrintDocument
    Private msRptString As StringBuilder ' holds the report string
   Private miNL As Integer ' number of lines in the report
   Private miNL2 As Integer ' number of sub lines in line
   Private mI As Integer
   Private mI2 As Integer
   Private CurrentLine As Integer ' curr print line on a page
   Private miChrPerLine As Integer ' nbr chars to print per line
   Public Title As String
   Private PageNbr As Integer
   Private Portrait As Boolean = True ' landscape if false
   Public SubTitle As String
   Public Heading As String
   Private DetailFontSize As Single
   Dim oUtil As CUtilities
   Dim oUtil2 As CUtilities
   Dim msLine As String
   Dim msToken As String
   Dim sFooter As String
   Dim miFileType As Integer
   Dim miWordWrap As Integer = 0
    Const DETAIL_FONT As String = "Courier New"
    Const DETAIL_FONT_SIZE_80 As Integer = 10
    Const DETAIL_FONT_SIZE_96 As Integer = 9
    Const DETAIL_FONT_SIZE_120 As Integer = 8
    'Const DETAIL_FONT_SIZE_132 As Integer  = 7
    Const DETAIL_FONT_SIZE_160 As Integer = 6
    Const DETAIL_FONT_BOLD As Boolean = True
   Private mbSepLines As Boolean
    Const CUST_BILL_FONT_SIZE As Integer = 10
    Const JOB_SITE_FONT_SIZE As Integer = 10
    Const INVOICE_DETAIL_FONT_SIXE As Integer = 9
    Const INVOICE_HDR_DATA_FONT_SIZE As Integer = 8
   'Private F As frmCustomers
   'Private FIn As frmCheckIn
   Private InvID As Integer
   ' following values come from either frmCustomer or frmCheckin
   Private m_ShipToName As String
   Private m_ShipToAddress1 As String
   Private m_ShipToCity As String
   Private m_ShipToState As String
   Private m_ShipToZip As String
   Private m_ComapanyName As String
   Private m_BillAddress1 As String
   Private m_BillingCity As String
   Private m_BillingState As String
   Private m_BillingZip As String
   Private m_InvoiceType As String
   Private m_CustomerID As String
   Private m_PONbr As String
   Private m_ContactName As String
   Private m_CheckNumber As String
   Private m_PaidOption As String
   Private m_InvoiceId As String
   Private m_InvoiceDate As DateTime
   Private m_TaxId As String
   Private _CheckOutEmployee As String


#End Region

#Region " Public Methods "
   Public Function HideCCNumber(ByVal ccn As String) As String
      Dim len As Short = ccn.Length
      Dim i As Integer
      Dim s As String = ccn
      If len > 4 Then
         For i = 1 To len - 4
            Mid(s, i, 1) = "X"
         Next
         Return s
      Else
         Return ccn
      End If
   End Function

   Public Function PadR(ByVal s As String, ByVal n As Integer) As String
      If s.Trim.Length > n - 1 Then
         Return s.Substring(0, n)
      Else
         Return s.Trim & Space(n - Len(s.Trim))
      End If
   End Function

   Public Function PadL(ByVal s As String, ByVal n As Integer) As String
      If s.Trim.Length > n - 1 Then
         Return s.Substring(0, n)
      Else
         Return Space(n - s.Trim.Length) & s.Trim
      End If
   End Function


    Friend Overloads Sub StartPrint(ByRef ps As StringBuilder, ByVal InvoiceID As Integer)


        '   Initializes a report
        '
        '   Input:  iWordWrap       I   - word wrapping column
        '           title           S   - title text for report
        '           subTitle        S   - subtitle text for report
        ' All we do here is save the parameters passed
        ' as we have no print object to work with here,
        ' the PagePrint will access the saved parameter
        ' data and set the printdoc parameters as each
        ' page is printed


        ' create two objects so that we can use
        ' nested calls to MemoLine w/o stepping
        ' on each other...


        Try
            Me.InvID = InvoiceID
            Me.msRptString = ps
            oUtil = New CUtilities()
            Me.miNL = oUtil.MLCount(ps.ToString, 0)
            Me.mI = 1
            InvID = InvoiceID

            ' set up the printdocument object
            PrintDoc = New Printing.PrintDocument()
            PrintDoc.Print() ' kick off the printing


        Catch ex As Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

    Friend Sub PrintPreview(ByRef ps As StringBuilder, ByVal InvoiceID As Integer)
        Dim previewDialog As New PrintPreviewDialog()

        Try
            Me.msRptString = ps
            Me.InvID = InvoiceID
            oUtil = New CUtilities()
            Me.miNL = oUtil.MLCount(ps.ToString, 0)
            Me.mI = 1
            PrintDoc = New Printing.PrintDocument()
            PrintDoc.DocumentName = "Pioneer Print"
            previewDialog.Document = PrintDoc
            previewDialog.ShowDialog()
            PrintDoc.Dispose()
            previewDialog.Dispose()

        Catch ex As Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

#End Region

#Region " Private Methods "
    Private Sub PrtDoc_PrintPage(ByVal sender As Object, _
       ByVal e As Printing.PrintPageEventArgs) _
       Handles PrintDoc.PrintPage

        '` This method handles the callback from the PrtDoc
        '` object.  It supplies the line(s) to be printed by
        '` calling the DrawString method of the events
        '` PrintPageEventArgs parameter, supplied to the
        '' event.
        '-------------------------------------------------------'
        ' method to determine how wide a sub string is
        ' Dim sz As Drawing.SizeF
        ' sz = e.Graphics.MeasureString("This string", PrintFont)
        ' h = sz.Height
        ' Dim w As Integer = sz.Width
        ' xPos += w
        '-------------------------------------------------------'
        Dim LineHeight As Single
        Dim LineWidth As Single
        Dim yPos As Single
        Dim xPos As Single
        Dim PrintFont As Font
        Dim yPosSave As Single

        Try
            Static i As Integer
            Static nL As Integer
            Static sLine As String
            Dim br As Brushes
            Dim o As Object
            Static Quotes As Integer
            Static iChrB4NonSpace As Integer
            Dim sHdrLine As String
            Dim CharWidth As Single
            Dim chartest As String = " "
            Dim sz As Drawing.SizeF
            Dim TextWidth As Single
            Dim sPageNbr As String
            Dim bDoneWithString As Boolean
            Dim x1 As Single
            Dim y1 As Single
            Dim x2 As Single
            Dim y2 As Single
            Dim drPen As New Pen(Color.Black, 1)
            'Dim drPen2 As New Pen(Color.Black, 1)
            Dim LeftMargin As Single
            Dim RightMargin As Single
            Dim TopMargin As Single
            Dim BottomMargin As Single
            Dim iLineLen As Integer = miChrPerLine
            'Static PageNbr As Integer

            ' Compute the margins after determining
            ' landscape or portrait
            PrintFont = New Font(DETAIL_FONT, 10, FontStyle.Bold)
            sz = e.Graphics.MeasureString("m", PrintFont)
            LeftMargin = e.MarginBounds.Left - (3 * sz.Width)
            RightMargin = e.MarginBounds.Right '+ (6 * sz.Width)

            sz = e.Graphics.MeasureString("M", PrintFont)
            TopMargin = e.MarginBounds.Top + (2 * sz.Height)
            BottomMargin = e.MarginBounds.Bottom + (10 * sz.Height)

            xPos = LeftMargin 'e.MarginBounds.Left
            yPos = TopMargin 'e.MarginBounds.Top

            ' Areas to Print:
            ' 1) Job site 10 point
            ' 2) Customer Billing info
            ' 3)                                  Invoice info
            ' 4) Invoice detail 9 point
            ' Print the header on this page first thing
            ' First, print the title
            PageNbr += 1
            PrintFont = New Font("Courier New", 10, FontStyle.Bold)
            LineHeight = PrintFont.GetHeight(e.Graphics)
            'sz = e.Graphics.MeasureString(Title, PrintFont)
            'TextWidth = sz.Width
            'LineWidth = e.MarginBounds.Right + e.MarginBounds.Left
            LineWidth = RightMargin + LeftMargin


            ' *************print the job site company name*************
            e.Graphics.DrawString(m_ShipToName, _
                                    PrintFont, _
                                    Brushes.Black, _
                                    xPos, _
                                    yPos, _
                                    New StringFormat())
            yPos += LineHeight

            ' print the job site Address1
            e.Graphics.DrawString(m_ShipToAddress1, _
                                    PrintFont, _
                                    Brushes.Black, _
                                    xPos, _
                                    yPos, _
                                    New StringFormat())
            yPos += LineHeight

            ' print the job site city,state zip
            Dim sCSZ As String = m_ShipToCity & ", " & _
                               m_ShipToState & "  " & _
                               m_ShipToZip
            e.Graphics.DrawString(sCSZ, _
                                    PrintFont, _
                                    Brushes.Black, _
                                    xPos, _
                                    yPos, _
                                    New StringFormat())
            ' skip 3 lines after the Job site
            yPos += LineHeight * 4



            ' ************print the company name****************
            e.Graphics.DrawString(m_ComapanyName, _
                                    PrintFont, _
                                    Brushes.Black, _
                                    xPos, _
                                    yPos, _
                                    New StringFormat())
            yPos += LineHeight

            ' print the Address1
            e.Graphics.DrawString(m_BillAddress1, _
                                    PrintFont, _
                                    Brushes.Black, _
                                    xPos, _
                                    yPos, _
                                    New StringFormat())
            yPos += LineHeight

            ' print the city,state zip
            sCSZ = m_BillingCity & ", " & _
                  m_BillingState & "  " & _
                  m_BillingZip
            e.Graphics.DrawString(sCSZ, _
                                    PrintFont, _
                                    Brushes.Black, _
                                    xPos, _
                                    yPos, _
                                    New StringFormat())
            ' skip 3 lines after the customer data
            yPos += LineHeight * 3

            ' **************print the invoice header data***********
            xPos = ((LineWidth - TextWidth) / 2) + (9 * sz.Width)
            yPosSave = yPos ' save for printing detail
            yPos = TopMargin
            PrintFont = New Font(DETAIL_FONT, 9, FontStyle.Bold)


            sCSZ = "Invoice Type: " & m_InvoiceType
            e.Graphics.DrawString(sCSZ, _
                                    PrintFont, _
                                    Brushes.Black, _
                                    xPos, _
                                    yPos, _
                                    New StringFormat())
            yPos += LineHeight

            ' print the date
            sCSZ = "Inv Date: " & m_InvoiceDate.ToString
            e.Graphics.DrawString(sCSZ, _
                                    PrintFont, _
                                    Brushes.Black, _
                                    xPos, _
                                    yPos, _
                                    New StringFormat())
            yPos += LineHeight

            ' print the date printed if a reprint
            If m_InvoiceType = "Reprint" Then
                sCSZ = "Printed: " & Now.ToString
                e.Graphics.DrawString(sCSZ, _
                                        PrintFont, _
                                        Brushes.Black, _
                                        xPos, _
                                        yPos, _
                                        New StringFormat())
                yPos += LineHeight
            End If

            sCSZ = "Customer #:  " & Format(Val(m_CustomerID), "0")
            e.Graphics.DrawString(sCSZ, _
                                    PrintFont, _
                                    Brushes.Black, _
                                    xPos, _
                                    yPos, _
                                    New StringFormat())
            yPos += LineHeight

            sCSZ = "Invoice #: " & Format(Val(InvID), "0")
            e.Graphics.DrawString(sCSZ, _
                                    PrintFont, _
                                    Brushes.Black, _
                                    xPos, _
                                    yPos, _
                                    New StringFormat())
            yPos += LineHeight
            sCSZ = "P. O. #: " & m_PONbr
            e.Graphics.DrawString(sCSZ, _
                                    PrintFont, _
                                    Brushes.Black, _
                                    xPos, _
                                    yPos, _
                                    New StringFormat())
            yPos += LineHeight

            sCSZ = "Contact: " & m_ContactName
            e.Graphics.DrawString(sCSZ, _
                                    PrintFont, _
                                    Brushes.Black, _
                                    xPos, _
                                    yPos, _
                                    New StringFormat())
            yPos += LineHeight

            If m_TaxId.Trim.Length > 0 Then
                sCSZ = "Tax ID: " & m_TaxId
                e.Graphics.DrawString(sCSZ, _
                                        PrintFont, _
                                        Brushes.Black, _
                                        xPos, _
                                        yPos, _
                                        New StringFormat())
                yPos += LineHeight
            End If

            'print paid option
            If m_PaidOption = "BC" Then
                sCSZ = "Blank Check #:  " & m_CheckNumber
            ElseIf m_PaidOption = "LC" Then
                sCSZ = "Left Card #:  " & m_CheckNumber
            ElseIf m_PaidOption = "CK" Then
                sCSZ = "Paid by Check #:  " & m_CheckNumber
            ElseIf m_PaidOption = "CC" Then
                sCSZ = "Paid by Card #:  " & m_CheckNumber
            ElseIf m_PaidOption = "BT" Then
                sCSZ = "Bill To #:  " & m_CustomerID.ToString 'Format(Val(m_CustomerID), "000000")
            ElseIf m_PaidOption = "CA" Then
                sCSZ = "Paid by Cash"
            End If
            e.Graphics.DrawString(sCSZ, _
                                    PrintFont, _
                                    Brushes.Black, _
                                    xPos, _
                                    yPos, _
                                    New StringFormat())
            yPos += LineHeight

            ' print the employee checking in/out
            If m_InvoiceType = "Rental Check Out" Then
                sCSZ = "CkOut Emp: " & Me._CheckOutEmployee
            ElseIf m_InvoiceType = "Rental Check In" Then
                sCSZ = "CkIn Emp: " & Me._CheckOutEmployee
            Else
                sCSZ = String.Empty
            End If
            If sCSZ.Length > 0 Then
                e.Graphics.DrawString(sCSZ, _
                                        PrintFont, _
                                        Brushes.Black, _
                                        xPos, _
                                        yPos, _
                                        New StringFormat())
                yPos += LineHeight * 2
            Else
                yPos += LineHeight
            End If


            ' set up font for detail of invoice
            yPos = yPosSave
            xPos = LeftMargin
            yPos += (4 * LineHeight)

            ' next print the page number on the right end
            ' of the first line
            'sPageNbr = PageNbr.ToString
            'xPos = RightMargin - (20 * sPageNbr.Length)

            'e.Graphics.DrawString(sPageNbr, _
            '                        PrintFont, _
            '                        Brushes.Black, _
            '                        xPos, _
            '                        yPos, _
            '                        New StringFormat())
            'yPos += LineHeight

            ' here we are printing simple, non color coded lines
            ' so we don't need to tokenize the line.
            ' we can just call memoline limiting the lines to
            ' a length that will fit on a line of print

            ' if we have a blank line, don't print a sep line
            ' after it...
            Dim bBlank As Boolean

            For i = mI To miNL
                ' just get a line and print it and then
                ' check to see if there is enough room to print another
                ' line.
                sLine = oUtil.MemoLine(msRptString.ToString, 0, i)

                If (sLine Is Nothing OrElse sLine.Length = 0) Then
                    ' dont print a blank line, just bump yPos
                    bBlank = True
                Else
                    LineHeight = PrintFont.GetHeight(e.Graphics)
                    xPos = LeftMargin 'e.MarginBounds.Left
                    bBlank = False
                End If

                PrintFont = New Font(DETAIL_FONT, 9, FontStyle.Bold)
                e.Graphics.DrawString(sLine, _
                                     PrintFont, _
                                     Brushes.Black, _
                                     xPos, _
                                     yPos, _
                                     New StringFormat())
                yPos += LineHeight


                ' check to see if we are at the end of the page
                If yPos >= (BottomMargin - LineHeight) Then
                    ' end of page, ck for more lines to print
                    ' after print the footer
                    ' put a little space before the line
                    yPos += 2
                    x1 = LeftMargin 'e.MarginBounds.Left
                    y1 = yPos
                    x2 = RightMargin 'e.MarginBounds.Right
                    y2 = yPos
                    e.Graphics.DrawLine(drPen, x1, y1, x2, y2)
                    yPos += 4

                    PrintFont.Dispose()
                    PrintFont = New Font("Arial", 10)
                    xPos = LeftMargin 'e.MarginBounds.Left
                    e.Graphics.DrawString(sFooter, _
                       PrintFont, _
                       Brushes.Black, _
                       xPos, _
                       yPos, _
                       New StringFormat())

                    ' ck for more lines to print
                    If i < miNL Then
                        e.HasMorePages = True
                        ' set ptr to next line back
                        mI = i + 1
                        Exit Sub
                    Else
                        e.HasMorePages = False
                        PageNbr = 0 ' in case called again from preview print
                        Exit Sub
                    End If
                End If
            Next

            e.HasMorePages = False
            PageNbr = 0 ' in case called again from preview print
            'print a footer on the last page
            yPos = BottomMargin - LineHeight 'e.MarginBounds.Bottom - LineHeight
            x1 = LeftMargin 'e.MarginBounds.Left
            y1 = yPos
            x2 = RightMargin 'e.MarginBounds.Right
            y2 = yPos
            e.Graphics.DrawLine(drPen, x1, y1, x2, y2)
            PrintFont.Dispose()
            PrintFont = New Font("Arial", 10)
            xPos = LeftMargin 'e.MarginBounds.Left
            e.Graphics.DrawString(sFooter, _
               PrintFont, _
               Brushes.Black, _
               xPos, _
               yPos, _
               New StringFormat())
        Catch ex As Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

#End Region

#Region " Constructor "
   Public Sub New()
      ' Caller will set properties prior to calling PrintSelectedInvoices
   End Sub

    '#If Reliable Then
   Public Sub New(ByRef f As frmCheckinNew)
      '#Else
      '   Public Sub New(byref f as frmCheckin)
      '#End If
      m_ShipToName = f.txtShipToCustomer.Text.Trim
      m_ShipToAddress1 = f.txtShipAddress1.Text.Trim
      m_ShipToCity = f.txtShipCity.Text
      m_ShipToState = f.txtShipState.Text
      m_ShipToZip = f.txtShipZip.Text
      m_ComapanyName = f.txtCompanyName.Text.Trim
      m_BillAddress1 = f.txtBillingAddress1.Text.Trim
      m_BillingCity = f.txtCity.Text
      m_BillingState = f.txtState.Text
      m_BillingZip = f.txtPostalCode.Text
      m_InvoiceType = "Rental Check In"
      m_CustomerID = f.txtCustomerID.Text
      m_CustomerID = f.txtCustomerID.Text
      m_PONbr = f.txtPONbr.Text
      m_ContactName = f.txtContactName.Text
      m_CheckNumber = f.txtCheckNumber.Text
      m_TaxId = f.txtTaxID.Text
      If f.optBillTo.Checked Then
         m_PaidOption = "BT"
      ElseIf f.optCash.Checked Then
         m_PaidOption = "CA"
      ElseIf f.optLeftBlankCheck.Checked Then
         m_PaidOption = "BC"
      ElseIf f.optLeftCardNumber.Checked Then
         m_PaidOption = "LC"
         m_CheckNumber = HideCCNumber(m_CheckNumber)
      ElseIf f.optPaidByCheck.Checked Then
         m_PaidOption = "CK"
      ElseIf f.optPaidByCreditCard.Checked Then
         m_PaidOption = "CC"
         m_CheckNumber = HideCCNumber(m_CheckNumber)
      End If
      m_InvoiceDate = Now
      Me._CheckOutEmployee = f.CheckOutEmployee
   End Sub

   Public Sub New(ByRef f As frmCustomers)

      Try
         m_ShipToName = f.txtShipToCustomer.Text.Trim
         m_ShipToAddress1 = f.txtShipAddress1.Text.Trim
         m_ShipToCity = f.txtShipCity.Text
         m_ShipToState = f.txtShipState.Text
         m_ShipToZip = f.txtShipZip.Text
         m_ComapanyName = f.txtCompanyName.Text.Trim
         m_BillAddress1 = f.txtBillingAddress1.Text.Trim
         m_BillingCity = f.txtCity.Text
         m_BillingState = f.txtState.Text
         m_BillingZip = f.txtPostalCode.Text
         If f.chkCkOutAndIN.Checked Then
            m_InvoiceType = "Check Out & In"
         Else
            m_InvoiceType = "Rental Check Out"
         End If
         m_CustomerID = f.txtCustomerID.Text
         m_CustomerID = f.txtCustomerID.Text
         m_PONbr = f.txtPONbr.Text
         m_ContactName = f.txtContactName.Text
         m_CheckNumber = f.txtCheckNumber.Text
         If f.optBillTo.Checked Then
            m_PaidOption = "BT"
         ElseIf f.optCash.Checked Then
            m_PaidOption = "CA"
         ElseIf f.optLeftBlankCheck.Checked Then
            m_PaidOption = "BC"
         ElseIf f.optLeftCardNumber.Checked Then
            m_PaidOption = "LC"
            m_CheckNumber = HideCCNumber(m_CheckNumber)
         ElseIf f.optPaidByCheck.Checked Then
            m_PaidOption = "CK"
         ElseIf f.optPaidByCreditCard.Checked Then
            m_PaidOption = "CC"
            m_CheckNumber = HideCCNumber(m_CheckNumber)
         End If
         m_InvoiceDate = Now
         m_TaxId = f.txtTaxID.Text
         Me._CheckOutEmployee = f.CheckOutEmployee
        Catch ex As Exception
            StructuredErrorHandler(ex)
      End Try
   End Sub


#End Region

#Region " Public Properties "
   Public Property ShipToName() As String
      Get
         Return m_ShipToName
      End Get
      Set(ByVal Value As String)
         m_ShipToName = Value
      End Set
   End Property

   Public Property ShipToAddress1() As String
      Get
         Return m_ShipToAddress1
      End Get
      Set(ByVal Value As String)
         m_ShipToAddress1 = Value
      End Set
   End Property

   Public Property ShipToCity() As String
      Get
         Return m_ShipToCity
      End Get
      Set(ByVal Value As String)
         m_ShipToCity = Value
      End Set
   End Property

   Public Property ShipToState() As String
      Get
         Return m_ShipToState
      End Get
      Set(ByVal Value As String)
         m_ShipToState = Value
      End Set
   End Property

   Public Property ShipToZip() As String
      Get
         Return m_ShipToZip
      End Get
      Set(ByVal Value As String)
         m_ShipToZip = Value
      End Set
   End Property

   Public Property ComapanyName() As String
      Get
         Return m_ComapanyName
      End Get
      Set(ByVal Value As String)
         m_ComapanyName = Value
      End Set
   End Property

   Public Property BillAddress1() As String
      Get
         Return m_BillAddress1
      End Get
      Set(ByVal Value As String)
         m_BillAddress1 = Value
      End Set
   End Property

   Public Property BillingCity() As String
      Get
         Return m_BillingCity
      End Get
      Set(ByVal Value As String)
         m_BillingCity = Value
      End Set
   End Property

   Public Property BillingState() As String
      Get
         Return m_BillingState
      End Get
      Set(ByVal Value As String)
         m_BillingState = Value
      End Set
   End Property

   Public Property BillingZip() As String
      Get
         Return m_BillingZip
      End Get
      Set(ByVal Value As String)
         m_BillingZip = Value
      End Set
   End Property

   Public Property InvoiceType() As String
      Get
         Return m_InvoiceType
      End Get
      Set(ByVal Value As String)
         m_InvoiceType = Value
      End Set
   End Property

   Public Property CustomerID() As String
      Get
         Return m_CustomerID
      End Get
      Set(ByVal Value As String)
         m_CustomerID = Value
      End Set
   End Property

   Public Property PONbr() As String
      Get
         Return m_PONbr
      End Get
      Set(ByVal Value As String)
         m_PONbr = Value
      End Set
   End Property

   Public Property ContactName() As String
      Get
         Return m_ContactName
      End Get
      Set(ByVal Value As String)
         m_ContactName = Value
      End Set
   End Property

   Public Property CheckNumber() As String
      Get
         Return m_CheckNumber
      End Get
      Set(ByVal Value As String)
         m_CheckNumber = Value
      End Set
   End Property

   Public Property PaidOption() As String
      Get
         Return m_PaidOption
      End Get
      Set(ByVal Value As String)
         m_PaidOption = Value
      End Set
   End Property

   Public Property InvoiceId() As String
      Get
         Return m_InvoiceId
      End Get
      Set(ByVal Value As String)
         m_InvoiceId = Value
      End Set
   End Property

   Public Property InvoiceDate() As DateTime
      Get
         Return m_InvoiceDate
      End Get
      Set(ByVal Value As DateTime)
         m_InvoiceDate = Value
      End Set
   End Property

   Public Property TaxId() As String
      Get
         Return m_TaxId
      End Get
      Set(ByVal Value As String)
         m_TaxId = Value
      End Set
   End Property

   Public Property CheckOutEmployee() As String
      Get
         Return _CheckOutEmployee
      End Get
      Set(ByVal Value As String)
         _CheckOutEmployee = Value
      End Set
   End Property


#End Region

End Class
