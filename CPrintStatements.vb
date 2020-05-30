Imports System.Text
Imports System
Imports System.Drawing
Imports System.Drawing.Text
Imports System.IO
Imports System.Windows.Forms
Public Class CPrintStatements
#Region " Class Level Variables "
   Private oDA As CDataAccess
   Private SQL As String
   Private sb As StringBuilder
   Private sb2 As StringBuilder
   Private oUtil As CUtilities
   Private m_Preview As Boolean
   Private m_CustomerID As String = String.Empty


#End Region


#Region " Public Methods "
   Public Sub PrintStatements()
      ' this method will get a recordset of unpaid invoices,
      ' joined with the customer billing address
      ' it will print each customers statment on a different page
      Dim iLastCustomer As Integer = 0
      Dim dt As New DataTable()
      Dim i As Integer

      Try
         Dim decInvoiceTotal As Decimal = 0
         Dim s As String
         Dim bNewCustomer As Boolean = True
         Dim oPS As PrintStatement
         Const SL = 5

         sb = New StringBuilder()
         sb2 = New StringBuilder()
         SQL = "select I.*,C.Companyname,c.ContactName,c.BillingAddress1, c.BillingAddress2, "
         SQL &= "c.city,c.state,c.postalcode, c.customerid "
         SQL &= "from invoices i, customers c "
         SQL &= "where i.customerid = c.customerid "
         SQL &= "and i.balancedue <> 0 "
         SQL &= "and i.status = 'OPEN' "
         If m_CustomerID.Length > 0 Then
            SQL &= "and c.customerid = " & m_CustomerID & " "
         End If
         SQL &= "order by c.customerid,i.invoiceid"
         Dim iRows As Integer = oDA.SendQuery(SQL, dt, ConnectString)
         If iRows < 1 Then
            MsgBox("There are no statements to print.", MsgBoxStyle.Information)
            Exit Sub
         End If

         ' here's the format:
         ' Bill To:                                   Remit To:
         ' Company Name                               Pioneer Rental, Inc.
         ' Billing Address                            Address
         ' City, St zip                               City, ST zip

         ' ______________________________________________________________________\
         ' 0000000001111111111222222222233333333334444444444555555555566666666667
         ' 1234567890123456789012345678901234567890123456789012345678901234567890
         ' Invoice #     Invoice Date   P.O. Number  Picked Up By: Invoice Amount
         '-----------------------------------------------------------------
         '     1001       10/21/2003    123456       Les Smith       $1,234.56
         '     1005       10/24/2003                                    987.34
         '                                                          ----------
         '                                          Invoice Total:   $n,nnn.nn

         iLastCustomer = dt.Rows(0).Item("c.customerid")
         Dim dr As DataRow

         With dt.Rows(0)
            dr = dt.Rows(0)
            ' print billing customer data
            sb.Append(Space(SL) & "Bill To:".PadRight(43) & "Remit To:" & vbCrLf)
            sb.Append(Space(SL) & CType(.Item("companyname"), String).PadRight(43) & _
               ReportName & vbCrLf)
            sb.Append(Space(SL) & CType(.Item("billingaddress1"), String).PadRight(43) & _
               Address1 & vbCrLf)
            If MNS(dr("billingaddress2")).Length > 0 Then
               sb.Append(Space(SL) & MNS(dr("billingaddress2")).PadRight(43) & Address2 & vbCrLf)
            End If
            sb.Append(Space(SL) & (CType(.Item("City"), String) & ", " & _
                CType(.Item("State"), String) & "  " & _
                CType(.Item("postalcode"), String)).PadRight(43) & _
                City & ", " & State & " " & Zip & vbCrLf)

            sb2.Append(Space(SL) & "Invoice #".PadRight(14) & _
                      "Invoice Date".PadRight(15) & _
                      "P. O. Number".PadRight(24) & _
                      "Invoice Amount" & vbCrLf)
         End With

         For i = 0 To dt.Rows.Count - 1
            With dt.Rows(i)
               ' print invoice detail data
               If .Item("c.customerid") = iLastCustomer Then
                  sb2.Append(Space(SL) & Format(.Item("invoiceid"), "0").PadLeft(9) & Space(5) & _
                            Format(DateValue(.Item("invoicedate")), "MM/dd/yyyy").PadRight(15) & _
                            CType(.Item("ponumber"), String).PadRight(24) & _
                            Format(.Item("balancedue"), "#,##0.00").PadLeft(14) & vbCrLf)
                  decInvoiceTotal += .Item("balancedue")
               Else
                  sb2.Append(Space(SL) & "Statement Total: ".PadLeft(55) & Format(decInvoiceTotal, "$#,##0.00").PadLeft(12) & vbCrLf)

                  ' print the data
                  oPS = New PrintStatement()
                  oPS.PrintFooter = False
                  oPS.TitleFontSize = 48
                  oPS.TitleFontStyle = "BI"
                  If m_Preview Then
                     oPS.PrintPreviewStatements(80, _
                        sb.ToString, _
                        sb2.ToString, _
                        ReportName, _
                        "Customer: " & iLastCustomer.ToString & " Statememnt Dated: " & Format(Today, "M/d/yyyy"))
                  Else
                     oPS.StartPrintStatements(80, _
                        sb.ToString, _
                        sb2.ToString, _
                        ReportName, _
                        "Customer: " & iLastCustomer.ToString & " Statememnt Dated: " & Format(Today, "M/d/yyyy"))
                  End If
                  ' start the next customer
                  sb = New StringBuilder()
                  sb2 = New StringBuilder()

                  sb.Append(Space(SL) & "Bill To:".PadRight(43) & "Remit To:" & vbCrLf)
                  sb.Append(Space(SL) & CType(.Item("companyname"), String).PadRight(43) & _
                     CorporateName & vbCrLf)
                  sb.Append(Space(SL) & CType(.Item("billingaddress1"), String).PadRight(43) & _
                     Address1 & vbCrLf)

                  sb.Append(Space(SL) & (CType(.Item("City"), String) & ", " & _
                      CType(.Item("State"), String) & "  " & _
                      CType(.Item("postalcode"), String)).PadRight(43) & _
                      City & ", " & State & " " & Zip & vbCrLf)

                  sb2.Append(Space(SL) & "Invoice #".PadRight(14) & _
                            "Invoice Date".PadRight(15) & _
                            "P. O. Number".PadRight(24) & _
                            "Invoice Amount" & vbCrLf)
                  iLastCustomer = .Item("c.customerid")
                  decInvoiceTotal = 0
                  sb2.Append(Space(SL) & Format(.Item("invoiceid"), "0").PadLeft(9) & Space(5) & _
                            Format(DateValue(.Item("invoicedate")), "MM/dd/yyyy").PadRight(15) & _
                            CType(.Item("ponumber"), String).PadRight(24) & _
                            Format(.Item("balancedue"), "#,##0.00").PadLeft(14) & vbCrLf)
                  decInvoiceTotal += .Item("balancedue")
               End If
            End With
         Next i

         ' print last customer's data
         If decInvoiceTotal > 0 Then
            sb2.Append(Space(SL) & "Statement Total: ".PadLeft(55) & Format(decInvoiceTotal, "$#,##0.00").PadLeft(12) & vbCrLf)

            oPS = New PrintStatement()
            oPS.PrintFooter = False
            oPS.TitleFontSize = 48
            oPS.TitleFontStyle = "BI"
            If m_Preview Then
               oPS.PrintPreviewStatements(80, _
                  sb.ToString, _
                  sb2.ToString, _
                  ReportName, _
                  "Customer: " & iLastCustomer.ToString & " Statememnt Dated: " & Format(Today, "M/d/yyyy"))
            Else
               oPS.StartPrintStatements(80, _
                  sb.ToString, _
                  sb2.ToString, _
                  ReportName, _
                  "Customer: " & iLastCustomer.ToString & " Statememnt Dated: " & Format(Today, "M/d/yyyy"))
            End If

         End If
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

#End Region



#Region " Constructor "
   Public Sub New()
      oDA = New CDataAccess()
      oUtil = New CUtilities()
   End Sub

#End Region


#Region " Public Properties "
   Public Property Preview() As Boolean
      Get
         Return m_Preview
      End Get
      Set(ByVal Value As Boolean)
         m_Preview = Value
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


#End Region

#Region " PrintStatement Class "

   Private Class PrintStatement
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
      Private WithEvents PrintStatements As Printing.PrintDocument
      Private msRptString As String ' holds the report string
      Private msRptString2 As String ' holds the statement details
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
      Private ColHdrArrayList As New ArrayList()
      Private ColHdrCount As Short = 0
      Public Heading As String
      Private DetailFontSize As Single
      Dim oUtil As CUtilities
      Dim oUtil2 As CUtilities
      Dim msLine As String
      Dim msToken As String
      Dim sFooter As String
      Dim miFileType As Integer
      Dim miWordWrap As Integer = 0
      Const DETAIL_FONT = "Courier New"
      Const DETAIL_FONT_SIZE_80 = 10
      Const DETAIL_FONT_SIZE_96 = 9
      Const DETAIL_FONT_SIZE_120 = 8 '118 with 1/4 margins
      'Const DETAIL_FONT_SIZE_132 = 7
      Const DETAIL_FONT_SIZE_160 = 6
      Const DETAIL_FONT_BOLD = True
      Private mbSepLines As Boolean
      Public TitleFontSize As Single = 12
      Public TitleFontStyle As String = "B"
      Public PrintFooter As Boolean = True
      Friend Sub StartPrintStatements(ByVal CharsPerLine As Integer, _
        ByRef sPrintBlock As String, _
        ByRef sPrintBlock2 As String, _
        ByVal sTitle As String, _
        ByVal sSubTitle As String, _
        Optional ByVal ColHdr1 As String = "", _
        Optional ByVal Landscape As Boolean = False, _
        Optional ByVal WordWrap As Integer = 0, _
        Optional ByVal SepLines As Boolean = False, _
        Optional ByVal ColHdr2 As String = "", _
        Optional ByVal ColHdr3 As String = "", _
        Optional ByVal ColHdr4 As String = "")

         'Dim previewDialog As New PrintPreviewDialog()

         Try
            mbSepLines = SepLines
            miWordWrap = WordWrap
            Portrait = Not Landscape
            msRptString = sPrintBlock
            msRptString2 = sPrintBlock2
            Title = sTitle
            SubTitle = sSubTitle
            SetUpColHdrArray(ColHdr1, ColHdr2, ColHdr3, ColHdr4)
            miChrPerLine = CharsPerLine
            ' create two objects so that we can use
            ' nested calls to MemoLine w/o stepping
            ' on each other...
            oUtil = New CUtilities()
            oUtil2 = New CUtilities()

            sFooter = "Printed on: " & Now.ToString
            Select Case CharsPerLine
               Case 80 : DetailFontSize = DETAIL_FONT_SIZE_80
               Case 96 : DetailFontSize = DETAIL_FONT_SIZE_96
               Case 120 : DetailFontSize = DETAIL_FONT_SIZE_120
               Case 160 : DetailFontSize = DETAIL_FONT_SIZE_160
               Case Else
                  MsgBox("You must specify a valid CharsPerLine parameter.", MsgBoxStyle.Exclamation)
                  Exit Sub
            End Select

            ' set up memoline
            miNL = oUtil.MLCount(msRptString2, 0) 'WordWrap)
            If miNL = 0 Then
               MsgBox("No lines to print in report string.", _
                  MsgBoxStyle.Exclamation)
               Exit Sub
            End If

            mI = 1

            ' set up the printdocument object
            PrintStatements = New Printing.PrintDocument()
            If Landscape Then
               PrintStatements.DefaultPageSettings.Landscape = True
            End If
            PrintStatements.Print() ' kick off the printing

         Catch ex As System.Exception
            StructuredErrorHandler(ex)
         End Try
      End Sub
      Friend Sub PrintPreviewStatements(ByVal CharsPerLine As Integer, _
       ByRef sPrintBlock As String, _
       ByRef sPrintBlock2 As String, _
       ByVal sTitle As String, _
       ByVal sSubTitle As String, _
       Optional ByVal ColHdr1 As String = "", _
       Optional ByVal Landscape As Boolean = False, _
       Optional ByVal WordWrap As Integer = 0, _
       Optional ByVal SepLines As Boolean = False, _
       Optional ByVal ColHdr2 As String = "", _
       Optional ByVal ColHdr3 As String = "", _
       Optional ByVal ColHdr4 As String = "")

         Dim previewDialog As New PrintPreviewDialog()

         Try
            mbSepLines = SepLines
            miWordWrap = WordWrap
            Portrait = Not Landscape
            msRptString = sPrintBlock
            msRptString2 = sPrintBlock2
            Title = sTitle
            SubTitle = sSubTitle
            SetUpColHdrArray(ColHdr1, ColHdr2, ColHdr3, ColHdr4)
            miChrPerLine = CharsPerLine
            ' create two objects so that we can use
            ' nested calls to MemoLine w/o stepping
            ' on each other...
            oUtil = New CUtilities()
            oUtil2 = New CUtilities()

            sFooter = "Printed on: " & Now.ToString
            Select Case CharsPerLine
               Case 80 : DetailFontSize = DETAIL_FONT_SIZE_80
               Case 96 : DetailFontSize = DETAIL_FONT_SIZE_96
               Case 120 : DetailFontSize = DETAIL_FONT_SIZE_120
               Case 160 : DetailFontSize = DETAIL_FONT_SIZE_160
            End Select

            ' set up memoline
            miNL = oUtil.MLCount(msRptString2, 0) 'WordWrap)
            If miNL = 0 Then
               MsgBox("No lines to print in report string.", _
                  MsgBoxStyle.Exclamation)
               Exit Sub
            End If

            mI = 1
            PrintStatements = New Printing.PrintDocument()
            If Landscape Then
               PrintStatements.DefaultPageSettings.Landscape = True
            End If
            PrintStatements.DocumentName = "Pioneer Print"
            previewDialog.Document = PrintStatements()
            previewDialog.ShowDialog()
            PrintStatements.Dispose()
            previewDialog.Dispose()

         Catch ex As System.Exception
            StructuredErrorHandler(ex)
         End Try
      End Sub

      Public Function PadR(ByVal s As String, ByVal n As Integer) As String
         If s.Trim.Length > n - 1 Then
            Return s.Substring(0, n)
         Else
            Return s.Trim & Space(n - Len(s.Trim))
         End If
      End Function
      Private Sub PrintStatements_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintStatements.PrintPage
         '` This method handles the callback from the PrintStatements
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
            PrintFont = New Font(DETAIL_FONT, DetailFontSize)
            sz = e.Graphics.MeasureString("m", PrintFont)
            LeftMargin = e.MarginBounds.Left - (7 * sz.Width)
            If Portrait Then
               RightMargin = e.MarginBounds.Right + (4 * sz.Width)
            Else
               RightMargin = e.MarginBounds.Right
            End If
            sz = e.Graphics.MeasureString("M", PrintFont)
            CharWidth = sz.Width
            TopMargin = e.MarginBounds.Top - (4 * sz.Height)
            BottomMargin = e.MarginBounds.Bottom + (2 * sz.Height)

            xPos = LeftMargin 'e.MarginBounds.Left
            yPos = TopMargin 'e.MarginBounds.Top

            ' Print the header on this page first thing
            ' First, print the title
            PageNbr += 1
            If TitleFontStyle.IndexOf("I") > -1 Then
               PrintFont = New Font("Arial", TitleFontSize, FontStyle.Italic Or FontStyle.Bold)
            Else
               PrintFont = New Font("Arial", TitleFontSize, FontStyle.Bold)
            End If
            LineHeight = PrintFont.GetHeight(e.Graphics)
            sz = e.Graphics.MeasureString(Title, PrintFont)
            TextWidth = sz.Width
            'LineWidth = e.MarginBounds.Right + e.MarginBounds.Left
            LineWidth = RightMargin + LeftMargin

            xPos = (LineWidth - TextWidth) / 2
            e.Graphics.DrawString(Title, _
                                    PrintFont, _
                                    Brushes.Black, _
                                    xPos, _
                                    yPos, _
                                    New StringFormat())
            yPos += (LineHeight * 1)

            ' Second, print the SubTitle
            PrintFont = New Font(DETAIL_FONT, DetailFontSize, FontStyle.Bold)
            xPos = LeftMargin + (5 * CharWidth) 'e.MarginBounds.Left
            LineHeight = PrintFont.GetHeight(e.Graphics)
            yPos += LineHeight
            ' remembe where hdr data starts
            Dim yPos2 As Single = yPos
            yPos += LineHeight

            sz = e.Graphics.MeasureString(sHdrLine, PrintFont)
            TextWidth = sz.Width
            sHdrLine = SubTitle.Trim
            ' xPos = (LineWidth - TextWidth) / 3
            e.Graphics.DrawString(sHdrLine, _
               PrintFont, _
               Brushes.Black, _
               xPos, _
               yPos, _
               New StringFormat())

            ' next print the page number on the right end
            ' of the first line
            sPageNbr = PageNbr.ToString
            'xPos = e.MarginBounds.Right - (8 * sPageNbr.Length)
            If Portrait Then
               xPos = RightMargin - (20 * sPageNbr.Length)
            Else
               xPos = RightMargin - (5 * sPageNbr.Length)
            End If

            e.Graphics.DrawString(sPageNbr, _
                                    PrintFont, _
                                    Brushes.Black, _
                                    xPos, _
                                    yPos, _
                                    New StringFormat())
            yPos += LineHeight * 2

            ' next print the bill and remit data
            PrintFont = New Font(DETAIL_FONT, DetailFontSize, FontStyle.Bold)
            xPos = LeftMargin + (5 * CharWidth) 'e.MarginBounds.Left
            LineHeight = PrintFont.GetHeight(e.Graphics)
            Dim iNL As Integer = oUtil2.MLCount(msRptString, 0)
            Dim k As Integer
            For k = 1 To iNL
               sLine = oUtil.MemoLine(msRptString, 0, k)
               e.Graphics.DrawString(sLine, _
                  PrintFont, _
                  Brushes.Black, _
                  xPos, _
                  yPos, _
                  New StringFormat())
               yPos += LineHeight
            Next k
            yPos += LineHeight

            ' now print a rectangle to house the header
            ' put a little space before the line
            Dim yPos3 As Single = yPos
            x1 = LeftMargin + (3 * CharWidth) 'e.MarginBounds.Left
            y1 = yPos2
            x2 = RightMargin 'e.MarginBounds.Right
            y2 = yPos
            e.Graphics.DrawRectangle(drPen, x1, y1, x2 - x1, yPos3 - yPos2)

            yPos += LineHeight

            ' now print a rectangle to house the details
            ' put a little space before the line
            x1 = LeftMargin + (3 * CharWidth) 'e.MarginBounds.Left
            y1 = yPos
            x2 = RightMargin 'e.MarginBounds.Right
            y2 = yPos
            e.Graphics.DrawRectangle(drPen, x1, y1, x2 - x1, BottomMargin - yPos)

            PrintFont = New Font(DETAIL_FONT, DetailFontSize, FontStyle.Bold)

            ' if a headings are extant, print them
            If ColHdrCount > 0 Then
               Dim p As Short
               For p = 0 To ColHdrCount - 1
                  yPos += 2
                  If ColHdrArrayList(p).Length > 0 Then
                     xPos = LeftMargin 'e.MarginBounds.Left
                     LineHeight = PrintFont.GetHeight(e.Graphics)
                     sHdrLine = SubTitle
                     e.Graphics.DrawString(ColHdrArrayList(0), _
                        PrintFont, _
                        Brushes.Black, _
                        xPos, _
                        yPos, _
                        New StringFormat())
                     yPos += LineHeight

                     ' now print a line
                     ' put a little space before the line
                     x1 = LeftMargin 'e.MarginBounds.Left
                     y1 = yPos
                     x2 = RightMargin 'e.MarginBounds.Right
                     y2 = yPos
                     e.Graphics.DrawLine(drPen, x1, y1, x2, y2)
                     yPos += 4
                     ' now print a line
                     ' put a little space before the line
                     x1 = LeftMargin 'e.MarginBounds.Left
                     y1 = yPos
                     x2 = RightMargin 'e.MarginBounds.Right
                     y2 = yPos
                     e.Graphics.DrawLine(drPen, x1, y1, x2, y2)
                  End If
               Next p
            End If

            yPos += 4

            ' we can just call memoline limiting the lines to
            ' a length that will fit on a line of print

            ' if we have a blank line, don't print a sep line
            ' after it...
            Dim bBlank As Boolean

            ' Print directives
            ' <:CH1 Heading 1
            ' <:CH2 Heading 2
            ' <:CH3 Heading 3
            ' <:CH4 Heading 4
            ' <:NEWPAGE
            ' <:NOLINES
            ' <:LINES
            ' <:SUBTITLE
            ' <:PAGENBR0
            For i = mI To miNL
               ' just get a line and print it and then
               ' check to see if there is enough room to print another
               ' line.
               sLine = oUtil.MemoLine(msRptString2, 0, i)

               ' ck for print directives
               If sLine.Length > 0 Then
                  If sLine.Substring(0, 2) = "<:" Then
                     ' we have a print directive telling us to change something
                     If sLine.Substring(2, 3) = "CH1" Then
                        ' <:CH>New col heading 1
                        ColHdrArrayList(0) = sLine.Substring(6)
                        GoTo GetNextLine
                     ElseIf sLine.Substring(2, 3) = "CH2" Then
                        ' <:CH>New col heading 1
                        ColHdrArrayList(1) = sLine.Substring(6)
                        GoTo GetNextLine
                     ElseIf sLine.Substring(2, 3) = "CH3" Then
                        ' <:CH>New col heading 1
                        ColHdrArrayList(2) = sLine.Substring(6)
                        GoTo GetNextLine
                     ElseIf sLine.Substring(2, 3) = "CH4" Then
                        ' <:CH>New col heading 1
                        ColHdrArrayList(3) = sLine.Substring(6)
                        GoTo GetNextLine
                     ElseIf sLine.Substring(2, 7).StartsWith("NEWPAGE") Then
                        GoTo EndPage
                     ElseIf sLine.Substring(2, 7).StartsWith("NOLINES") Then
                        mbSepLines = False
                     ElseIf sLine.Substring(2, 5).StartsWith("LINES") Then
                        mbSepLines = True
                     ElseIf sLine.Substring(2, 8).StartsWith("PAGENBR0") Then
                        PageNbr = 0 : GoTo GetNextLine
                     ElseIf sLine.Substring(2, 8).StartsWith("SUBTITLE") Then
                        ' <:SUBTITLE>NEW sub title line
                        Me.SubTitle = sLine.Substring(11)
                        GoTo GetNextLine
                     End If
                  End If
               End If


               If (sLine Is Nothing OrElse sLine.Length = 0) Then
                  ' dont print a blank line, just bump yPos
                  bBlank = True
               Else
                  LineHeight = PrintFont.GetHeight(e.Graphics)
                  xPos = LeftMargin 'e.MarginBounds.Left
                  bBlank = False
               End If

               PrintFont = New Font(DETAIL_FONT, DetailFontSize, FontStyle.Bold)
               e.Graphics.DrawString(sLine, _
                                    PrintFont, _
                                    Brushes.Black, _
                                    xPos, _
                                    yPos, _
                                    New StringFormat())
               yPos += LineHeight

               If mbSepLines Then
                  If Not bBlank Then
                     ' insert the print of a seperator line if SepLines=True
                     ' put a little space before the line
                     yPos += 2
                     x1 = LeftMargin 'e.MarginBounds.Left
                     y1 = yPos
                     x2 = RightMargin 'e.MarginBounds.Right
                     y2 = yPos
                     e.Graphics.DrawLine(drPen, x1, y1, x2, y2)
                     yPos += 2
                  End If
               End If

               ' check to see if we are at the end of the page
               'If yPos >= (e.MarginBounds.Bottom - LineHeight) Then
               If yPos >= (BottomMargin - LineHeight) Then
                  ' end of page, ck for more lines to print
                  ' after print the footer
                  ' put a little space before the line
EndPage:
                  yPos += 2
                  x1 = LeftMargin 'e.MarginBounds.Left
                  y1 = yPos
                  x2 = RightMargin 'e.MarginBounds.Right
                  y2 = yPos
                  If Not PrintFooter Then GoTo EndPage2
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
EndPage2:
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
GetNextLine:
            Next

            e.HasMorePages = False
            PageNbr = 0 ' in case called again from preview print
            'print a footer on the last page
            yPos = BottomMargin - LineHeight 'e.MarginBounds.Bottom - LineHeight
            x1 = LeftMargin 'e.MarginBounds.Left
            y1 = yPos
            x2 = RightMargin 'e.MarginBounds.Right
            y2 = yPos
            'e.Graphics.DrawLine(drPen, x1, y1, x2, y2)
            PrintFont.Dispose()
            PrintFont = New Font("Arial", 10)
            xPos = LeftMargin 'e.MarginBounds.Left
            'e.Graphics.DrawString(sFooter, _
            '   PrintFont, _
            '   Brushes.Black, _
            '   xPos, _
            '   yPos, _
            '   New StringFormat())
         Catch ex As System.Exception
            StructuredErrorHandler(ex)
         End Try
      End Sub
      Private Sub SetUpColHdrArray(ByVal ColHdr1 As String, ByVal ColHdr2 As String, _
         ByVal ColHdr3 As String, ByVal ColHdr4 As String)
         If ColHdr1.Length > 0 Then
            ColHdrArrayList.Add(ColHdr1)
            ColHdrCount += 1
         End If
         If ColHdr2.Length > 0 Then
            ColHdrArrayList.Add(ColHdr2)
            ColHdrCount += 1
         End If
         If ColHdr3.Length > 0 Then
            ColHdrArrayList.Add(ColHdr3)
            ColHdrCount += 1
         End If
         If ColHdr4.Length > 0 Then
            ColHdrArrayList.Add(ColHdr4)
            ColHdrCount += 1
         End If
      End Sub
   End Class

#End Region
End Class
