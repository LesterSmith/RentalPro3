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
Public Class CPrintString
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
   Private msRptString As String ' holds the report string
   Private miNL As Integer ' number of lines in the report
   Private miNL2 As Integer ' number of sub lines in line
   Private mI As Integer
   Private mI2 As Integer
   Private CurrentLine As Integer ' curr print line on a page
   Private miChrPerLine As Integer ' nbr chars to print per line
   Public Title As String
   Public PageNbr As Integer
   Private Portrait As Boolean = True ' landscape if false
   Public SubTitle As String
   Public Heading As String
   Public NewPage As Boolean = False
   Public DetailFontSize As Single
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
   Const DETAIL_FONT_SIZE_120 = 8
   'Const DETAIL_FONT_SIZE_132 = 7
   Const DETAIL_FONT_SIZE_160 = 6
   Const DETAIL_FONT_BOLD = True
   Private mbSepLines As Boolean
   Protected TitleFontSize As Single = 12
   Protected TitleFontStyle As String = "B"



#End Region

#Region " Public Methods "
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


   Sub StartPrint(ByVal iChrPerLine As Integer, _
         ByRef sPrintBlock As String, _
         ByVal sTitle As String, _
         ByVal sSubTitle As String, _
         Optional ByVal sHeading As String = "", _
         Optional ByVal Landscape As Boolean = False, _
         Optional ByVal WordWrap As Integer = 0, _
         Optional ByVal SepLines As Boolean = False)


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
         mbSepLines = SepLines
         oUtil = New CUtilities()
         sFooter = "Printed on: " & Now.ToString
         miWordWrap = WordWrap
         msRptString = sPrintBlock
         Title = sTitle
         SubTitle = sSubTitle
         Heading = sHeading
         miChrPerLine = iChrPerLine
         Portrait = Not Landscape
         Select Case iChrPerLine
            Case 80 : DetailFontSize = DETAIL_FONT_SIZE_80
            Case 96 : DetailFontSize = DETAIL_FONT_SIZE_96
            Case 120 : DetailFontSize = DETAIL_FONT_SIZE_120
            Case 160 : DetailFontSize = DETAIL_FONT_SIZE_160
         End Select

         ' set up memoline
         miNL = oUtil.MLCount(msRptString, WordWrap)
         If miNL = 0 Then
            MsgBox("No lines to print in report string.", _
               MsgBoxStyle.Exclamation)
            Exit Sub
         End If

         mI = 1

         ' set up the printdocument object
         PrintDoc = New Printing.PrintDocument()
         If Landscape Then
            PrintDoc.DefaultPageSettings.Landscape = True
         End If
         PrintDoc.Print() ' kick off the printing


      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

   Friend Sub PrintPreview(ByVal iChrPerLine As Integer, _
      ByRef sPrintBlock As String, _
      ByVal sTitle As String, _
      ByVal sSubTitle As String, _
      Optional ByVal sHeading As String = "", _
      Optional ByVal Landscape As Boolean = False, _
      Optional ByVal WordWrap As Integer = 0, _
      Optional ByVal SepLines As Boolean = False)


      Dim previewDialog As New PrintPreviewDialog()



      Try
         mbSepLines = SepLines
         miWordWrap = WordWrap
         Portrait = Not Landscape
         msRptString = sPrintBlock
         Title = sTitle
         SubTitle = sSubTitle
         Heading = sHeading
         miChrPerLine = iChrPerLine
         ' create two objects so that we can use
         ' nested calls to MemoLine w/o stepping
         ' on each other...
         oUtil = New CUtilities()
         sFooter = "Printed on: " & oUtil.GetNowPrintString()
         Select Case iChrPerLine
            Case 80 : DetailFontSize = DETAIL_FONT_SIZE_80
            Case 96 : DetailFontSize = DETAIL_FONT_SIZE_96
            Case 120 : DetailFontSize = DETAIL_FONT_SIZE_120
            Case 160 : DetailFontSize = DETAIL_FONT_SIZE_160
         End Select

         ' set up memoline
         miNL = oUtil.MLCount(msRptString, WordWrap)
         If miNL = 0 Then
            MsgBox("No lines to print in report string.", _
               MsgBoxStyle.Exclamation)
            Exit Sub
         End If

         mI = 1
         PrintDoc = New Printing.PrintDocument()
         If Landscape Then
            PrintDoc.DefaultPageSettings.Landscape = True
         End If
         PrintDoc.DocumentName = "NETCommander Print"
         previewDialog.Document = PrintDoc
         previewDialog.ShowDialog()
         PrintDoc.Dispose()
         previewDialog.Dispose()

      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

#End Region




#Region " Private Methods "
   Private Sub PrtDoc_PrintPage(ByVal sender As Object, _
      ByVal e As System.Drawing.Printing.PrintPageEventArgs) _
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
         yPos += LineHeight

         ' Second, print the SubTitle
         PrintFont = New Font("Arial", 10)
         xPos = LeftMargin 'e.MarginBounds.Left
         LineHeight = PrintFont.GetHeight(e.Graphics)
         sHdrLine = SubTitle
         e.Graphics.DrawString(sHdrLine, _
            PrintFont, _
            Brushes.Black, _
            xPos, _
            yPos, _
            New StringFormat())

         ' next print the page number on the right end
         ' of the first line
         sPageNbr = PageNbr.ToString
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
         yPos += LineHeight

         ' now print a line
         ' put a little space before the line
         x1 = LeftMargin 'e.MarginBounds.Left
         y1 = yPos
         x2 = RightMargin 'e.MarginBounds.Right
         y2 = yPos
         e.Graphics.DrawLine(drPen, x1, y1, x2, y2)

         PrintFont = New Font(DETAIL_FONT, DetailFontSize, FontStyle.Bold)

         ' if a heading is extant, print it
         yPos += 2
         If Heading.Length > 0 Then
            xPos = LeftMargin 'e.MarginBounds.Left
            LineHeight = PrintFont.GetHeight(e.Graphics)
            sHdrLine = SubTitle
            e.Graphics.DrawString(Heading, _
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
         yPos += 4

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
            sLine = oUtil.MemoLine(msRptString, miWordWrap, i)

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
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Sub

#End Region


End Class
