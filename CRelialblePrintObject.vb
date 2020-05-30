   ''' This is a standalone print class.
   ''' The constructor accepts the text for the print objects,
   ''' Once instantiated,
   ''' Simply call the Print or Preview method.
   Public Class CRelialblePrintObject
   'Private _Left As Single
   'Private _Top As Single
   'Private _Height As Single
   'Private _Width As Single
   'Private _FontName As String
   'Private _FontSize As Single
   'Private _FontBold As Boolean
   'Private _FontItalic As Boolean
   'Private _FontUnderline As Boolean
   'Private _Text As String
   'Private _Box As Boolean
   'Private _PenWidth As Single
      Dim previewDialog As New PrintPreviewDialog
      WithEvents PrintDoc As Printing.PrintDocument
   Public Sub New( _
      ByVal tbnameBillTo As String, _
      ByVal tbnameShipTo As String, _
      ByVal tbnameInvoice As String, _
      ByVal tbnameTimeOut As String, _
      ByVal tbnameTimeIn As String, _
      ByVal tbnameElapsed As String, _
      ByVal tbnameJobPhone As String, _
      ByVal tbnameCheckedInBy As String, _
      ByVal tbnameDLNumber As String, _
      ByVal tbnameAgent As String, _
      ByVal tbnamePONumber As String, _
      ByVal tbnameDueInTime As String, _
      ByVal tbnameJobLocation As String, _
      ByVal tbnameTotalDesc As String, _
      ByVal tbnameTaxid As String, _
      ByVal tbnamePaidOption As String, _
      ByVal tbnameDetailLines As String, _
      ByVal tbnameWrittenBy As String)
      AddObjectToList("Courier New", _
         11.25, _
         106.462472088623, _
         125.249989306641, _
         100.199973730469, _
         100.199991445312, _
         tbnameBillTo, _
         False, _
         False, _
         False, _
         "Text", _
         1, _
         "Black")
      AddObjectToList("Courier New", _
         11.25, _
         106.462472088623, _
         237.974979682617, _
         100.199973730469, _
         100.199991445312, _
         tbnameShipTo, _
         False, _
         False, _
         False, _
         "Text", _
         1, _
         "Black")
      AddObjectToList("Courier New", _
         11.25, _
         588.674845666504, _
         87.6749925146484, _
         100.199973730469, _
         100.199991445312, _
         tbnameInvoice, _
         False, _
         False, _
         False, _
         "Text", _
         1, _
         "Black")
      AddObjectToList("Courier New", _
         11.25, _
         588.674845666504, _
         162.824986098633, _
         100.199973730469, _
         100.199991445312, _
         tbnameTimeOut, _
         False, _
         False, _
         False, _
         "Text", _
         1, _
         "Black")
      AddObjectToList("Courier New", _
         11.25, _
         588.674845666504, _
         125.249989306641, _
         100.199973730469, _
         100.199991445312, _
         tbnameTimeIn, _
         False, _
         False, _
         False, _
         "Text", _
         1, _
         "Black")
      AddObjectToList("Courier New", _
         11.25, _
         588.674845666504, _
         187.874983959961, _
         100.199973730469, _
         100.199991445312, _
         tbnameElapsed, _
         False, _
         False, _
         False, _
         "Text", _
         1, _
         "Black")
      AddObjectToList("Courier New", _
         11.25, _
         12.5249967163086, _
         356.962469523926, _
         100.199973730469, _
         100.199991445312, _
         tbnameJobPhone, _
         False, _
         False, _
         False, _
         "Text", _
         1, _
         "Black")
      AddObjectToList("Courier New", _
         11.25, _
         200.399947460937, _
         325.649972197266, _
         100.199973730469, _
         100.199991445312, _
         tbnameCheckedInBy, _
         False, _
         False, _
         False, _
         "Text", _
         1, _
         "Black")
      AddObjectToList("Courier New", _
         11.25, _
         200.399947460937, _
         356.962469523926, _
         100.199973730469, _
         100.199991445312, _
         tbnameDLNumber, _
         False, _
         False, _
         False, _
         "Text", _
         1, _
         "Black")
      AddObjectToList("Courier New", _
         11.25, _
         407.062393280029, _
         325.649972197266, _
         100.199973730469, _
         100.199991445312, _
         tbnameAgent, _
         False, _
         False, _
         False, _
         "Text", _
         1, _
         "Black")
      AddObjectToList("Courier New", _
         11.25, _
         407.062393280029, _
         356.962469523926, _
         100.199973730469, _
         100.199991445312, _
         tbnamePONumber, _
         False, _
         False, _
         False, _
         "Text", _
         1, _
         "Black")
      AddObjectToList("Courier New", _
         11.25, _
         607.462340740967, _
         356.962469523926, _
         100.199973730469, _
         100.199991445312, _
         tbnameDueInTime, _
         False, _
         False, _
         False, _
         "Text", _
         1, _
         "Black")
      AddObjectToList("Courier New", _
         11.25, _
         607.462340740967, _
         325.649972197266, _
         100.199973730469, _
         100.199991445312, _
         tbnameJobLocation, _
         False, _
         False, _
         False, _
         "Text", _
         1, _
         "Black")
      AddObjectToList("Courier New", _
         9.75, _
         538.574767181923, _
         814.124899833984, _
         100.199956685009, _
         100.199987671875, _
         tbnameTotalDesc, _
         False, _
         False, _
         False, _
         "Text", _
         1, _
         "Black")
      AddObjectToList("Courier New", _
         11.25, _
         576.149848950195, _
         244.237479147949, _
         100.199973730469, _
         100.199991445312, _
         tbnameTaxid, _
         False, _
         False, _
         False, _
         "Text", _
         1, _
         "Black")
      AddObjectToList("Courier New", _
         11.25, _
         576.149848950195, _
         275.549976474609, _
         100.199973730469, _
         100.199991445312, _
         tbnamePaidOption, _
         False, _
         False, _
         False, _
         "Text", _
         1, _
         "Black")
      AddObjectToList("Courier New", _
         9.75, _
         12.5249945856261, _
         413.324949146484, _
         100.199956685009, _
         100.199987671875, _
         tbnameDetailLines, _
         False, _
         False, _
         False, _
         "Text", _
         1, _
         "Black")
      AddObjectToList("Courier New", _
         11.25, _
         12.5249967163086, _
         325.649972197266, _
         100.199973730469, _
         100.199991445312, _
         tbnameWrittenBy, _
         False, _
         False, _
         False, _
         "Text", _
         1, _
         "Black")

         PrintDoc = New Printing.PrintDocument
     End Sub
   Public Structure PrintObjects
      Dim FontName As String
      Dim FontSize As Single
      Dim XPos As Single
      Dim YPos As Single
      Dim yHeight As Single
      Dim xWidth As Single
      Dim FontBold As Boolean
      Dim FontUnderlne As Boolean
      Dim FontItalic As Boolean
      Dim Text As String
      Dim Name As String
      Dim Left As Single
      Dim Top As Single
      Dim Height As Single
      Dim Width As Single
      Dim HCI As Single
      Dim VCI As Single
      Dim CharWidth As Single
      Dim LineHeight As Single
      Dim Points As Single
      Dim Box As String
      Dim PenWidth As Single
      Dim Color As String
   End Structure
   Public printObjectArray As New ArrayList()
   Private printObject As PrintObjects

   Public Sub AddObjectToList(ByVal FontName As String, _
                              ByVal FontSize As Single, _
                              ByVal Left As Single, _
                              ByVal Top As Single, _
                              ByVal Height As Single, _
                              ByVal Width As Single, _
                              ByVal TextString As String, _
                              ByVal FontBold As Boolean, _
                              ByVal FontItalic As Boolean, _
                              ByVal FontUnderline As Boolean, _
                              ByVal Box As String, _
                              ByVal PenWidth As Single, _
                              ByVal Color As String)
      With printObject
         .FontName = FontName
         .FontSize = FontSize
         .XPos = Left
         .YPos = Top
         .xWidth = Width
         .yHeight = Height
         .Text = TextString
         .FontBold = FontBold
         .FontItalic = FontItalic
         .FontUnderlne = FontUnderline
         .Box = Box
         .PenWidth = PenWidth
         .Color = Color
      End With
      printObjectArray.Add(printObject)
   End Sub
   Public Sub Print()
      Dim PD As New Printing.PrintDocument()
      PD.Print()
   End Sub
   Public Sub Preview()
      PrintDoc.DocumentName = "Print Object Test"
      previewDialog.Document = PrintDoc
      previewDialog.ShowDialog()
      PrintDoc.Dispose()
      previewDialog.Dispose()
   End Sub
   Private Sub PrintDoc_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDoc.PrintPage
      Dim i As Integer

      For i = 0 To printObjectArray.Count - 1
         printObject = printObjectArray.Item(i)
         With printObject
            Dim fontAttributes As System.Drawing.FontStyle = _
              (IIf(printObject.FontBold, 1, 0) Or _
              IIf(printObject.FontItalic, 2, 0) Or _
              IIf(printObject.FontUnderlne, 4, 0))
            Dim printFont As Font
            If fontAttributes <> 0 And fontAttributes <> FontStyle.Regular Then
               printFont = New Font(.FontName, .FontSize, fontAttributes)
            Else
               printFont = New Font(.FontName, .FontSize)
            End If
            If .Box = "Box" Then
               Dim c As Color
               c = Color.FromName(.Color)
               Dim myPen As New Pen(c, .PenWidth)
               e.Graphics.DrawRectangle(myPen, printObject.XPos, printObject.YPos, printObject.xWidth, printObject.yHeight)
            ElseIf .Box = "Line" Then
               Dim c As Color
               c = Color.FromName(.Color)
               Dim myPen As New Pen(c, .PenWidth)
               e.Graphics.DrawLine(myPen, printObject.XPos, printObject.YPos, printObject.XPos + printObject.xWidth, printObject.YPos + printObject.yHeight)
            ElseIf .Box = "Text" Then
               Dim b As New SolidBrush(Color.FromName(.Color))
               e.Graphics.DrawString(printObject.Text, _
                                     printFont, _
                                     b, _
                                     .XPos, _
                                     .YPos, _
                                     New StringFormat())
            End If
         End With
      Next i
      e.HasMorePages = False
   End Sub

End Class
