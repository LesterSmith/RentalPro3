   ''' This is a standalone print class.
   ''' The constructor accepts the text for the print objects,
   ''' Once instantiated,
   ''' Simply call the Print or Preview method.
   Public Class CPioneerPrintObject
      Private _Left As Single
      Private _Top As Single
      Private _Height As Single
      Private _Width As Single
      Private _FontName As String
      Private _FontSize As Single
      Private _FontBold As Boolean
      Private _FontItalic As Boolean
      Private _FontUnderline As Boolean
      Private _Text As String
      Private _Box As Boolean
      Private _PenWidth As Single
      Dim previewDialog As New PrintPreviewDialog
      WithEvents PrintDoc As Printing.PrintDocument
   Public Sub New( _
      ByVal tbnameShipTo As String, _
      ByVal tbnameBillTo As String, _
      ByVal tbnameInvoiceHdrData As String, _
      ByVal tbnameDetails_Total As String)
      AddObjectToList("Courier New", _
         11.25, _
         25.0499934326172, _
         100.199991445312, _
         75.1499802978516, _
         75.1499935839844, _
         tbnameShipTo, _
         False, _
         False, _
         False, _
         0, _
         1)
      AddObjectToList("Courier New", _
         11.25, _
         25.0499934326172, _
         212.924981821289, _
         75.1499802978516, _
         75.1499935839844, _
         tbnameBillTo, _
         False, _
         False, _
         False, _
         0, _
         1)
      AddObjectToList("Courier New", _
         11.25, _
         488.474871936035, _
         100.199991445312, _
         75.1499802978516, _
         75.1499935839844, _
         tbnameInvoiceHdrData, _
         False, _
         False, _
         False, _
         0, _
         1)
      AddObjectToList("Courier New", _
         11.25, _
         25.0499934326172, _
         350.699970058594, _
         75.1499802978516, _
         75.1499935839844, _
         tbnameDetails_Total, _
         False, _
         False, _
         False, _
         0, _
         1)

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
         Dim Box As Boolean
         Dim PenWidth As Single
      End Structure
      Public printObjectArray As New ArrayList
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
                                 ByVal Box As Boolean, _
                                 ByVal PenWidth As Single)
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
         End With
         printObjectArray.Add(printObject)
      End Sub
      Public Sub Print()
         Dim PD As New Printing.PrintDocument
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
              If .Box Then
                 Dim myPen As New Pen(Color.Black, .PenWidth)
                 e.Graphics.DrawRectangle(myPen, printObject.XPos, printObject.YPos, printObject.xWidth, printObject.yHeight)
              Else
                 e.Graphics.DrawString(printObject.Text, _
                                       printFont, _
                                       Brushes.Black, _
                                       .XPos, _
                                       .YPos, _
                                       New StringFormat)
              End If
           End With
         Next i
         e.HasMorePages = False
      End Sub
      Public Property Left() As Single
         Get
            Return _Left
         End Get
         Set(ByVal Value As Single)
            _Left = Value
         End Set
      End Property
      Public Property Top() As Single
         Get
            Return _Top
         End Get
         Set(ByVal Value As Single)
            _Top = Value
         End Set
      End Property

      Public Property Height() As Single
         Get
            Return _Height
         End Get
         Set(ByVal Value As Single)
            _Height = Value
         End Set
      End Property
      Public Property Width() As Single
         Get
            Return _Width
         End Get
         Set(ByVal Value As Single)
            _Width = Value
         End Set
      End Property

      Public Property FontName() As String
         Get
            Return _FontName
         End Get
         Set(ByVal Value As String)
            _FontName = Value
         End Set
      End Property
      Public Property FontSize() As Single
         Get
            Return _FontSize
         End Get
         Set(ByVal Value As Single)
            _FontSize = Value
         End Set
      End Property
      Public Property FontBold() As Boolean
         Get
            Return _FontBold
         End Get
         Set(ByVal Value As Boolean)
            _FontBold = Value
         End Set
      End Property
      Public Property FontItalic() As Boolean
         Get
            Return _FontItalic
         End Get
         Set(ByVal Value As Boolean)
            _FontItalic = Value
         End Set
      End Property
      Public Property FontUnderline() As Boolean
         Get
            Return _FontUnderline
         End Get
         Set(ByVal Value As Boolean)
            _FontUnderline = Value
         End Set
      End Property

      Public Property Text() As String
         Get
            Return _Text
         End Get
         Set(ByVal Value As String)
            _Text = Value
         End Set
      End Property
         Public Property PenWidth() As Single
            Get
               Return _PenWidth
            End Get
            Set(ByVal Value As Single)
              _PenWidth = Value
            End Set
         End Property

   End Class
