Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms
''' This is a standalone print class.
''' The constructor accepts the text for the print objects,
''' Once instantiated,
''' Simply call the Print or Preview method.
Public Class PioneerNewPrintObject
    Dim previewDialog As New PrintPreviewDialog
    WithEvents PrintDoc As Printing.PrintDocument
    Private _printerName As String = String.Empty
    Public Property PrinterName As String
        Get
            Return _printerName
        End Get
        Set(ByVal value As String)
            _printerName = value
        End Set
    End Property
    Public Sub New( _
       ByVal tbnameLogo As String, _
       ByVal tbnameCompanyName As String,
       ByVal tbnameTitle As String,
       ByVal tbnameShipTo As String,
       ByVal tbnameBillTo As String,
       ByVal tbnameInvoiceHdrData As String,
       ByVal tbnameDetails As String,
       ByVal printerName As String)
        AddObjectToList("Courier New", _
           18, _
           37.5749983905995, _
           25.0499978298926, _
           150.299993562398, _
           62.6249945747316, _
           tbnameLogo, _
           True, _
           True, _
           False, _
           "Text", _
           3, _
           "Black")
        AddObjectToList("Courier New", _
           12, _
           225.450008945925, _
           25.0500003455172, _
           300.600011927899, _
           62.6250008637929, _
           tbnameCompanyName, _
           True, _
           False, _
           False, _
           "Text", _
           1, _
           "Black")
        AddObjectToList("Courier New", _
           12, _
           576.150022861807, _
           25.0500003455172, _
           200.400007951933, _
           50.1000006910343, _
           tbnameTitle, _
           True, _
           False, _
           False, _
           "Text", _
           1, _
           "Black")
        AddObjectToList("Courier New", _
           11.25, _
           37.5749828151941, _
           112.724988810047, _
           400.799816695404, _
           100.199990053375, _
           tbnameShipTo, _
           True, _
           False, _
           False, _
           "Text", _
           1, _
           "Black")
        AddObjectToList("Courier New", _
           11.25, _
           37.5749828151941, _
           212.924978863422, _
           400.799816695404, _
           100.199990053375, _
           tbnameBillTo, _
           True, _
           False, _
           False, _
           "Text", _
           1, _
           "Black")
        AddObjectToList("Courier New", _
           11.25, _
           488.474776597524, _
           112.724988810047, _
           325.649851065016, _
           112.724988810047, _
           tbnameInvoiceHdrData, _
           True, _
           False, _
           False, _
           "Text", _
           1, _
           "Black")
        AddObjectToList("Courier New", _
           11.25, _
           25.0499885434628, _
           375.749962700157, _
           776.549644847345, _
           488.474951510204, _
           tbnameDetails, _
           True, _
           False, _
           False, _
           "Text", _
           1, _
           "Black")


        PrintDoc = New Printing.PrintDocument
        If Not String.IsNullOrEmpty(printerName) Then
            Dim ps As New PrinterSettings
            ps.PrinterName = printerName

            PrintDoc.PrinterSettings = ps
        End If
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
        PrintDoc.Print()
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
                                          New StringFormat)
                End If
            End With
        Next i
        e.HasMorePages = False
    End Sub
End Class
