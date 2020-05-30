Imports System
Imports System.Drawing.Printing
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Public Class Printers
    Public Shared Function GetDefaultPrinter() As String
        Dim i As Integer = 0
        Dim defaultPrinter As String = String.Empty

        For Each printerName As String In PrinterSettings.InstalledPrinters
            Dim objPrinter As New PrinterSettings
            objPrinter.PrinterName = printerName
            If objPrinter.IsDefaultPrinter Then
                defaultPrinter = printerName
            End If
            i += 1
        Next
        Return defaultPrinter
    End Function
End Class
