Imports System
Imports System.Drawing.Printing
Imports System.Windows.Forms
Imports System.Runtime.InteropServices

Public Class frmPrinters
    Private _dirty As Boolean
    Private Sub frmPrinters_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ListAllPrinters()
    End Sub

    Private Sub ListAllPrinters()
        lbPrinters.Items.Clear()
        Dim i As Integer = 0
        Dim defaultPrinter As String = Printers.GetDefaultPrinter()

        For Each item As String In PrinterSettings.InstalledPrinters
            lbPrinters.Items.Add(item.ToString())
            Dim printer As New PrinterSettings
            printer.PrinterName = item
            If printer.PrinterName = defaultPrinter Then 'printer.IsDefaultPrinter Then
                lbPrinters.SetItemChecked(i, True)
            End If
            i += 1
        Next
    End Sub

    Private Sub lbPrinters_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbPrinters.SelectedIndexChanged
    End Sub

    Private Sub lbPrinters_ItemCheck(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles lbPrinters.ItemCheck
        For i As Integer = 0 To lbPrinters.Items.Count - 1
            If i <> e.Index Then
                lbPrinters.SetItemCheckState(i, False)
            End If
        Next
        _dirty = True
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        If _dirty Then
            If MsgBox("You have changed the selection.  Are you sure you want to close without saving the new default printer?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Confirm Close Without Saving") = MsgBoxResult.Yes Then
                Me.Close()
            End If
        End If
    End Sub

    Private Sub btnSetDefault_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSetDefault.Click
        If Not _dirty Then
            MsgBox("You have not changed the default printer that was already selected.")
            Exit Sub
        End If

        Dim pname As String = lbPrinters.SelectedItem.ToString()
        SetDefaultPrinter(pname)

    End Sub

#Region "GetDefaultPrinter"
    <DllImport("winspool.drv", EntryPoint:="GetDefaultPrinter", _
         SetLastError:=True, CharSet:=CharSet.Auto, _
         ExactSpelling:=False, _
         CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Function GetDefaultPrinter(ByVal pszBuffer As System.Text.StringBuilder, _
                                              ByRef BufferSize As Int32) As Boolean

    End Function
#End Region

#Region "SetDefaultPrinter"
    <DllImport("winspool.drv", EntryPoint:="SetDefaultPrinter", _
         SetLastError:=True, CharSet:=CharSet.Auto, _
         ExactSpelling:=False, _
         CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Function SetDefaultPrinter(ByVal PrinterName As String) As Boolean

    End Function
#End Region

    Public Shared Property DefaultPrinterName() As String
        Get
            '\\ Go through the list of printers and return the default one
            Dim lpsRet As New System.Text.StringBuilder(256), chars As Integer = 256
            If GetDefaultPrinter(lpsRet, chars) Then

            End If
            Return lpsRet.ToString
        End Get
        Set(ByVal value As String)
            '\\ Go through the list of printers and if you find the one named as above make it the default
            If Not SetDefaultPrinter(value) Then
                Trace.WriteLine("Failed to set printer to : " & value)
            End If
        End Set
    End Property


End Class