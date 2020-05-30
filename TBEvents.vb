Imports System.Windows.Forms
Namespace VBNetCommander
   Public Class TBEvents
      Private Shared Function CkKeyPressNumeric(ByVal KeyAscii As Integer, ByVal roTB As TextBox) As Integer
         ' allow 0-9,., Back, Del,-,Ins, and / if in tag format
         Try
            Dim retValue As Integer = KeyAscii
            If KeyAscii = Keys.Back OrElse _
               KeyAscii = Keys.Insert OrElse _
               KeyAscii = Keys.Delete OrElse _
               KeyAscii = 46 OrElse _
               (KeyAscii >= Keys.D0 AndAlso KeyAscii <= Keys.D9) Or _
               KeyAscii = 45 OrElse _
               KeyAscii = 46 _
               Then
               If roTB.SelectionLength = 0 Then
                  Dim idx As Integer = roTB.Text.IndexOf(".")
                  If idx > -1 Then
                     If roTB.Text.Substring(idx).Length > 1 Then
                        SendKeys.Send("{TAB}")
                        Return 0
                     End If
                  End If
               Else
                  roTB.Text = ""
               End If
               Return retValue
            End If
         Catch
         End Try
         Return 0
      End Function

      Public Function UnFmt_T_B(ByVal roTB As TextBox) As Object
         Dim dec As Decimal = 0
         Try
            Dim text As String = roTB.Text
            text = text.Replace("$", String.Empty)
            text = text.Replace(",", String.Empty)
            text = text.Replace("(", String.Empty)
            text = text.Replace(")", String.Empty)
            text = text.Replace("%", String.Empty)
            dec = CDbl(text)
            If InStr(roTB.Text, "%") Then
               dec /= 100
            End If
            If InStr(roTB.Text, "(") > 0 And InStr(roTB.Text, ")") > 0 Then
               dec *= -1
            End If
         Catch
         End Try
         Return dec
      End Function

      Public Function Fmt_T_B(ByVal roTB As TextBox) As String
         On Error Resume Next
         Dim text As Single = CDbl(roTB.Text)
         Dim fmt As String = roTB.Tag
         Return text.ToString(fmt)
      End Function

      Public Function Fmt_D_F(ByVal rsTxt As Object, ByVal roTB As TextBox) As String
         On Error Resume Next

         If InStr(1, roTB.Tag, ";", 1) > 0 Then
            If InStr(rsTxt, "-") Then

               Fmt_D_F = Format$(Replace(rsTxt, "-", ""), Mid$(roTB.Tag, InStr(roTB.Tag, ";") + 1))
            Else
               Fmt_D_F = Format$(rsTxt, Microsoft.VisualBasic.Left(roTB.Tag, InStr(roTB.Tag, ";") - 1))
            End If
         ElseIf InStr(1, roTB.Tag, "%", 1) > 0 Then
            Fmt_D_F = Format$(rsTxt, roTB.Tag)
         Else
            Fmt_D_F = Format$(rsTxt, roTB.Tag)
         End If
      End Function

   End Class
End Namespace