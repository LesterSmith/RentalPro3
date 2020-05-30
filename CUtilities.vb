'****************************************
'* Purpose: Provides utility methods used 
'* by the whole application.
'*
'* Author:  Les Smith
'* Date Created: 07/18/2002 at 09:16:57
'* CopyRight:  InfoProGroup, Inc.
'****************************************
Imports System.Data.OleDb
#Region " Imports "
Imports System
Imports System.Windows.Forms.Application
Imports System.Text


#End Region

Public Class CUtilities
#Region " Class Level Variables "
   Public LastLineWrapped As Boolean


#End Region

#Region " Public Methods "
   Public Function GetAppPath() As String
      ' returns the path from which the exe is executing
      Return System.Reflection.Assembly.GetExecutingAssembly.Location()
   End Function
   Friend Function GetNowPrintString() As String
      ' Returns date/time string as mm/dd/yyyy at hh:mm:ss
      ' which is not necessaarily easy in .NET
      ' Note that MM is capitalized and dd/yyyy are not;
      ' that is the way .NET requires it to get the right
      ' value that VB6 Format returned!  Duh!!!!
      ' WHY??? I give up...
      Return Format(DateValue(Now()), "MM/dd/yyyy") & " at " & Format(Now(), "hh:mm:ss")
   End Function
   Friend Function CountOccurrences(ByVal rsExp As String, _
                                       ByVal rsStr As Object) As Integer
      ' Returns the number of occurrences of rsExp (expression)
      ' found in rsStr (string)
      ' Returns 0 of no occurrences found.
      Dim pPos As Integer
      Dim lPos As Integer
      Dim nPos As Integer
      Dim nFirst As Integer
      Dim lCnt As Integer

      Try

         pPos = 0 ' previous find
         lPos = 0 ' return position of right char
         nPos = 1 ' position of next right most char
         nFirst = 1
         lCnt = 0

         ' loop thru every char in string until we
         ' find the last occurrence
         Do
            lPos = InStr(nPos, rsStr, rsExp, 1)
            If lPos > 0 Then
               nPos = lPos + 1
               pPos = lPos
               lCnt = lCnt + 1
            Else
               Exit Do
            End If
         Loop

         Return lCnt
      Catch e As System.Exception
      End Try
   End Function
   Friend Function GetNextCommaDelimitedToken(ByRef TxPtr As Integer, ByRef TxLine As String, Optional ByVal Delimiter As String = ",") As String
      ' Return all characters preceding the next ","
      ' Picks up Left$ of string beginning at 'i' where
      Try
         Dim s As String
         Dim j As Integer = InStr(TxPtr, TxLine, Delimiter)

         If j > 0 Then
            s = Mid$(TxLine, TxPtr, j - TxPtr)
            TxPtr = j + 1
            Return s
         Else
            Return Mid$(TxLine, TxPtr)
         End If
      Catch ex As System.Exception
      End Try
   End Function
   Friend Function RepString(ByVal rsChar As String, ByVal riNbr As Integer) As String
      ' This method will return a string of riNbr of rsChar.
      ' It replaces the vb6 String$ Function.
      Try
         Dim i As Integer
         Dim s As New System.Text.StringBuilder()
         For i = 1 To riNbr
            s.Append(rsChar)
         Next
         Return s.ToString
      Catch
      End Try
   End Function

   Function IsALLAlpha(ByVal cS As String) As Boolean
      Dim i As Integer = 0
      Dim nc As String ' next character
      Dim AsciiVal As Integer

      If cS.Length = 0 Then Return False

      For i = 0 To cS.Length - 1
         nc = cS.Substring(i, 1)
         AsciiVal = Asc(nc)
         If Not ((AsciiVal > 64 And AsciiVal < 91) Or _
                (AsciiVal > 96 And AsciiVal < 123)) _
                Then
            Return False
         End If
      Next i
      Return True

      'While cS.Length > i
      '   CharTemp = temp.Substring(0, 1) ' get first character first
      '   temp = Right(temp, Len(temp) - 1)
      '   i += 1
      '   AsciiVal = Asc(CharTemp)

      '   If Not ((AsciiVal > 64 And AsciiVal < 91) Or (AsciiVal > 96 And AsciiVal < 123)) Then
      '      Return False
      '   End If
      'End While
      'Return True
   End Function

   Function IsDigit(ByVal cS As String) As Boolean
      ' Visual Basic replacement for Clipper isdigit() function.
      ' Call is retlogical = isdigit(cString)
      ' Returns True if first character of cString is digit,
      ' otherwise False.
      Dim cTemp As String
      Dim vTemp As Integer

      cTemp = cS.Substring(0, 1) ' get first character first
      If cTemp.Trim.Length = 0 Then ' cant pass Asc() an empty string
         Return False
      End If

      vTemp = Asc(cTemp)
      Return (vTemp > 47 AndAlso vTemp < 58)
      'If (vTemp > 47 And vTemp < 58) Then
      '   Return True
      'Else
      '   Return False
      'End If
   End Function




   Friend Function GetToken(ByRef srcline As String, _
                             Optional ByVal rsNonDelimiters As String = "", _
                             Optional ByVal rsDel As String = "N") _
                             As String


      '-----
      ' If rsDel = "N" then the rsNondelimiters is a list of non delimters
      ' which is added to a list of AN Chars (a-z, A-Z, 0-9), which are
      ' always assumed to be non delimiters.
      ' If rsDel="D" then rsNonDelimiters is the list of delimiters, anything
      ' else in the string is assumed to be non deliter.
      ' Get Next word from srcLine.  An alphanumeric and any character
      ' found in strDelimtrs is a valid char for the word.  i.e. a char
      ' which is not alphanumeric and not found in the delimiter string
      ' will terminate the word.  If space is not a delimiter it must be
      ' included in the strNonDelimitrs.
      ' Typicall call is:
      '     srcLine = GetToken(srcLine, " ().!" or
      '     srcLine = GetToken(srcLine, " ,") where space and comma are the delimiters.
      ' Any non alphanumeric and not in the " ().!" would terminate the string
      ' To include " in the set of allowable chars, concatenate chr(34) with the
      ' other non delimiters.
      ' If non delimiters are not supplied, dont compare for them
      ' and performance is increased...
      '-----
      Dim n_w As New StringBuilder() ' staging area for return string
      Dim FC As String ' first char of string
      Dim lsTemp As String
      Dim lsTemp2 As String
      Const AN_DIGITS = "abcdefghijklmnopqrstuvwxyz" & _
                        "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"
      Try
         If rsDel = "N" Then
            lsTemp2 = AN_DIGITS & rsNonDelimiters
         Else
            lsTemp2 = rsNonDelimiters
         End If

         Do While Trim$(srcline) <> ""
            FC = srcline.Substring(0, 1)
            If rsDel = "N" Then
               If lsTemp2.IndexOf(FC) = -1 Then
                  srcline = Mid(srcline, 2) ' save all but first char for next call
                  If n_w.Length > 0 Then
                     Return n_w.ToString
                  End If
               Else
                  n_w.Append(FC)
                  srcline = Mid(srcline, 2)
               End If
            Else
               If lsTemp2.IndexOf(FC) > 0 Then
                  srcline = Mid(srcline, 2) ' save all but first char for next call
                  If n_w.Length > 0 Then
                     Return n_w.ToString()
                  End If
               Else
                  n_w.Append(FC)
                  srcline = Mid(srcline, 2)
               End If
            End If
         Loop

         Return n_w.ToString

      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Function
   '   Friend Function MLCount(ByVal cStrng As String, Optional ByVal nL As Integer = 0) As Integer
   '      '` VB Replacement for Clipper MLCount Function.  
   '      '` It does handle word wrap, nL is the max char
   '      '' count per line.
   '      Dim nStptr As Integer, nLenStr As Integer, nLineCtr As Integer
   '      Dim sTemp As String
   '      Dim i As Integer
   '      Dim sdel As String
   '      Dim k As Short

   '      ' nStPtr is the pointer to position in cStrng

   '      Try
   '         If cStrng.IndexOf(vbCrLf) > -1 Then
   '            sdel = vbCrLf
   '         ElseIf cStrng.IndexOf(vbCr) > -1 Then
   '            sdel = vbCr
   '         ElseIf cStrng.IndexOf(vbLf) > -1 Then
   '            sdel = vbLf
   '         Else
   '            Return CInt(1)
   '         End If

   '         k = sdel.Length
   '         nStptr = 1
   '         nLenStr = cStrng.Length
   '         nLineCtr = 0

   '         While True
   '            ' If the pointer to the beginning of the next line
   '            ' is >= the length of the string, we are outta here!
   '            If nStptr >= nLenStr Then
   '               Return nLineCtr
   '               Exit Function
   '            End If

   '            ' Get the next line, not to exceed the length of nL
   '            ' if nL was greater than 0
   '            If nL > 0 Then
   '               sTemp = Mid$(cStrng, nStptr, nL)
   '               If sTemp.IndexOf(sdel) > -1 Then
   '                  ' there is a delimiter in the string
   '                  sTemp = sTemp.Substring(0, sTemp.IndexOf(sdel) + 1) ' - 1)
   '                  nStptr = nStptr + sTemp.Length + (k - 1)
   '               Else
   '                  ' new code to handle lines with no crlf
   '                  If sTemp.Length = nL Then
   '                     ' we have a full line left (at least)
   '                     ' check to see if the next char after stemp can be examined
   '                     If nStptr + nL + 1 <= nLenStr Then
   '                        ' check to see if the next char is a space
   '                        If Mid(cStrng, nStptr + nL, 1) = " " Then
   '                           nStptr += sTemp.Length
   '                           ' check to see if the next character is a special char
   '                           ' which should be printed immediately following the 
   '                           ' last word
   '                        Else
   '                           ' last word is not complete, truncate to preceding space
   '                           i = sTemp.LastIndexOf(" ")
   '                           ' truncate the partial word from the end
   '                           sTemp = Mid(sTemp, 1, i) '  - 1)
   '                           'set the pointer to start the next line at
   '                           'current start point + len(stemp)
   '                           nStptr += sTemp.Length
   '                        End If
   '                     Else
   '                        'set the pointer to start the next line at
   '                        'current start point + len(stemp)
   '                        nStptr += sTemp.Length
   '                     End If
   '                  Else
   '                     ' this is the last line, because the string is
   '                     ' shorter than the nL length
   '                     Return nLineCtr + 1
   '                     Exit Function
   '                  End If
   '               End If
   '            Else
   '               ' nL was supplied as 0 meaning we just look for CRLf
   '               nStptr = InStr(nStptr, cStrng, sdel) + k
   '            End If

   '            ' if the ptr = 2 then there was no crlf in the line
   '            If nStptr = k Then
   '               Return nLineCtr + 1
   '            End If

   '            nLineCtr += 1
   '            If nStptr + k - 1 > nLenStr Then
   '               Return nLineCtr
   '            End If
   '         End While
   '         Exit Function
   '      Catch e As System.Exception
   '         MsgBox("Error: " & e.Message, vbCritical, "MLCount")
   '      End Try
   '   End Function

   '   Friend Function MemoLine(ByVal cStrng As String, ByVal nLL As Integer, ByVal nL As Integer) As String
   '      '***************************************
   '      '* Name: MemoLine
   '      '* Purpose:
   '      '*   VB Replacement for Clipper MemoLine() Function.
   '      '*   Handles Word Wrap.  nLL is the max char/line.
   '      '*   Note that if the user asks for a line that is beyond the
   '      '*   end of the string, i.e. more lines than are in the string
   '      '*   unpredictable results will be returned, assuming we
   '      '*   return at all.  Therefore, MLCount() must be called
   '      '*   before calling MemoLine() and MemoLine must not be called
   '      '*   to return a line numbered higher than MLCount() returened.
   '      '*
   '      '* Parameters:
   '      '*   cStrng
   '      '*   nLL As Integer
   '      '*   nL As Integer
   '      '*
   '      '* Returns:
   '      '*
   '      '* Author: Les Smith
   '      '* Date Created: 11/10/1997
   '      '* Copyright: HHI Software
   '      '* Date Last Changed: to allow fetch of any line
   '      '* in word wrap.
   '      '* Liscened to InfoProGroup by Les Smith
   '      '***************************************



   '      Try
   '         Static nStptr As Long
   '         Dim i As Long
   '         Dim nTmpPtr As Long
   '         Dim sTemp As String
   '         Static j As Long
   '         Dim iSt As Long
   '         Dim sdel As String
   '         Dim k As Short

   '         If cStrng.IndexOf(vbCrLf) > -1 Then
   '            sdel = vbCrLf
   '         ElseIf cStrng.IndexOf(vbCr) > -1 Then
   '            sdel = vbCr
   '         ElseIf cStrng.IndexOf(vbLf) > -1 Then
   '            sdel = vbLf
   '         Else
   '            Return cStrng
   '         End If

   '         k = sdel.Length

   '         ' if NL is 1 > than J then
   '         ' this is a subsequent call to get the next
   '         ' line
   '         If j = 0 Then
   '            nStptr = 1
   '         End If
   '         If nL - j = 1 Then
   '            iSt = nL
   '         Else
   '            nStptr = 1
   '            iSt = 1
   '         End If

   '         LastLineWrapped = False

   '         If nStptr >= Len(cStrng) Then
   '            MemoLine = ""
   '            Exit Function
   '         End If

   '         ' Loop through the string until we find the requested line.
   '         For j = iSt To nL
   '            ' Get the next line, not to exceed the length of nLL
   '            ' if nL was greater than 0
   '            If nLL = 0 Then
   '               ' nL was supplied as 0 meaning we just look for vbCrLf
   '               i = InStr(nStptr, cStrng, sdel, 1)

   '               If i = 0 Then
   '                  ' no vbcrlf, return the whole remaining portion of string
   '                  MemoLine = Mid(cStrng, nStptr) 'Trim(Mid(cStrng, nStptr))

   '                  ' set the next ptr at the end of the string
   '                  ' in case the user calls for the next line, which
   '                  ' if mlcount worked properly, they should not do...
   '                  nStptr = Len(cStrng)
   '                  Exit Function
   '               ElseIf i = nStptr Then
   '                  ' the first chars in the current line are vbcrlf
   '                  nStptr += k
   '                  MemoLine = ""
   '                  If j < nL Then
   '                     GoTo BottomOfLoop
   '                  Else
   '                     Exit Function
   '                  End If
   '               Else
   '                  MemoLine = Mid(cStrng, nStptr, i - nStptr)
   '                  nStptr = i + k
   '                  If j < nL Then
   '                     GoTo BottomOfLoop
   '                  Else
   '                     Exit Function
   '                  End If
   '               End If
   '            Else
   '               ' user specified max length of lines to be returned,
   '               ' i.e. word wrap is called for...
   '               sTemp = Mid$(cStrng, nStptr, nLL)
   '               If sTemp.IndexOf(sdel) > -1 Then
   '                  ' there is a vbCrLf in the string
   '                  sTemp = Mid(sTemp, 1, sTemp.IndexOf(sdel) - (k - 1)) ' question - 1 [k-1]?
   '                  nStptr += sTemp.Length + k
   '                  MemoLine = sTemp
   '                  If j < nL Then
   '                     GoTo BottomOfLoop
   '                  Else
   '                     Exit Function
   '                  End If
   '               Else
   '                  ' no vbCrLf in string, find end of last full word
   '                  ' see if the line is shorter than the requested line
   '                  If Len(sTemp) < nLL Then
   '                     ' line is less than requested length,
   '                     ' we are at the end of the input string
   '                     ' set the pointer to the next line past the
   '                     ' end of the string
   '                     nStptr = cStrng.Length + 1
   '                     MemoLine = sTemp
   '                     Exit Function
   '                  Else
   '                     ' this is not the last line, .'. find the
   '                     ' last space in the line, assuming there is one...
   '                     i = sTemp.LastIndexOf(" ")

   '                     If i = 0 Then
   '                        ' there is no space in the line
   '                        MemoLine = sTemp
   '                        nStptr += sTemp.Length
   '                        If j < nL Then
   '                           GoTo BottomOfLoop
   '                        Else
   '                           LastLineWrapped = True
   '                           Exit Function
   '                        End If
   '                     Else
   '                        ' there is a space in the line
   '                        sTemp = Mid(sTemp, 1, i)
   '                        MemoLine = sTemp
   '                        nStptr += i
   '                        If j < nL Then
   '                           GoTo BottomOfLoop
   '                        Else
   '                           LastLineWrapped = True
   '                           Exit Function
   '                        End If
   '                     End If
   '                  End If
   '               End If
   '            End If
   'BottomOfLoop:
   '         Next j
   '      Catch ex As System.Exception
   '         StructuredErrorHandler(ex)
   '      End Try
   '   End Function
   Friend Overloads Function MLCount(ByVal cStrng As String, ByVal nL As Integer) As Integer
      '-----
      ' VB Replacement for Clipper MLCount Function
      ' It does handle word wrap, nL is the max char
      ' count per line.
      '-----
      Dim nStptr As Integer, nLenStr As Integer, nLineCtr As Integer
      Dim sTemp As String
      Dim i As Integer
      Dim ptr As Integer

      ' nStPtr is the pointer to position in cStrng

      Try
         nStptr = 1
         nLenStr = Len(cStrng)
         nLineCtr = 0

         While True
            ' If the pointer to the beginning of the next line
            ' is >= the length of the string, we are outta here!
            If nStptr >= nLenStr Then
               Return nLineCtr
               Exit Function
            End If

            ' Get the next line, not to exceed the length of nL
            ' if nL was greater than 0
            If nL > 0 Then
               sTemp = Mid$(cStrng, nStptr, nL)
               ptr = InStr(sTemp, vbCrLf)
               If ptr > 0 Then
                  ' there is a CRLF in the string
                  If ptr - 1 > 0 Then
                     sTemp = Mid(sTemp, 1, ptr - 1)
                  End If
                  nStptr += Len(sTemp) + 2
               Else
                  ' new code to handle lines with no crlf
                  If Len(sTemp) = nL Then
                     ' we have a full line left (at least)
                     i = sTemp.LastIndexOf(" ")
                     ' truncate the partial word from the end
                     If i > 0 Then
                        sTemp = Mid(sTemp, 1, i - 1)
                     End If
                     'set the pointer to start the next line at
                     'current start point + len(stemp)
                     nStptr = nStptr + Len(sTemp)
                  Else
                     ' this is the last line, because the string is
                     ' shorter than the nL length
                     Return nLineCtr + 1
                     Exit Function
                  End If
                  End If
            Else
               ' nL was supplied as 0 meaning we just look for CRLf
               nStptr = InStr(nStptr, cStrng, vbCrLf) + 2
            End If

            ' if the ptr = 2 then there was no crlf in the line
            If nStptr = 2 Then
               Return nLineCtr + 1
            End If

            nLineCtr = nLineCtr + 1
            If nStptr + 1 > nLenStr Then
               Return nLineCtr
            End If
         End While
         Exit Function
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Function

   Friend Overloads Function MemoLine(ByVal cStrng As String, ByVal nLL As Integer, ByVal nL As Integer) As String
      '***************************************
      '* Name: MemoLine
      '* Purpose:
      '*   VB Replacement for Clipper MemoLine() Function.
      '*   Handles Word Wrap.  nLL is the max char/line.
      '*   Note that if the user asks for a line that is beyond the
      '*   end of the string, i.e. more lines than are in the string
      '*   unpredictable results will be returned, assuming we
      '*   return at all.  Therefore, MLCount() must be called
      '*   before calling MemoLine() and MemoLine must not be called
      '*   to return a line numbered higher than MLCount() returened.
      '*
      '* Parameters:
      '*   cStrng
      '*   nLL As Integer
      '*   nL As Integer
      '*
      '* Returns:
      '*
      '* Author: Les Smith
      '* Date Created: 11/10/1997
      '* Copyright: HHI Software
      '* Date Last Changed: to allow fetch of any line
      '* in word wrap.
      '***************************************



      Try
         Static nStptr As Long
         Dim i As Long
         Dim nTmpPtr As Long
         Dim sTemp As String
         Static j As Long
         Dim iSt As Long

         ' if NL is 1 > than J then
         ' this is a subsequent call to get the next
         ' line
         If j = 0 Then
            nStptr = 1
         End If
         If nL - j = 1 Then
            iSt = nL
         Else
            nStptr = 1
            iSt = 1
         End If

         LastLineWrapped = False

         If nStptr >= Len(cStrng) Then
            MemoLine = ""
            Exit Function
         End If

         ' Loop through the string until we find the requested line.
         For j = iSt To nL
            ' Get the next line, not to exceed the length of nLL
            ' if nL was greater than 0
            If nLL = 0 Then
               ' nL was supplied as 0 meaning we just look for vbCrLf
               i = InStr(nStptr, cStrng, vbCrLf, 1)

               If i = 0 Then
                  ' no vbcrlf, return the whole remaining portion of string
                  MemoLine = Mid(cStrng, nStptr) 'Trim(Mid(cStrng, nStptr))

                  ' set the next ptr at the end of the string
                  ' in case the user calls for the next line, which
                  ' if mlcount worked properly, they should not do...
                  nStptr = Len(cStrng)
                  Exit Function
               ElseIf i = nStptr Then
                  ' the first chars in the current line are vbcrlf
                  nStptr = nStptr + 2
                  MemoLine = ""
                  If j < nL Then
                     GoTo BottomOfLoop
                  Else
                     Exit Function
                  End If
               Else
                  MemoLine = Mid(cStrng, nStptr, i - nStptr) 'Trim(Mid(cStrng, nStptr, i - nStptr))
                  nStptr = i + 2
                  If j < nL Then
                     GoTo BottomOfLoop
                  Else
                     Exit Function
                  End If
               End If
            Else
               ' user specified max length of lines to be returned,
               ' i.e. word wrap is called for...
               sTemp = Mid$(cStrng, nStptr, nLL)
               If InStr(sTemp, vbCrLf) > 0 Then
                  ' there is a vbCrLf in the string
                  sTemp = Mid(sTemp, 1, InStr(sTemp, vbCrLf) - 1)
                  nStptr = nStptr + Len(sTemp) + 2
                  MemoLine = sTemp 'Trim(sTemp)
                  If j < nL Then
                     GoTo BottomOfLoop
                  Else
                     Exit Function
                  End If
               Else
                  ' no vbCrLf in string, find end of last full word
                  ' see if the line is shorter than the requested line
                  If Len(sTemp) < nLL Then
                     ' line is less than requested length,
                     ' we are at the end of the input string
                     ' set the pointer to the next line past the
                     ' end of the string
                     nStptr = Len(cStrng) + 1
                     MemoLine = sTemp
                     Exit Function
                  Else
                     ' this is not the last line, .'. find the
                     ' last space in the line, assuming there is one...
                     i = InStrRev(sTemp, " ")

                     If i = 0 Then
                        ' there is no space in the line
                        MemoLine = sTemp 'Trim(sTemp)
                        nStptr = nStptr + Len(sTemp) '+ 1
                        If j < nL Then
                           GoTo BottomOfLoop
                        Else
                           LastLineWrapped = True
                           Exit Function
                        End If
                     Else
                        ' there is a space in the line
                        sTemp = Mid(sTemp, 1, i)
                        MemoLine = sTemp 'Trim(sTemp)
                        nStptr = nStptr + i
                        If j < nL Then
                           GoTo BottomOfLoop
                        Else
                           LastLineWrapped = True
                           Exit Function
                        End If
                     End If
                  End If
               End If
            End If
BottomOfLoop:
         Next j
      Catch ex As System.Exception
         StructuredErrorHandler(ex)
      End Try
   End Function

   Function IsALLNumeric(ByVal cS As String) As Boolean
      Dim i As Integer = 0
      Dim nc As String ' next character
      Dim AsciiVal As Integer

      If cS.Length = 0 Then Return False

      For i = 0 To cS.Length - 1
         nc = cS.Substring(i, 1)
         If Not (nc >= "0" And nc <= "9") Then
            Return False
         End If
      Next i
      Return True
   End Function
   Public Overloads Function CountSpacesBeforeFirstChar(ByVal psIN As String) As Integer
      ' Count the number of spaces before first
      ' non space character in a source line
      Dim iSpCnt As Integer

      For iSpCnt = 0 To Len(psIN) - 1
         If Mid$(psIN, iSpCnt + 1, 1) <> " " Then
            CountSpacesBeforeFirstChar = iSpCnt
            Exit Function
         End If
      Next iSpCnt
      CountSpacesBeforeFirstChar = iSpCnt

   End Function
   Public Overloads Function CountSpacesBeforeFirstChar(ByVal psIN As String, ByVal piSt As Integer) As Integer
      ' Count the number of spaces before first
      ' non space character in a source line, 
      ' beginning at piST
      Dim iSpCnt As Integer

      For iSpCnt = piSt To Len(psIN) - 1
         If Mid$(psIN, iSpCnt + 1, 1) <> " " Then
            CountSpacesBeforeFirstChar = iSpCnt - piSt
            Exit Function
         End If
      Next iSpCnt
      CountSpacesBeforeFirstChar = iSpCnt - piSt

   End Function


#End Region

End Class