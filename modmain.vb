Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports System.IO
Module modMain
    'Public colItems As New Collection()
    Public fMainForm As frmRental 'frmMain
    'Public oProgram As CProgram
    Public giChildCount As Short
    Public Const SETTINGS As String = "SETTINGS"
    Public ConnectString As String
    'Public colItems As colItems
    Public AppPath As String
    Public UserName As String
    Public DatabaseName As String
    Public Const RENTALPRO = "RentalProV2"
    ' currently support customers
    Public Const PIONEER = "PIONEER"
    Public Const RELIABLE = "RELIABLE"

    ' this boolean will be set to true by frmCustomers
    ' if items are cancelled and that will cause frmRental
    ' to reload the grid
    Public ReloadRentalGrid As Boolean = False
    Public Const INVOICE_DETAIL_RECORD = 15
    Public Const DEPOSIT_RECORD = 25
    Public Const TAX_RECORD = 35
    Public Const DELIVERY_RECORD = 45
    Public Const PICKUP_RECORD = 46
    Public Const AMTPAID_RECORD = 55
    Public Const REFUND_AT_CHECKIN_RECORD = 65
    Public Const BALANCE_DUE_AFTER_CHECKIN_RECORD = 75
    Public Const CASH_ON_ACCOUNT_RECORD = 67
    Public Const CREDIT_MEMO_RECORD = 66
    Public Const DEBIT_MEMO_RECORD = 68

    Public Const DAILY = "Daily"
    Public Const HALF_DAY = "Half Day"
    Public Const WEEKLY = "Weekly"
    Public Const MONTHLY = "Monthly"
    Public Const WEEK_END = "Week End"
    Public Const SALE = "Sale"
    Public Const RENT = "Rent"
    Public Const HOURLY = "Hourly"
    Public Const RERENT = "ReRent"
    Public Const REPORT_TITLE_FONT_SIZE = 16
    ' company config variables
    Public ReportName As String
    Public CorporateName
    Public Address1 As String
    Public Address2 As String
    Public City As String
    Public State As String
    Public Zip As String
    Public Phone As String
    Public Fax As String
    Public EMail As String
    Public EmailServer As String
    Public EmailGreeting As String
    Public EmailBodyStart As String
    Public EmailSubject As String
    Public EmailBody As String
    Public CostBasis As String
    Public UseDeposits As Boolean
    Public UseHourlyRates As Boolean
    Public UseWeekEndRates As Boolean
    Public TaxRate As Single
    Public AccountingBasis As String
    Public PrintInitialsOnly As Boolean
    Public HoursPerMonth As Integer
    Public DaysPerMonth As Integer
    Public CalcByMonth As Boolean
    Public GraceHoursForDayRent As Integer
    Public GraceHoursForHalfDayRent As Integer
    Public UseHalfDays As Boolean = True ' false means 4 hour half days
    Public AutoCalcOn As Boolean = True ' turns autocalc on/off
    Public MonthlyBreakDays As Integer = 17 ' over this amount moves to monthly rntal
    Public WeeklyBreakDays As Integer = 3
    Public StoreOpenHour As Integer = 8 ' 8 am
    Public AfterNoonRentBeginHour As Integer = 13 '1 pm
    Public FridayWeekEndStartHour As Integer = 16 ' 4pm
    Public CalcBestRate As Boolean = False
    Public CustomerEmail As String
    Public CutePDFFilePath As String
    Public CustomerNameForInvoiceFile As String
    Public EmailSSL As Boolean = False
    Public EmailPort As Integer = 0

    ''' <summary>
    ''' Limits the size of a string to the length of cnt.
    ''' </summary>
    ''' <param name = "str"></param>
    ''' <param name = "cnt"></param>
    ''' <returns>String</returns>
    Public Function LS(ByVal str As String, ByVal cnt As Short) As String
        If str.Length > cnt Then
            Return str.Substring(0, cnt)
        Else
            Return str
        End If
    End Function

    Public Sub PositionForm(ByRef oFrm As Form)

        Try
            oFrm.Left = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width - oFrm.Width
            oFrm.Top = Screen.PrimaryScreen.Bounds.Height / 4
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub
    Public Sub ComboAutoSearch(ByRef cbo As System.Windows.Forms.ComboBox, ByRef KeyAscii As Short, Optional ByRef CompareMode As CompareMethod = CompareMethod.Text)
        ' Called from KeyPress Event of a a standard combo box
        ' requires the cbo and the KeyAscii of the Keypress event to be passed
        ' as parameters
        Dim index As Integer
        Dim Text As String

        ' if user pressed a control key, do noting
        If KeyAscii <= 32 Then Exit Sub

        ' produce new text, cancel automatic key processing
        Text = Left(cbo.Text, cbo.SelectionStart) & Chr(KeyAscii) & Mid(cbo.Text, cbo.SelectionStart + 1 + cbo.SelectionLength)
        KeyAscii = 0

        ' search the current item in the list
        For index = 0 To cbo.Items.Count - 1
            If InStr(1, VB6.GetItemString(cbo, index), Text, CompareMode) > 1 Then
                ' we've found a match
                cbo.SelectedIndex = index
                Exit For
            End If
        Next

        ' if no matching item
        If index = cbo.Items.Count Then
            cbo.Text = Text
        End If

        ' highlight trailing chars in the edit area
        cbo.SelectionStart = Len(Text)
        cbo.SelectionLength = 9999
    End Sub

    Public Sub IncrementChildCount(ByRef roFrm As Object)
        'roFrm.Move giChildCount * 300, giChildCount * 300
        giChildCount = giChildCount + 1
    End Sub

    Public Sub DecrementChildCount()
        giChildCount = giChildCount - 1
    End Sub

    Sub CenterForm(ByRef oFrm As System.Windows.Forms.Form)
        oFrm.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) / 2) - (VB6.PixelsToTwipsX(oFrm.Width) / 2))
        oFrm.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) / 2) - (VB6.PixelsToTwipsY(oFrm.Height) / 2))
    End Sub
    ''' <summary>
    ''' Return $nnn without pennies.
    ''' </summary>
    ''' <param name = "curAmt"></param>
    ''' <returns>String</returns>
    Public Function FormatDollars(ByVal curAmt As Decimal) As String
        Return Format(MND(curAmt), "$#,##0")
    End Function
    ''' <summary>
    '''    ''' Return formatted currency amt with $ and cents.
    ''' </summary>
    ''' <param name = "curAmt"></param>
    ''' <returns>String</returns>
    Public Function FormatCurrency(ByVal curAmt As Object) As String
        Try
            Try
                If curAmt = String.Empty Then curAmt = 0
            Catch
            End Try
            Return VB6.Format(curAmt, "$#,##0.00")
        Catch
            Return "$0.00"
        End Try
    End Function
    ''' <summary>
    ''' Returns formatted currency without $ sign.
    ''' </summary>
    ''' <param name = "curAmt"></param>
    ''' <returns>String</returns>
    Public Function FormatNDS(ByVal curAmt As Decimal) As String
        Return Format(MND(curAmt), "#,##0.00")
    End Function

    Public Function UnFormat(ByRef rsAmt As String) As Decimal
        UnFormat = Val(Replace(Replace(rsAmt, "$", ""), ",", ""))
    End Function

    Friend Function GetToken(ByRef srcline As String, _
                         ByVal rsNonDelimiters As String) _
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
        Dim n_w As String ' staging area for return string
        Dim FC As String ' first char of string
        Dim lsTemp As String
        Dim lsTemp2 As String
        Const AN_DIGITS = "abcdefghijklmnopqrstuvwxyz" & _
                          "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"
        Try
            n_w = ""
            lsTemp2 = AN_DIGITS & rsNonDelimiters

            Do While Trim$(srcline) <> ""
                FC = Mid(srcline, 1, 1)
                If InStr(lsTemp2, FC) = 0 Then
                    srcline = Mid(srcline, 2) ' save all but first char for next call
                    If Trim$(n_w) <> "" Then
                        GetToken = n_w
                        Exit Function
                    End If
                Else
                    n_w = n_w & FC
                    srcline = Mid(srcline, 2)
                End If
            Loop

            GetToken = n_w
            Exit Function
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
            GetToken = n_w
        End Try
    End Function

    Friend Function GetToken(ByRef srcline As String, _
                            ByVal rsNonDelimiters As String, _
                            ByVal rsDel As String) _
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
        Dim n_w As String ' staging area for return string
        Dim FC As String ' first char of string
        Dim lsTemp As String
        Dim lsTemp2 As String
        Const AN_DIGITS = "abcdefghijklmnopqrstuvwxyz" & _
                          "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"
        Try
            n_w = ""
            lsTemp2 = rsNonDelimiters

            Do While Trim$(srcline) <> ""
                FC = Mid(srcline, 1, 1)
                If InStr(lsTemp2, FC) > 0 Then '(lsTemp2 Like lsTemp) Then
                    srcline = Mid(srcline, 2) ' save all but first char for next call
                    If Trim$(n_w) <> "" Then
                        GetToken = n_w
                        Exit Function
                    End If
                Else
                    n_w = n_w & FC
                    srcline = Mid(srcline, 2)
                End If
            Loop

            GetToken = n_w
            Exit Function
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
            GetToken = n_w
        End Try
    End Function

    Public Sub FixGridCaptions(ByRef dt As DataTable)
        Dim i As Short
        Dim lsTemp As String
        Dim dr As DataRow
        Dim dc As DataColumn

        For Each dr In dt.Rows
            For Each dc In dt.Columns
                dc.Caption = MakeCaptionPretty(dc.Caption)
            Next
        Next
    End Sub
    Public Function MakeCaptionPretty(ByVal s As String) As String
        ' column_name_test
        ' make first char upper case
        Dim i As Short

        s = s.Substring(0, 1).ToUpper & s.Substring(1)
        Do
            If s.IndexOf("_") = 0 Then Return s
            i = s.IndexOf("_")
            s = s.Substring(i - 1) & " " & s.Substring(i + 1).ToUpper & s.Substring(i - 2)
        Loop

    End Function


    ''' <summary>
    ''' Return 0 if object is null, else decimal value
    ''' </summary>
    ''' <param name = "o"></param>
    ''' <returns>Decimal</returns>
    Public Function MND(ByVal o As Object) As Decimal
        If IsDBNull(o) Then
            Return 0
        Else
            Return CType(o, Decimal)
        End If
    End Function


    ''' <summary>
    ''' Return 0 if null, else integer value of object.
    ''' </summary>
    ''' <param name = "i"></param>
    ''' <returns>Integer</returns>
    Public Function MNI(ByVal i As Object) As Integer
        If IsDBNull(i) Then
            Return 0
        Else
            Return CType(i, Integer)
        End If
    End Function
    ''' Return String if object is not null, else return empty.string
    Public Function MNS(ByVal s As Object) As String
        If IsDBNull(s) Then
            Return String.Empty
        Else
            Return CType(s, String)
        End If
    End Function
    Public Function MNB(ByVal o As Object) As Boolean
        If IsDBNull(o) Then
            Return False
        Else
            Return (CType(o, Boolean))
        End If
    End Function

    ''' <summary>
    ''' Returns date string as MM/dd/yy hh:mm tt
    ''' </summary>
    ''' <param name = "dte"></param>
    ''' <returns>String</returns>
    Public Function FMDNS(ByVal dte As Object) As String
        If IsDBNull(dte) OrElse Not IsDate(dte) Then
            Return String.Empty
        Else
            Return Format(dte, "MM/dd/yy hh:mm tt")
        End If
    End Function

    ''' <summary>
    ''' Return date if object is not null, else return empty string
    ''' </summary>
    ''' <param name = "d"></param>
    ''' <returns>String</returns>
    Public Function MNDS(ByVal d As Object) As String
        If IsDBNull(d) Then
            Return String.Empty
        Else
            Return CType(d, Date)
        End If
    End Function

    ''' <summary>
    ''' Return single if not null.
    ''' </summary>
    ''' <param name = "bybal"></param>
    ''' <returns>Single</returns>
    Public Function MNSng(ByVal d As Object) As Single
        If IsDBNull(d) Then
            Return 0.0
        Else
            Return CSng(d)
        End If
    End Function

    Public Function GetAppPath() As String
        ' returns the path from which the exe is executing
        Return System.Reflection.Assembly.GetExecutingAssembly.Location()
    End Function
    ''' <summary>
    ''''' Get Database to open.
    ''' </summary>
    ''' <returns>String</returns>
    Public Function SelectDatabase(ByRef od As OpenFileDialog, Optional ByVal Required As Boolean = False) As String
        With od
            Do
                .Filter = "Database Files (*.mdb)|*.mdb"
                .ShowDialog()
                If .FileName.Length = 0 OrElse Dir(.FileName) = String.Empty Then
                    If Required Then
                        Dim sMsg As String
                        Dim iRV As Integer
                        sMsg = "You must select a database so that RentalPro can run." & Chr(10)
                        sMsg &= " " & Chr(10)
                        sMsg &= "Click Ok to select a database or Cancel to Quit." & Chr(10)
                        sMsg &= "" & Chr(10)
                        iRV = MsgBox(sMsg, CType(49, Microsoft.VisualBasic.MsgBoxStyle), "Select Database")

                        If iRV = 1 Then
                            ' Ok Code goes here
                        Else
                            ' Cancel code goes here
                            Return String.Empty
                        End If
                    Else
                        Return String.Empty
                    End If
                Else
                    SaveSetting(modMain.RENTALPRO, modMain.SETTINGS, "DBNAME", .FileName)
                    Return .FileName
                End If
            Loop
        End With
    End Function


    Public Sub Main()
        Dim Start As Integer
        Dim oCF As New CConfig()
        Dim oFrm As New frmRental()

        Try
            If PrevInstance() Then Exit Sub
            UserName = Environ("COMPUTERNAME")
            DatabaseName = GetSetting(RENTALPRO, SETTINGS, "DBNAME", "")
            If DatabaseName.Length = 0 Then
                DatabaseName = SelectDatabase(oFrm.OpenFileDialog1)
                If DatabaseName.Length = 0 Then
                    Exit Sub
                End If
            End If


            Dim oRES As CTransaction
            AppPath = GetAppPath()
            AppPath = AppPath.Substring(0, AppPath.LastIndexOf("\"))
            ConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabaseName

            Dim oFS As New frmSplash()
            oCF.GetConfig()
            oFS.Show()
            Start = VB.Timer()

            oRES = New CTransaction()
            Call oRES.RemoveTempReservation("")
            Do While VB.Timer() - Start < 2
                System.Windows.Forms.Application.DoEvents()
            Loop
            oFS.Close()
            System.Windows.Forms.Application.DoEvents()

            fMainForm = oFrm ' so frmCustomers can cause frmRental grid to refresh
            'oFrm.ShowDialog()
            Application.Run(oFrm)
        Catch ex As System.Exception
            StructuredErrorHandler(ex)
        End Try
    End Sub

    ' Visual Basic .NET
    Function PrevInstance() As Boolean
        ' Debug.Assert(UBound(Diagnostics.Process.GetProcessesByName(Diagnostics.Process.GetCurrentProcess.ProcessName)) > 0, "It worked")
        'If UBound(Diagnostics.Process.GetProcessesByName(Diagnostics.Process.GetCurrentProcess.ProcessName)) > 0 Then
        'Return True
        'Else
        Return False
        'End If
    End Function

    Public Sub StructuredErrorHandler(ByVal ex As System.Exception, Optional ByVal ShowMsgBox As Boolean = True)
        Try
            Dim msg As String = "Error------" & vbCrLf & _
               ex.Message & vbCrLf & _
               "Error Type-----" & vbCrLf & ex.GetType().ToString & vbCrLf & _
               "Error Details-----" & vbCrLf & ex.ToString
            If ShowMsgBox Then
                MsgBox(msg, MsgBoxStyle.Exclamation)
            End If

            Try
                Dim st As New System.Diagnostics.StackTrace(True)
                msg &= "Stack Trace------" & vbCrLf
                msg &= st.ToString
            Catch
            End Try
            WriteErrLog(msg)
        Catch
        End Try
    End Sub
    Public Sub StructuredErrorHandler(ByVal err As String)
        Try
            Dim msg As String = "Error------" & vbCrLf & err & vbCrLf
            MsgBox(msg, MsgBoxStyle.Exclamation)
            WriteErrLog(msg)
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Write a text file to the exe path for logging the error.
    ''' </summary>
    ''' <param name = "ex"></param>
    Public Sub WriteErrLog(ByRef s As String)
        Try
            Dim h As Integer = FreeFile()
            Dim ap As String = GetErrFileName()
            FileOpen(h, ap, OpenMode.Output)
            Print(h, s)
            FileClose(h)
        Catch
        End Try
    End Sub
    ''' <summary>
    ''' returns file to write error log.
    ''' </summary>
    ''' <returns>String</returns>
    Public Function GetErrFileName() As String
        Try
            Dim AppPath As String = GetAppPath()
            AppPath = AppPath.Substring(0, AppPath.LastIndexOf("\")) & "\ErrLogs"
            Dim di As New DirectoryInfo(AppPath)
            If Not di.Exists Then
                di.Create()
            End If
            Dim errFile As String = AppPath & "\ErrLog_" & Format(Now, "MMddyyyyHHmmss") & ".txt"
            Return errFile
        Catch
        End Try
    End Function

End Module