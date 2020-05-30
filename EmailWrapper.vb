Imports System.IO
'Imports Utilities.GIS.Interfaces
Imports System.Net.Mail
Imports System.Net
 
Namespace HHISoftware
    Public Class EmailWrapperLib
        Private _password As String = "Prayer09"

        Private _sb As System.Text.StringBuilder
        Private _logPath As String = String.Empty
        Private _mailServer As String = String.Empty
        Private _msgFrom As String = String.Empty

        Public Property sb() As System.Text.StringBuilder
            Get
                Return _sb
            End Get
            Set(ByVal Value As System.Text.StringBuilder)
                _sb = Value
            End Set
        End Property

        Public Property logPath() As String
            Get
                Return _logPath
            End Get
            Set(ByVal Value As String)
                _logPath = Value
            End Set
        End Property

        Public Property mailServer() As String
            Get
                Return _mailServer
            End Get
            Set(ByVal Value As String)
                _mailServer = Value
            End Set
        End Property

        Public Property msgFrom() As String
            Get
                Return _msgFrom
            End Get
            Set(ByVal Value As String)
                _msgFrom = Value
            End Set
        End Property

        ''' <summary>

        ''' Any values passed to the constructor can be changed by
        ''' changing the respective property.
        ''' </summary>
        ''' <param name="msg"></param>
        ''' <param name="subject"></param>
        ''' <param name="msgTO"></param>
        ''' <param name="cc"></param>
        ''' <param name="attach"></param>
        ''' <remarks></remarks>
        Public Sub SendEmail( _
           ByVal msg As String, _
           ByVal subject As String, _
           ByVal msgTO As String, _
           Optional ByVal cc As String = "", _
           Optional ByVal attach As String = "")

            Try
                Dim message As New MailMessage

                message.From = New MailAddress(msgFrom)
                AddAddressesToMessageObject(msgTO, message, "to")

                message.Subject = subject
                message.Body = msg
                If Not String.IsNullOrEmpty(cc) Then
                    AddAddressesToMessageObject(cc, message, "cc")
                End If
                'check for an attachments this will only work with 1 attachments.
                If attach IsNot Nothing AndAlso IO.File.Exists(Path.Combine(CutePDFFilePath, attach)) Then
                    message.Attachments.Add(New Attachment(Path.Combine(CutePDFFilePath, attach)))
                End If
                'message.DeliveryNotificationOptions = DeliveryNotificationOptions.

                Dim sc As New SmtpClient()
                sc.Port = 587
                sc.Host = mailServer
                sc.EnableSsl = True
                sc.Credentials = New NetworkCredential(msgFrom, _password)
                'sc.Port = 465
                sc.Send(message)
                MsgBox("Your email was sent.")
            Catch ex As Exception
                StructuredErrorHandler(ex)
            End Try
        End Sub

        ''' <summary>
        ''' Parses a mail string and returns a MailAddress object.
        ''' You can no longer have ; in to/from string object.
        ''' </summary>
        ''' <param name="emailAddresses"></param>
        ''' <param name="mm"></param>
        ''' <remarks></remarks>
        Private Sub AddAddressesToMessageObject(ByVal emailAddresses As String, ByVal mm As MailMessage, ByVal type As String)

            If emailAddresses.Trim.Length = 0 Then Exit Sub

            Dim addresses() As String = emailAddresses.Split(";"c)

            For Each address As String In addresses
                If Not String.IsNullOrEmpty(address) AndAlso address.Trim.Length > 0 Then
                    Select Case type.ToLower
                        Case "to" : mm.To.Add(address)
                        Case "cc" : mm.CC.Add(address)
                    End Select
                End If
            Next
        End Sub

        ''' <summary>
        ''' The stringbuilder sb has an accumulation of all exceptions created
        ''' during the processing of the batch.  Build a .txt file for attachment
        ''' to email that is going to the client.
        ''' </summary>
        ''' <returns>String</returns>
        ''' <remarks></remarks>
        Public Function CreateEMailAttachment() As String
            If sb.Length > 0 Then
                Dim fn As String = Path.Combine(logPath, "Exceptions_" & Format(Now, "yyyyMMddHHmmss") & ".txt")
                Write2File(sb.ToString, fn)
                sb = New System.Text.StringBuilder(5000)
                Return fn
            Else
                Return String.Empty
            End If
        End Function

        Public Sub AddMessageToEmailAttachment(ByVal msg As String)
            sb.Append(msg & vbCrLf)
        End Sub

        ''' <summary>
        ''' logPath is where an email attachment can be built for you
        ''' when calling the AddMessageToEmailAttachment.
        ''' </summary>
        ''' <param name="emailServer"></param>
        ''' <param name="msgFrom"></param>
        ''' <param name="attachmentPath"></param>
        ''' <remarks></remarks>
        Public Sub New(ByVal emailServer As String, ByVal msgFrom As String, ByVal attachmentPath As String)
            logPath = attachmentPath
            mailServer = emailServer
            Me.msgFrom = msgFrom
            sb = New System.Text.StringBuilder(5000)
        End Sub

        ''' <summary>
        ''' This method will either create a new file or append to it and write
        ''' the passed msg, adding a vbcrlf to it.  The file is then closed so that if we
        ''' crash, we always have the audtit file for as far as we have gone.
        ''' On 2/25/2008 changed from Sub to Function to return a false if unsuccessful
        ''' write.  Cannot email from here b/c we have no email object in this module.
        ''' </summary>
        ''' <param name="msg" ></param>
        ''' <param name="filePath" ></param>
        ''' <remarks></remarks>
        Public Overloads Function Write2File(ByVal msg As String, ByVal filePath As String) As Boolean
            Write2File(msg, filePath, False)
        End Function

        ''' <summary>
        ''' This method will either create a new file or append to it and write
        ''' the passed msg, adding a vbcrlf to it.  The file is then closed so that if we
        ''' crash, we always have the audtit file for as far as we have gone.
        ''' Prefixes msg with datetime if includTime=true
        ''' On 2/25/2008 changed from Sub to Function to return a false if unsuccessful
        ''' write.  Cannot email from here b/c we have no email object in this module.
        ''' </summary>
        ''' <param name="msg"></param>
        ''' <param name="filePath"></param>
        ''' <param name="includeTime"></param>
        ''' <remarks></remarks>
        Public Overloads Function Write2File(ByVal msg As String, ByVal filePath As String, ByVal includeTime As Boolean) As Boolean
            Using fs As New FileStream(filePath, FileMode.Append, FileAccess.Write, FileShare.ReadWrite)
                Try
                    Dim sw As StreamWriter = New StreamWriter(fs)
                    Dim timeStamp As String = String.Empty

                    If includeTime Then
                        timeStamp = "AT: " & Now.ToString & "  "
                    End If

                    sw.WriteLine(timeStamp & msg)
                    sw.Flush()
                    sw.Close()
                    fs.Close()
                Catch ex As System.Exception
                    If fs IsNot Nothing Then fs.Close()
                    Throw ex
                End Try
            End Using
            Return True
        End Function
    End Class
End Namespace
