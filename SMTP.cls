VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SMTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mMath As Math
Private mMem As Memory
Private mStrings As Strings
Private mComputer As Computer

Private mWinsock As Winsock

Private mcolRecipientFrom As Collection
Private mcolRecipientTo As Collection
Private mcolRecipientCC As Collection
Private mcolRecipientBCC As Collection
Private mcolReplyTo As Collection
Private mstrSubject As String
Private mstrBody As String

Private mstrSmtpServer As String
Private mstrSmtpServerIP As String
Private mlngSmtpPort As Long

Private mlngLocalPort As Long
Private mlngSocketState As WinsockStateConstants

Private mobjLogViewer As Object

Private Sub Class_Initialize()
    Set mMath = New Math
    Set mMem = New Memory
    Set mStrings = New Strings
    Set mComputer = New Computer
        
    Set mcolRecipientFrom = New Collection
    Set mcolRecipientTo = New Collection
    Set mcolRecipientCC = New Collection
    Set mcolRecipientBCC = New Collection
    Set mcolReplyTo = New Collection
    
    mstrSmtpServer = mComputer.Name
    mlngSmtpPort = 25
    mlngLocalPort = 1001
    
    Set mWinsock = New Winsock
End Sub

Public Function GetLastError(nErrorCode As Long, lpBuffer As String)
    nErrorCode = mWinsock.GetLastError
    mWinsock.FormatErrorMessage nErrorCode, lpBuffer
End Function

Public Property Get SmtpServer() As String
    SmtpServer = mstrSmtpServer
End Property

Public Property Let SmtpServer(vData As String)
    mstrSmtpServer = vData
End Property

Public Property Get SmtpPort() As Long
    SmtpPort = mlngSmtpPort
End Property

Public Property Let SmtpPort(vData As Long)
    mlngSmtpPort = vData
End Property

Public Property Get SmtpServerIP() As String
    SmtpServerIP = mstrSmtpServerIP
End Property

Public Property Set LogViewer(vData As Object)
    Set mobjLogViewer = vData
End Property

Public Function Disconnect() As Long
    mWinsock.Shutdown
    SendListen 30
    If mWinsock.IsConnected Then
        mWinsock.CloseSocket
    End If
    
    Set mWinsock = Nothing
    Set mWinsock = New Winsock
    
    mlngSocketState = sckClosed
End Function

Public Function Connect() As Long
    Dim lngRetVal As Long
    
    Connect = 0
    
    mlngSocketState = sckConnecting
    Connect = mWinsock.Startup(&H202)
    If Connect = SOCKET_ERROR Then
        mlngSocketState = sckError
        Exit Function
    End If
    
    If Not mWinsock.SocketHandle = INVALID_SOCKET Then mWinsock.CloseSocket
    Connect = mWinsock.OpenSocket(afINET, stSockStream, IPPROTO_TCP)
    If Connect <> 0 Then
        mlngSocketState = sckError
        Exit Function
    End If
    
BIND_LABEL:
    Connect = mWinsock.Bind(mWinsock.GetLocalHostName, mlngLocalPort)
    If Connect <> 0 Then
        If mWinsock.GetLastError = 10048 Then
            mlngLocalPort = mlngLocalPort + 1
            DoEvents
            GoTo BIND_LABEL
        End If
        mlngSocketState = sckError
        Exit Function
    End If
    
    Connect = mWinsock.Connect(mstrSmtpServer, mlngSmtpPort)
    If Connect <> 0 Then
        mlngSocketState = sckError
        Exit Function
    End If
    
    mlngSocketState = sckOpen
    mstrSmtpServerIP = mWinsock.ResolveHostToIP(mstrSmtpServer, mlngSmtpPort)
End Function

Public Function Log(vLog As String)
    If mobjLogViewer Is Nothing Then Exit Function
    mobjLogViewer.Text = vLog & mobjLogViewer.Text
End Function

Public Property Let RecipientFrom(vData As String)
    Dim saryData, saryData2
    Dim udtRecipient As RECIPIENT
    
    Set mcolRecipientFrom = New Collection
    
    saryData = Split(vData, ";")
    saryData2 = Split(saryData(0), "<")
    
    If UBound(saryData2) = 0 Then
        ' Just an email address and no name
        saryData2(0) = Replace(saryData2(0), ">", "")
        With udtRecipient
            .Name = saryData2(0)
            .EMail = saryData2(0)
        End With
        mcolRecipientFrom.Add udtRecipient, saryData2(0)
    Else
        ' Name and email address supplied
        saryData2(1) = Replace(saryData2(1), ">", "")
        With udtRecipient
            .Name = Trim$(saryData2(0))
            .EMail = saryData2(1)
        End With
        mcolRecipientFrom.Add udtRecipient, saryData2(1)
    End If
End Property

Public Property Let RecipientTo(vData As String)
    Dim saryData, saryData2
    Dim i As Long
    Dim udtRecipient As RECIPIENT
    
    Set mcolRecipientTo = New Collection

    saryData = Split(vData, ";")
    For i = 0 To UBound(saryData)
        saryData2 = Split(saryData(i), "<")
        
        If UBound(saryData2) = 0 Then
            ' Just an email address and no name
            saryData2(0) = Replace(saryData2(0), ">", "")
            DeleteIfExists mcolRecipientTo, saryData2(0)
            With udtRecipient
                .Name = Trim$(saryData2(0))
                .EMail = saryData2(0)
            End With
            mcolRecipientTo.Add udtRecipient, saryData2(0)
        Else
            ' Name and email address supplied
            saryData2(1) = Replace(saryData2(1), ">", "")
            DeleteIfExists mcolRecipientTo, saryData2(1)
            With udtRecipient
                .Name = Trim$(saryData2(0))
                .EMail = saryData2(1)
            End With
            mcolRecipientTo.Add udtRecipient, saryData2(1)
        End If
    Next i
End Property

Public Property Let RecipientCC(vData As String)
    Dim saryData, saryData2
    Dim i As Long
    Dim udtRecipient As RECIPIENT
    
    Set mcolRecipientCC = New Collection

    saryData = Split(vData, ";")
    For i = 0 To UBound(saryData)
        saryData2 = Split(saryData(i), "<")
        
        If UBound(saryData2) = 0 Then
            ' Just an email address and no name
            saryData2(0) = Replace(saryData2(0), ">", "")
            DeleteIfExists mcolRecipientCC, saryData2(0)
            With udtRecipient
                .Name = Trim$(saryData2(0))
                .EMail = saryData2(0)
            End With
            mcolRecipientCC.Add udtRecipient, saryData2(0)
        Else
            ' Name and email address supplied
            saryData2(1) = Replace(saryData2(1), ">", "")
            DeleteIfExists mcolRecipientCC, saryData2(1)
            With udtRecipient
                .Name = Trim$(saryData2(0))
                .EMail = saryData2(1)
            End With
            mcolRecipientCC.Add udtRecipient, saryData2(1)
        End If
    Next i
End Property

Public Property Let RecipientBCC(vData As String)
    Dim saryData, saryData2
    Dim i As Long
    Dim udtRecipient As RECIPIENT
    
    Set mcolRecipientBCC = New Collection

    saryData = Split(vData, ";")
    For i = 0 To UBound(saryData)
        saryData2 = Split(saryData(i), "<")
        
        If UBound(saryData2) = 0 Then
            ' Just an email address and no name
            saryData2(0) = Replace(saryData2(0), ">", "")
            DeleteIfExists mcolRecipientBCC, saryData2(0)
            With udtRecipient
                .Name = Trim$(saryData2(0))
                .EMail = saryData2(0)
            End With
            mcolRecipientBCC.Add udtRecipient, saryData2(0)
        Else
            ' Name and email address supplied
            saryData2(1) = Replace(saryData2(1), ">", "")
            DeleteIfExists mcolRecipientBCC, saryData2(1)
            With udtRecipient
                .Name = Trim$(saryData2(0))
                .EMail = saryData2(1)
            End With
            mcolRecipientBCC.Add udtRecipient, saryData2(1)
        End If
    Next i
End Property

Public Property Let ReplyTo(vData As String)
    Dim saryData, saryData2
    Dim udtRecipient As RECIPIENT
    
    If vData = "" Then Set mcolReplyTo = mcolRecipientFrom: Exit Property
    
    Set mcolReplyTo = New Collection

    saryData = Split(vData, ";")
    saryData2 = Split(saryData(0), "<")
    
    If UBound(saryData2) = 0 Then
        ' Just an email address and no name
        saryData2(0) = Replace(saryData2(0), ">", "")
        With udtRecipient
            .Name = saryData2(0)
            .EMail = saryData2(0)
        End With
        mcolReplyTo.Add udtRecipient, saryData2(0)
    Else
        ' Name and email address supplied
        saryData2(1) = Replace(saryData2(1), ">", "")
        With udtRecipient
            .Name = Trim$(saryData2(0))
            .EMail = saryData2(1)
        End With
        mcolReplyTo.Add udtRecipient, saryData2(1)
    End If
End Property

Public Property Let Subject(vData As String)
    mstrSubject = vData
End Property

Public Property Let Body(vData As String)
    mstrBody = vData
End Property

Public Function SendMail(ByVal MailFormat As MailFormatConstants, _
                                    Optional ByVal charset As CharsetConstants = cUsAscii, _
                                    Optional ByVal Importance As ImportanceConstants = iNormal) As Long
    Dim lngRetVal As Long
    Dim lpBuffer As String
    Dim strMailFormat As String, strCharset As String
    Dim strImportance As String
    Dim i As Long
    
    If State = sckError Then Exit Function
    
    Select Case Importance
        Case iLow: strImportance = "Low"
        Case iNormal: strImportance = "Normal"
        Case iHigh: strImportance = "High"
    End Select
    
    Select Case MailFormat
        Case mfTextPlain
            strMailFormat = "text/plain"
        Case mfTextHtml
            strMailFormat = "text/html"
        Case Else
            strMailFormat = "text/plain"
    End Select
    
    Select Case charset
        Case cUsAscii
            strCharset = "us-ascii"
        Case Else
            strCharset = "us-ascii"
    End Select
    
    With mWinsock
        lpBuffer = "HELO " & Mid$(mstrSmtpServer, InStr(mstrSmtpServer, ".") + 1, Len(mstrSmtpServer)) & vbCrLf
        Log lpBuffer
        lngRetVal = .Send(lpBuffer)
        If lngRetVal = SOCKET_ERROR Then GoTo ERR_HANDLER
        SendMail = SendListen(5)
        If SendMail <> 0 Then Exit Function
        
        lpBuffer = "MAIL FROM: " & mcolRecipientFrom(1).Name & "<" & mcolRecipientFrom(1).EMail & ">" & vbCrLf
        lngRetVal = .Send(lpBuffer)
        If lngRetVal = SOCKET_ERROR Then GoTo ERR_HANDLER
        SendMail = SendListen(5)
        If SendMail <> 0 Then Exit Function
        
        For i = 1 To mcolRecipientTo.Count
            lpBuffer = mcolRecipientTo(i).Name & "<" & mcolRecipientTo(i).EMail & ">"
            lpBuffer = "RCPT TO: " & lpBuffer & vbCrLf
            lngRetVal = .Send(lpBuffer)
            If lngRetVal = SOCKET_ERROR Then GoTo ERR_HANDLER
            SendMail = SendListen(5)
            If SendMail <> 0 Then Exit Function
        Next i
        
        For i = 1 To mcolRecipientCC.Count
            lpBuffer = mcolRecipientCC(i).Name & "<" & mcolRecipientCC(i).EMail & ">"
            lpBuffer = "RCPT TO: " & lpBuffer & vbCrLf
            lngRetVal = .Send(lpBuffer)
            If lngRetVal = SOCKET_ERROR Then GoTo ERR_HANDLER
            SendMail = SendListen(5)
            If SendMail <> 0 Then Exit Function
        Next i
        
        For i = 1 To mcolRecipientBCC.Count
            lpBuffer = mcolRecipientBCC(i).Name & "<" & mcolRecipientBCC(i).EMail & ">"
            lpBuffer = "RCPT TO: " & lpBuffer & vbCrLf
            lngRetVal = .Send(lpBuffer)
            If lngRetVal = SOCKET_ERROR Then GoTo ERR_HANDLER
            SendMail = SendListen(5)
            If SendMail <> 0 Then Exit Function
        Next i
        
        lpBuffer = "DATA" & vbCrLf
        lngRetVal = .Send(lpBuffer)
        If lngRetVal = SOCKET_ERROR Then GoTo ERR_HANDLER
        SendMail = SendListen(5)
        If SendMail <> 0 Then Exit Function
        
        Log "Beginning transfer of body..." & vbCrLf
        lpBuffer = GenerateMessageID(Mid(mcolRecipientFrom(1).EMail, InStr(mcolRecipientFrom(1).EMail, "@") + 1, Len(mcolRecipientFrom(1).EMail))) & vbCrLf
        Log lpBuffer
        lngRetVal = .Send(lpBuffer)
        If lngRetVal = SOCKET_ERROR Then GoTo ERR_HANDLER
        
        lpBuffer = "DATE: " & WeekdayName(Weekday(Now), True, vbSunday) & " " & FormatDateTime(Now, vbShortDate) & " " & FormatDateTime(Now, vbShortTime) & vbCrLf
        Log lpBuffer
        lngRetVal = .Send(lpBuffer)
        If lngRetVal = SOCKET_ERROR Then GoTo ERR_HANDLER
                
        lpBuffer = "FROM: " & mcolRecipientFrom(1).Name & "<" & mcolRecipientFrom(1).EMail & ">" & vbCrLf
        Log lpBuffer
        lngRetVal = .Send(lpBuffer)
        If lngRetVal = SOCKET_ERROR Then GoTo ERR_HANDLER
        
        lpBuffer = ""
        For i = 1 To mcolRecipientTo.Count
            lpBuffer = lpBuffer & mcolRecipientTo(i).Name & "<" & mcolRecipientTo(i).EMail & ">;"
        Next i
        If Not lpBuffer = "" Then
            lpBuffer = "TO: " & lpBuffer & vbCrLf
            Log lpBuffer
            lngRetVal = .Send(lpBuffer)
            If lngRetVal = SOCKET_ERROR Then GoTo ERR_HANDLER
        End If
        
        lpBuffer = ""
        For i = 1 To mcolRecipientCC.Count
            lpBuffer = lpBuffer & mcolRecipientCC(i).Name & "<" & mcolRecipientCC(i).EMail & ">;"
        Next i
        If Not lpBuffer = "" Then
            lpBuffer = "CC: " & lpBuffer & vbCrLf
            Log lpBuffer
            lngRetVal = .Send(lpBuffer)
            If lngRetVal = SOCKET_ERROR Then GoTo ERR_HANDLER
        End If
        
        lpBuffer = ""
        For i = 1 To mcolRecipientBCC.Count
            lpBuffer = lpBuffer & mcolRecipientBCC(i).Name & "<" & mcolRecipientBCC(i).EMail & ">;"
        Next i
        If Not lpBuffer = "" Then
            lpBuffer = "BCC: " & lpBuffer & vbCrLf
            Log lpBuffer
            lngRetVal = .Send(lpBuffer)
            If lngRetVal = SOCKET_ERROR Then GoTo ERR_HANDLER
        End If
        
        lpBuffer = "Reply-To: " & mcolReplyTo(1).Name & "<" & mcolReplyTo(1).EMail & ">" & vbCrLf
        Log lpBuffer
        lngRetVal = .Send(lpBuffer)
        If lngRetVal = SOCKET_ERROR Then GoTo ERR_HANDLER
        
        lpBuffer = "IMPORTANCE: " & strImportance & vbCrLf
        Log lpBuffer
        lngRetVal = .Send(lpBuffer)
        If lngRetVal = SOCKET_ERROR Then GoTo ERR_HANDLER
        
        lpBuffer = "SUBJECT: " & mstrSubject & vbCrLf
        Log lpBuffer
        lngRetVal = .Send(lpBuffer)
        If lngRetVal = SOCKET_ERROR Then GoTo ERR_HANDLER
        
        lpBuffer = "MIME-Version: 1.0" & vbCrLf
        Log lpBuffer
        lngRetVal = .Send(lpBuffer)
        If lngRetVal = SOCKET_ERROR Then GoTo ERR_HANDLER
        
        lpBuffer = "Content-Type: " & strMailFormat & "; charset=" & strCharset & vbCrLf
        Log lpBuffer
        lngRetVal = .Send(lpBuffer)
        If lngRetVal = SOCKET_ERROR Then GoTo ERR_HANDLER
        
        lpBuffer = vbCrLf & vbCrLf & mstrBody & vbCrLf
        Log lpBuffer
        lngRetVal = .Send(lpBuffer)
        If lngRetVal = SOCKET_ERROR Then GoTo ERR_HANDLER
        
        lpBuffer = vbCrLf & "." & vbCrLf
        Log lpBuffer
        lngRetVal = .Send(lpBuffer)
        If lngRetVal = SOCKET_ERROR Then GoTo ERR_HANDLER
        
        lpBuffer = "QUIT" & vbCrLf
        Log lpBuffer
        lngRetVal = .Send(lpBuffer)
        If lngRetVal = SOCKET_ERROR Then GoTo ERR_HANDLER
        SendMail = SendListen(30)
        If SendMail <> 0 Then Exit Function
    End With
    
EXIT_HANDLER:
    Log "Mail sent successfully." & vbCrLf
    Exit Function
    
ERR_HANDLER:
    SendMail = SOCKET_ERROR
End Function

Public Function Listen() As String
    Dim lpBuffer As String

    mWinsock.Receive lpBuffer
End Function

Private Sub DeleteIfExists(mColToDelete As Collection, ByVal Index)
On Error Resume Next
    mColToDelete.Remove Index
End Sub

Public Property Get State() As WinsockStateConstants
    State = mlngSocketState
End Property

Private Function GenerateMessageID(ByVal Host As String) As String
    Dim idNum As Double
    Dim sMessageID As String
    sMessageID = "Message-ID: "
    Randomize Int(CDbl((Now))) + Timer
    idNum = GetRandom(9999999999999#, 99999999999999#)
    sMessageID = sMessageID & CStr(idNum)
    idNum = GetRandom(9999, 99999)
    sMessageID = sMessageID & "." & CStr(idNum) & ".qmail@" & Host
    GenerateMessageID = sMessageID
End Function

Private Function GetRandom(ByVal dFrom As Double, ByVal dTo As Double) As Double
    Dim x As Double
    Randomize
    x = dTo - dFrom
    GetRandom = Int((x * Rnd) + 1) + dFrom
End Function

Private Function SendListen(Optional ByVal Interval As Long = 10) As Long
    Dim lngTimer As Long
    Dim lpBuffer As String
    
    SendListen = 0
    
    lngTimer = Timer + Interval
    Do Until Timer > lngTimer
        mWinsock.Receive lpBuffer
        If Not lpBuffer = "" Then
            Log lpBuffer
            lpBuffer = ""
            Select Case Left$(lpBuffer, 3)
                Case 220
                Case 354, 250
                Case 221
                Case 550, 551, 552, 553, 554, 451, 452, 500
                    SendListen = Left$(lpBuffer, 3)
                    Exit Do
            End Select
        End If
        DoEvents
    Loop
End Function
