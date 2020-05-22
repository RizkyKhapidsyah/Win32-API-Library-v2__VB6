VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oldWinsock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mMath As Math
Private mStrings As Strings
Private mMem As Memory

Private mudtWSAData As WSAData

Private mlngLastError As Long

Private mstrRemoteHost As String
Private mlngRemotePort As Long
Private mstrRemoteHostIP As String

Private mstrLocalHost As String
Private mlngLocalPort As Long
Private mstrLocalHostIP As String

Private mlngSocketHandle As Long
Private mlngAddrFamily As Long

Private mvarInternalState As WinsockInternalStateConstants
Private mvarProtocol As WinsockProtocolConstants
Private mvarState As WinsockStateConstants

Private saZero As sockaddr_in

Public Property Get WSAVersion() As String
    WSAVersion = mudtWSAData.wVersion \ 256 & "." & mudtWSAData.wVersion Mod 256
End Property

Public Property Get WSAVersionHigh() As String
    WSAVersionHigh = mudtWSAData.wHighVersion \ 256 & "." & mudtWSAData.wHighVersion Mod 256
End Property

Public Property Get WSADescription() As String
    WSADescription = mStrings.TrimNull(mudtWSAData.szDescription)
End Property

Public Property Get WSASystemStatus() As String
    WSASystemStatus = mStrings.TrimNull(mudtWSAData.szSystemStatus)
End Property

Public Property Get WSAMaxSockets() As Long
    WSAMaxSockets = mMath.IntegerToUnsigned(mudtWSAData.iMaxSockets)
End Property

Public Property Get WSAMaxUdpDg() As Long
    WSAMaxUdpDg = mMath.IntegerToUnsigned(mudtWSAData.iMaxUdpDg)
End Property

Public Property Get WSAVendorInfo() As Long
    WSAVendorInfo = mudtWSAData.lpVendorInfo
End Property

Public Function FormatMessage(ByVal nErrorNumber As Long, lpMsgBuffer As String, ByVal MsgBuffLen As Long) As Long
    lpMsgBuffer = Space$(MsgBuffLen)
    
    Select Case nErrorNumber
        Case WSASYSNOTREADY
            lpMsgBuffer = "The underlying network subsystem is not ready for network communication." & vbNullChar & lpMsgBuffer
        Case WSAVERNOTSUPPORTED
            lpMsgBuffer = "The version of Windows Sockets API support requested is not provided by this particular Windows Sockets implementation." & vbNullChar & lpMsgBuffer
        Case WSAEINVAL
            lpMsgBuffer = "The Windows Sockets version specified by the application is not supported by this DLL." & vbNullChar & lpMsgBuffer
    End Select
    
    lpMsgBuffer = Left$(lpMsgBuffer, MsgBuffLen)
End Function

Public Function GetLastError() As Long
    GetLastError = mlngLastError
End Function

Public Function Startup(ByVal lVersion As Long) As Long
    Dim lngRetVal As Long
    
    lngRetVal = WSAStartup(lVersion, mudtWSAData)
    If lngRetVal <> 0 Then
        mlngLastError = lngRetVal
    End If
End Function

Private Sub Class_Initialize()
    Set mMath = New Math
    Set mStrings = New Strings
    Set mMem = New Memory
    
    Dim mComputer As New Computer
    
    mlngRemotePort = 25
    mstrRemoteHost = mComputer.Name
    mstrRemoteHostIP = "127.0.0.1"
    
    mstrLocalHost = mComputer.Name
    mstrLocalHostIP = "127.0.0.1"
    mlngLocalPort = 25
    
    mlngSocketHandle = INVALID_SOCKET
    
    Set mComputer = Nothing
End Sub

Private Sub Class_Terminate()
    Set mMath = Nothing
    Set mStrings = Nothing
    Set mMem = Nothing
    
    WSACleanup
End Sub

Public Property Let RemoteHost(vData As String)
    mstrRemoteHost = vData
    
    mstrRemoteHostIP = ResolveHostIP(mstrRemoteHost, mlngRemotePort)
End Property

Public Property Get RemoteHost() As String
    RemoteHost = mstrRemoteHost
End Property

Public Property Get RemoteHostIP() As String
    RemoteHostIP = mstrRemoteHostIP
End Property

Public Property Let RemotePort(vData As Long)
    mlngRemotePort = vData
End Property

Public Property Get RemotePort() As Long
    RemotePort = mlngRemotePort
End Property

Public Function ResolveHostIP(ByVal vHost As String, ByVal vPort As Long) As String
    Dim sockin As sockaddr_in
    Dim retIpPort As String

    sockin = saZero
    sockin.sin_family = afINET
    sockin.sin_port = htons(vPort)
    If sockin.sin_port = INVALID_SOCKET Then
        ResolveHostIP = "127.0.0.1"
        Exit Function
    End If

    sockin.sin_addr = GetHostByNameAlias(vHost)
    If sockin.sin_addr = INADDR_NONE Then
        ResolveHostIP = "127.0.0.1"
        Exit Function
    End If
    ResolveHostIP = GetAscIp$(sockin.sin_addr)
End Function

Private Function GetHostByNameAlias(ByVal Host As String) As Long
On Error Resume Next
    Dim phe As Long
    Dim lpDestHost As HOSTENT
    Dim addrList As Long
    Dim retIP As Long
    
    retIP = inet_addr(Host)
    If retIP = INADDR_NONE Then
        phe = gethostbyname(Host)
        If phe <> 0 Then
            RtlMoveMemory lpDestHost, ByVal phe, LenB(lpDestHost)
            RtlMoveMemory addrList, ByVal lpDestHost.hAddrList, 4
            RtlMoveMemory retIP, ByVal addrList, lpDestHost.hLength
        Else
            retIP = INADDR_NONE
        End If
    End If
    GetHostByNameAlias = retIP
    If err Then GetHostByNameAlias = INADDR_NONE
End Function

Private Function GetAscIp(ByVal inn As Long) As String
On Error Resume Next
    Dim lpStr As Long
    Dim nStr As Long
    Dim retStr As String
    
    retStr = String(32, 0)
    lpStr = inet_ntoa(inn)
    If lpStr = 0 Then
        GetAscIp = "255.255.255.255"
        Exit Function
    End If
    
    nStr = lstrlen(lpStr)
    If nStr > 32 Then nStr = 32
    RtlMoveMemory ByVal retStr, ByVal lpStr, nStr
    retStr = Left(retStr, nStr)
    GetAscIp = retStr
    If err Then GetAscIp = "255.255.255.255"
End Function

Public Property Let LocalHost(vData As String)
    mstrLocalHost = vData
    
    mstrLocalHostIP = ResolveHostIP(mstrLocalHost, mlngLocalPort)
End Property

Public Property Get LocalHost() As String
    LocalHost = mstrLocalHost
End Property

Public Property Get LocalHostIP() As String
    LocalHostIP = mstrLocalHostIP
End Property

Public Property Let LocalPort(vData As Long)
    mlngLocalPort = vData
End Property

Public Property Get LocalPort() As Long
    LocalPort = mlngLocalPort
End Property

Public Function CreateSocket(ByVal AddressFamily As AddressFamilyConstants, _
                                        ByVal SocketType As SocketTypeConstants, _
                                        ByVal Protocol As WinsockProtocolConstants) As Long
On Error GoTo ERR_HANDLER
    Dim lngSocket As Long
    
    lngSocket = socket(AddressFamily, SocketType, Protocol)
    
    If lngSocket = INVALID_SOCKET Then
        CreateSocket = err.LastDllError
    Else
        mlngSocketHandle = lngSocket
        mlngAddrFamily = AddressFamily
    End If

EXIT_HANDLER:
    Exit Function
    
ERR_HANDLER:
    CreateSocket = INVALID_SOCKET
End Function

Public Function CloseSocket(lngSocketHandle As Long) As Long
    api_closesocket lngSocketHandle
    mlngSocketHandle = INVALID_SOCKET
End Function

Public Property Get SocketHandle() As Long
    SocketHandle = mlngSocketHandle
End Property

Public Function Connect(ByVal strRemoteHost As String, ByVal lngRemotePort As Long) As Long
On Error GoTo ERR_HANDLER
    Dim udtSocketAddress As sockaddr_in
    Dim lngRetVal As Long
    Dim lngAddress As Long
    
    Connect = SOCKET_ERROR
    
    If Not mlngSocketHandle > 0 Then
        Connect = sckNotSocket
        Exit Function
    End If
    
    If IsMissing(strRemoteHost) Then
        If Len(mstrRemoteHost) = 0 Then
            Connect = sckHostNotFound
        End If
    End If
    If IsMissing(lngRemotePort) Then
        If mlngRemotePort = 0 Then
            Connect = sckPortNotSupported
        End If
    End If
    
    lngAddress = GetHostByNameAlias(strRemoteHost)
    If lngAddress = INADDR_NONE Then
        Connect = sckAddressNotAvailable
        Exit Function
    End If
    
    With udtSocketAddress
        .sin_addr = lngAddress
        .sin_port = htons(mMath.UnsignedToInteger(CLng(lngRemotePort)))
        .sin_family = mlngAddrFamily
    End With
    
    Connect = api_connect(mlngSocketHandle, udtSocketAddress, LenB(udtSocketAddress))

EXIT_HANDLER:
    Exit Function
    
ERR_HANDLER:
    Connect = SOCKET_ERROR
End Function

Public Function SocketExists() As Boolean
    If mlngSocketHandle = INVALID_SOCKET Then
        Call CreateSocket(afINET, stSockStream, sckTCPProtocol)
        If mlngSocketHandle = SOCKET_ERROR Then
            Exit Function
        Else
        End If
    End If
    SocketExists = True
End Function

Public Function Bind(ByVal strLocalHost As String, ByVal lngLocalPort As Long) As Long
On Error GoTo ERR_HANDLER
    Dim udtSocketAddress As sockaddr_in
    Dim lngRetVal As Long
    Dim lngAddress As Long
    
    Bind = SOCKET_ERROR
    
    If Not mlngSocketHandle > 0 Then
        Bind = sckNotSocket
        Exit Function
    End If
    
    If IsMissing(strLocalHost) Then
        If Len(mstrLocalHost) = 0 Then
            Bind = sckHostNotFound
        End If
    End If
    If IsMissing(lngLocalPort) Then
        If mlngLocalPort = 0 Then
            Bind = sckPortNotSupported
        End If
    End If
    
    lngAddress = GetHostByNameAlias(strLocalHost)
    If lngAddress = INADDR_NONE Then
        Bind = sckAddressNotAvailable
        Exit Function
    End If
    
    With udtSocketAddress
        .sin_addr = lngAddress
        .sin_port = htons(mMath.UnsignedToInteger(CLng(lngLocalPort)))
        .sin_family = mlngAddrFamily
    End With
    
    Bind = api_bind(mlngSocketHandle, udtSocketAddress, LenB(udtSocketAddress))

EXIT_HANDLER:
    Exit Function
    
ERR_HANDLER:
    Bind = SOCKET_ERROR
End Function

Public Function Send(ByVal vData As String) As Long
    Dim arrBuffer() As Byte
    Dim lngBytesSent As Long
    Dim lngBufferLength As Long
    
    lngBufferLength = Len(vData)
    
    If IsConnected And lngBufferLength > 0 Then
        arrBuffer() = StrConv(vData, vbFromUnicode)
        lngBytesSent = api_send(mlngSocketHandle, arrBuffer(0), lngBufferLength, 0&)
        Send = lngBytesSent
    Else
        Send = SOCKET_ERROR
    End If
End Function

Public Function Receive(strBuffer As String) As Long
    Const MAX_BUFFER_LENGTH = 8192
    
    Dim arrBuffer(1 To MAX_BUFFER_LENGTH) As Byte
    Dim lngBytesReceived As Long
    Dim strTempBuffer As String
    
    If IsDataAvailable Then
        lngBytesReceived = recv(mlngSocketHandle, arrBuffer(1), MAX_BUFFER_LENGTH, 0&)
        If lngBytesReceived > 0 Then
            strTempBuffer = StrConv(arrBuffer, vbUnicode)
            strBuffer = Left$(strTempBuffer, lngBytesReceived)
        End If
        Receive = lngBytesReceived
    Else
        Receive = SOCKET_ERROR
    End If
End Function

Public Function IsDataAvailable() As Boolean
    Dim udtRead_fd As fd_set
    Dim udtWrite_fd As fd_set
    Dim udtError_fd As fd_set
    Dim lngSoccketCount As Long
    
    udtRead_fd.fd_count = 1
    udtRead_fd.fd_array(1) = mlngSocketHandle
    lngSoccketCount = api_select(0&, udtRead_fd, udtWrite_fd, udtError_fd, 0&)
    IsDataAvailable = CBool(lngSoccketCount)
End Function

Public Function GenerateMessageID(ByVal Host As String) As String
    Dim idNum As Double
    Dim sMessageID As String
    sMessageID = "Message-ID: "
    Randomize Int(CDbl((Now))) + Timer
    idNum = GetRandom(9999999999999#, 9999999999999#)
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

Private Function IsConnected() As Boolean
    Dim udtRead_fd As fd_set
    Dim udtWrite_fd As fd_set
    Dim udtError_fd As fd_set
    Dim lngSocketCount As Long
    
    udtWrite_fd.fd_count = 1
    udtWrite_fd.fd_array(1) = mlngSocketHandle
    
    lngSocketCount = api_select(0&, udtRead_fd, udtWrite_fd, udtError_fd, 0&)
    
    IsConnected = CBool(lngSocketCount)
End Function
