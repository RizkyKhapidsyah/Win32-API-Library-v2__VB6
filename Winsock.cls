VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Winsock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mMath As Math
Private mMem As Memory
Private mStrings As Strings
Private mComputer As Computer

' Startup parameters
Private mudtWSAData As WSAData

' Socket parameters
Private mhSocketHandle As Long
Private mlngAddressFamily As Long
Private mlngSocketType As Long
Private mlngProtocol As Long

' Error parameter
Private mlngLastError As Long

Private mSocketState As WinsockStateConstants

Private Sub Class_Initialize()
    Set mMath = New Math
    Set mMem = New Memory
    Set mStrings = New Strings
    Set mComputer = New Computer
    
    mhSocketHandle = INVALID_SOCKET
    mlngAddressFamily = afINET
    mlngSocketType = stSockStream
    mlngProtocol = IPPROTO_TCP
    mSocketState = sckClosed
End Sub

Public Function ResolveHostToIP(ByVal HostName As String, ByVal PortNumber As Long) As String
    Dim sockin As sockaddr_in
    Dim retIpPort
    
    sockin.sin_family = mlngAddressFamily
    sockin.sin_port = htons(PortNumber)
    If sockin.sin_port = INVALID_SOCKET Then
        ResolveHostToIP = "127.0.0.1"
        Exit Function
    End If
    
    sockin.sin_addr = GetHostByNameAlias(HostName)
    If sockin.sin_addr = INADDR_NONE Then
        ResolveHostToIP = "127.0.0.1"
        Exit Function
    End If
    ResolveHostToIP = GetAscIp(sockin.sin_addr)
End Function

Private Function GetHostByNameAlias(ByVal HostName As String) As Long
On Error Resume Next
    Dim phe As Long
    Dim lpDestHost As HOSTENT
    Dim addrList As Long
    Dim retIP As Long
    
    retIP = inet_addr(HostName)
    If retIP = INADDR_NONE Then
        phe = gethostbyname(HostName)
        If Not phe = 0 Then
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
    Dim retString As String
    
    retString = String(32, 0)
    lpStr = inet_ntoa(inn)
    If lpStr = 0 Then
        GetAscIp = "255.255.255.255"
        Exit Function
    End If
    
    nStr = lstrlen(lpStr)
    If nStr > 32 Then nStr = 32
    RtlMoveMemory ByVal retString, ByVal lpStr, nStr
    retString = Left(retString, nStr)
    GetAscIp = retString
    If err Then GetAscIp = "255.255.255.255"
End Function

Public Function Startup(ByVal wVersion As Long) As Long
    Dim lngRetVal As Long
    
    lngRetVal = WSAStartup(wVersion, mudtWSAData)
    If Not lngRetVal = 0 Then
        mlngLastError = err.LastDllError
        Startup = SOCKET_ERROR
    End If
End Function

Public Function Shutdown() As Long
    Dim lngRetVal As Long
    lngRetVal = api_shutdown(mhSocketHandle, SD_SEND)
    If lngRetVal = SOCKET_ERROR Then
        mlngLastError = err.LastDllError
    End If
    
    Dim lpBuffer As String
    FormatErrorMessage GetLastError, lpBuffer
End Function

Private Sub Class_Terminate()
    Set mMath = Nothing
    Set mMem = Nothing
    Set mStrings = Nothing
    Set mComputer = Nothing
    
    WSACleanup
End Sub

Public Function OpenSocket(ByVal AddressFamily As AddressFamilyConstants, _
                                        ByVal SocketType As SocketTypeConstants, _
                                        ByVal Protocol As ProtocolConstants) As Long
On Error GoTo ERR_HANDLER
    Dim hSocket As Long
    
    mlngAddressFamily = AddressFamily
    mlngSocketType = SocketType
    mlngProtocol = Protocol
    
    hSocket = socket(mlngAddressFamily, mlngSocketType, mlngProtocol)
    If hSocket = INVALID_SOCKET Then
        mlngLastError = err.LastDllError
        OpenSocket = SOCKET_ERROR
        mSocketState = sckError
    Else
        mhSocketHandle = hSocket
        mSocketState = sckOpen
    End If
    
EXIT_HANDLER:
    Exit Function
    
ERR_HANDLER:
    mlngLastError = err.LastDllError
    OpenSocket = SOCKET_ERROR
    mSocketState = sckError
End Function

Public Function CloseSocket() As Long
    api_closesocket mhSocketHandle
    
    mhSocketHandle = INVALID_SOCKET
    mlngAddressFamily = afINET
    mlngSocketType = stSockStream
    mlngProtocol = IPPROTO_TCP
    mSocketState = sckClosed
End Function

Public Property Get SocketHandle() As Long
    SocketHandle = mhSocketHandle
End Property

Public Function Bind(ByVal strLocalHost As String, ByVal lngLocalPort As Long) As Long
On Error GoTo ERR_HANDLER
    Dim udtSocketAddress As sockaddr_in
    Dim lngRetVal As Long
    Dim lngAddress As Long
    
    Bind = SOCKET_ERROR
    
    If Not mhSocketHandle > 0 Then
        mlngLastError = WSANOTINITIALIZED
        Bind = SOCKET_ERROR
        Exit Function
    End If
    
    If IsMissing(strLocalHost) Then
        mlngLastError = WSAHOSTNOTFOUND
        Bind = SOCKET_ERROR
        Exit Function
    End If
    If IsMissing(lngLocalPort) Then
        mlngLastError = WSAEINVAL
        Bind = SOCKET_ERROR
        Exit Function
    End If
    
    lngAddress = GetHostByNameAlias(strLocalHost)
    If lngAddress = INADDR_NONE Then
        mlngLastError = WSAHOSTNOTFOUND
        Bind = SOCKET_ERROR
        Exit Function
    End If
    
    With udtSocketAddress
        .sin_addr = lngAddress
        .sin_port = htons(mMath.UnsignedToInteger(CLng(lngLocalPort)))
        .sin_family = mlngAddressFamily
    End With
    
    Bind = api_bind(mhSocketHandle, udtSocketAddress, LenB(udtSocketAddress))
    If Bind = SOCKET_ERROR Then mlngLastError = err.LastDllError
    
EXIT_HANDLER:
    Exit Function
    
ERR_HANDLER:
    mlngLastError = err.LastDllError
    Bind = SOCKET_ERROR
End Function

Public Function Connect(ByVal strRemoteHost As String, ByVal lngRemotePort As Long) As Long
On Error GoTo ERR_HANDLER
    Dim udtSocketAddress As sockaddr_in
    Dim lngRetVal As Long
    Dim lngAddress As Long
    
    Connect = SOCKET_ERROR
    
    If Not mhSocketHandle > 0 Then
        mlngLastError = WSANOTINITIALIZED
        Connect = SOCKET_ERROR
        Exit Function
    End If
    
    If Len(strRemoteHost) = 0 Then
        mlngLastError = WSAEDESTADDRREQ
        Connect = SOCKET_ERROR
        Exit Function
    End If
    If lngRemotePort = 0 Then
        lngRemotePort = 25
    End If
    
    lngAddress = GetHostByNameAlias(strRemoteHost)
    If lngAddress = INADDR_NONE Then
        mlngLastError = WSAHOSTNOTFOUND
        Connect = SOCKET_ERROR
        Exit Function
    End If
    
    With udtSocketAddress
        .sin_addr = lngAddress
        .sin_port = htons(mMath.UnsignedToInteger(CLng(lngRemotePort)))
        .sin_family = mlngAddressFamily
    End With
    
    Connect = api_connect(mhSocketHandle, udtSocketAddress, LenB(udtSocketAddress))
    If Connect = SOCKET_ERROR Then mlngLastError = err.LastDllError
    
EXIT_HANDLER:
    Exit Function
    
ERR_HANDLER:
    mlngLastError = err.LastDllError
    Connect = SOCKET_ERROR
End Function

Public Function SocketExists() As Boolean
    If mhSocketHandle = INVALID_SOCKET Then
        OpenSocket mlngAddressFamily, mlngSocketType, mlngProtocol
        If mhSocketHandle = SOCKET_ERROR Then
            mlngLastError = err.LastDllError
            Exit Function
        End If
    End If
    SocketExists = True
End Function

Public Function IsConnected() As Boolean
    Dim udtRead_fd As fd_set
    Dim udtWrite_fd As fd_set
    Dim udtError_fd As fd_set
    Dim lngSocketCount As Long
    
    udtWrite_fd.fd_count = 1
    udtWrite_fd.fd_array(1) = mhSocketHandle
    
    lngSocketCount = api_select(0&, udtRead_fd, udtWrite_fd, udtError_fd, 0&)
    
    IsConnected = CBool(lngSocketCount)
End Function

Public Function IsDataAvailable() As Boolean
    Dim udtRead_fd As fd_set
    Dim udtWrite_fd As fd_set
    Dim udtError_fd As fd_set
    Dim lngSocketCount As Long
    
    udtRead_fd.fd_count = 1
    udtRead_fd.fd_array(1) = mhSocketHandle
    
    lngSocketCount = api_select(0&, udtRead_fd, udtWrite_fd, udtError_fd, 0&)
    
    IsDataAvailable = CBool(lngSocketCount)
End Function

Public Function Send(ByVal vData As String) As Long
    Dim arrBuffer() As Byte
    Dim lngBytesSent As Long
    Dim lngBufferLength As Long
    
    lngBufferLength = Len(vData)
    
    If IsConnected And lngBufferLength > 0 Then
        arrBuffer() = StrConv(vData, vbFromUnicode)
        lngBytesSent = api_send(mhSocketHandle, arrBuffer(0), lngBufferLength, 0&)
        Send = lngBytesSent
    Else
        mlngLastError = WSANODATA
        Send = SOCKET_ERROR
    End If
End Function

Public Function Receive(strBuffer As String) As Long
    Const MAX_BUFFER_LENGTH = 8192
    
    Dim arrBuffer(1 To MAX_BUFFER_LENGTH) As Byte
    Dim lngBytesReceived As Long
    Dim strTempBuffer As String
        
    If IsDataAvailable Then
        lngBytesReceived = recv(mhSocketHandle, arrBuffer(1), MAX_BUFFER_LENGTH, 0&)
        If lngBytesReceived > 0 Then
            strTempBuffer = StrConv(arrBuffer, vbUnicode)
            strBuffer = Left$(strTempBuffer, lngBytesReceived)
        End If
        Receive = lngBytesReceived
    Else
        strBuffer = ""
        Receive = SOCKET_ERROR
    End If
End Function

Public Function GetLastError() As Long
    GetLastError = mlngLastError
End Function

Public Function FormatErrorMessage(ByVal ErrorNumber As Long, lpBuffer As String) As Long
    lpBuffer = ""
    
    Select Case ErrorNumber
        Case WSAEACCESS: lpBuffer = "Permission denied."
        Case WSAEADDRINUSE: lpBuffer = "Address already in use."
        Case WSAEADDRNOTAVAIL: lpBuffer = "Cannot assign requested address."
        Case WSAEAFNOSUPPORT: lpBuffer = "Address family not supported by protocol family."
        Case WSAEALREADY: lpBuffer = "Operation already in progress."
        Case WSAEBADF: lpBuffer = "WSAEBADF"
        Case WSAECANCELLED: lpBuffer = "WSAECANCELLED"
        Case WSAECONNABORTED: lpBuffer = "Connection aborted."
        Case WSAECONNREFUSED: lpBuffer = "Connection refused."
        Case WSAECONNRESET: lpBuffer = "Connection reset by peer."
        Case WSAEDESTADDRREQ: lpBuffer = "Destination address required."
        Case WSAEDISCON: lpBuffer = "Graceful shutdown in progress."
        Case WSAEDQUOT: lpBuffer = "WSAEDQUOT"
        Case WSAEFAULT: lpBuffer = "Bad address."
        Case WSAEHOSTDOWN: lpBuffer = "Host is down."
        Case WSAEHOSTUNREACH: lpBuffer = "Host is unreachable."
        Case WSAEINPROGRESS: lpBuffer = "Operation now in progress."
        Case WSAEINTR: lpBuffer = "Interrupted function call."
        Case WSAEINVAL: lpBuffer = "Invalid argument."
        Case WSAEINVALIDPROCTABLE: lpBuffer = "WSAEINVALIDPROCTABLE"
        Case WSAEINVALIDPROVIDER: lpBuffer = "WSAEINVALIDPROVIDER"
        Case WSAEISCONN: lpBuffer = "Socket is already connected."
        Case WSAELOOP: lpBuffer = "WSAELOOP"
        Case WSAEMFILE: lpBuffer = "Too many open files."
        Case WSAEMSGSIZE: lpBuffer = "Message is too long."
        Case WSAENAMETOOLONG: lpBuffer = "WSAENAMETOOLONG"
        Case WSAENETDOWN: lpBuffer = "Network is down."
        Case WSAENETRESET: lpBuffer = "Network dropped connection on reset."
        Case WSAENETUNREACH: lpBuffer = "Network is unreachable."
        Case WSAENOBUFS: lpBuffer = "No buffer space available."
        Case WSAENOMORE: lpBuffer = "WSAENOMORE"
        Case WSAENOPROTOOPT: lpBuffer = "Bad protocol option."
        Case WSAENOTCONN: lpBuffer = "Socket is not connected."
        Case WSAENOTEMPTY: lpBuffer = "WSAENOTEMPTY"
        Case WSAENOTSOCK: lpBuffer = "Socket operation on nonsocket."
        Case WSAEOPNOTSUPP: lpBuffer = "Operation not supported."
        Case WSAEPFNOSUPPORT: lpBuffer = "Protocol family not supprted."
        Case WSAEPROCLIM: lpBuffer = "Too many processes."
        Case WSAEPROTONOSUPPORT: lpBuffer = "Protocol not supported."
        Case WSAEPROTOTYPE: lpBuffer = "Protocol wrong type for socket."
        Case WSAEPROVIDERFAILEDINIT: lpBuffer = "WSAEPROVIDERFAILEDINIT"
        Case WSAEREFUSED: lpBuffer = "WSAEREFUSED"
        Case WSAEREMOTE: lpBuffer = "WSAEREMOTE"
        Case WSAESHUTDOWN: lpBuffer = "Cannot send after socket shutdown."
        Case WSAESOCKTNOSUPPORT: lpBuffer = "Socket type not supported."
        Case WSAESTALE: lpBuffer = "WSAESTALE"
        Case WSAETIMEDOUT: lpBuffer = "Connection timed out."
        Case WSAETOOMANYREFS: lpBuffer = "WSAETOOMANYREFS"
        Case WSAEUSERS: lpBuffer = "WSAEUSERS"
        Case WSAEWOULDBLOCK: lpBuffer = "Resource temporarily unavailable."
        Case WSAHOSTNOTFOUND: lpBuffer = "Host not found."
        Case WSANODATA: lpBuffer = "Valid name, no data record of requested type."
        Case WSANORECOVERY: lpBuffer = "Nonrecoverable error."
        Case WSANOTINITIALIZED: lpBuffer = "Successful WSAStartup not yet performed."
        Case WSASERVICENOTFOUND: lpBuffer = "WSASERVICENOTFOUND"
        Case WSASYSCALLFAILED: lpBuffer = "System call failure."
        Case WSASYSNOTREADY: lpBuffer = "Network subsystem is unavailable."
        Case WSATRYAGAIN: lpBuffer = "Nonauthoritative host not found."
        Case WSATYPENOTFOUND: lpBuffer = "Class type not found."
        Case WSAVERNOTSUPPORTED: lpBuffer = "Winsock.dll version out of range."
        Case Else: lpBuffer = "Unsupported error."
    End Select
    
    lpBuffer = lpBuffer & String$(255 - Len(lpBuffer), vbNullChar)
End Function

Public Function GetLocalHostName() As String
    Dim lngRetVal As Long
    Dim strHostName As String * 256
    
    lngRetVal = gethostname(strHostName, 256)
    If lngRetVal = SOCKET_ERROR Then
    Else
        GetLocalHostName = mStrings.TrimNull(strHostName)
    End If
End Function
