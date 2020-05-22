Attribute VB_Name = "mdlDeclare"
'*************************************************************
'
' Mmodule containing all the API declarations needed for
' this DLL and private userdefined types
'
' OWNER: Neil Stevens
' CREATED: 30/01/2003
' HISTORY:
'   1.0.0 - Created
'
'*************************************************************
Option Explicit

' General declarations
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" _
    (ByVal dwExStyle As Long, ByVal lpClassName As String, _
    ByVal lpWindowName As String, ByVal dwStyle As Long, _
    ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal hWndParent As Long, _
    ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" _
    (ByVal lpString As String) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc&, ByVal hwnd&, ByVal Msg&, ByVal wParam&, ByVal lParam&) As Long
Public Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, _
    lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function SetWindowPlacement Lib "user32" (ByVal hwnd As Long, _
    lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BROWSEINFO) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, _
    ByVal lpBuffer As String) As Long
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
    (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal dwFlags As Long) As Long
Public Declare Sub RtlZeroMemory Lib "kernel32" (Destination As Any, ByVal Length As Long)
Public Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, source As Any, ByVal Length As Long)
Public Declare Sub RtlCopyMemory Lib "kernel32" (Destination As Any, source As Any, ByVal Length As Long)
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" _
    (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" _
    (ByVal lpString As Any) As Long
Public Declare Function lstrcmp Lib "kernel32" Alias "lstrcmpA" _
    (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function IsCharLower Lib "user32" Alias "IsCharLowerA" _
    (ByVal cChar As Byte) As Long
Public Declare Function IsCharUpper Lib "user32" Alias "IsCharUpperA" _
    (ByVal cChar As Byte) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

' Event log declarations
Public Declare Function ReportEvent Lib "advapi32.dll" Alias "ReportEventA" (ByVal hEventLog As Long, _
    ByVal wType As Long, ByVal wCategory As Long, ByVal dwEventID As Long, lpUserSid As Any, _
    ByVal wNumStrings As Long, ByVal dwDataSize As Long, ByVal lpStrings As Long, lpRawData As Any) As Long
Public Declare Function RegisterEventSource Lib "advapi32.dll" Alias "RegisterEventSourceA" _
    (ByVal lpUNCServerName As String, ByVal lpSourceName As String) As Long
Public Declare Function DeregisterEventSource Lib "advapi32.dll" _
    (ByVal hEventLog As Long) As Long

' Internet connection declarations
Public Declare Function InternetGetConnectedState Lib "wininet.dll" _
    (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long

' Network declarations
Public Declare Function WNetAddConnection Lib "mpr.dll" Alias "WNetAddConnection2A" _
    (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUsername As String, _
    ByVal dwFlags As Long) As Long
Public Declare Function WNetCancelConnection Lib "mpr.dll" Alias "WNetCancelConnection2A" _
    (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long
Public Declare Function WNetGetUser Lib "mpr.dll" Alias "WNetGetUserA" _
    (ByVal lpName As String, ByVal lpUsername As String, lpnLength As Long) As Long
Public Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" _
    (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, _
    cbRemoteName As Long) As Long
Public Declare Function WNetConnectDialog Lib "mpr.dll" _
    (ByVal hwnd As Long, ByVal dwType As Long) As Long
Public Declare Function WNetDisconnectDialog Lib "mpr.dll" _
    (ByVal hwnd As Long, ByVal dwType As Long) As Long

' Registry declarations
Public Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" _
    Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
    lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
    ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, _
    ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, _
    lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, _
    lpdwDisposition As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
    (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
    (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" _
    (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, _
    lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, _
    lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
    (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
    lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, _
    lpData As Byte, lpcbData As Long) As Long
Public Declare Function RegConnectRegistry Lib "advapi32.dll" _
    Alias "RegConnectRegistryA" (ByVal lpMachineName As String, _
    ByVal hKey As Long, phkResult As Long) As Long
Public Declare Function ExpandEnvironmentStrings Lib "kernel32" _
    Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, _
    ByVal lpDst As String, ByVal nSize As Long) As Long

' System declarations
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" _
    (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" _
    (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" _
    (ByVal hWndOwner As Long, ByVal nFolder As CSIDL, ByVal hToken As Long, _
    ByVal dwFlags As CSIDL_FOLDERPATH, ByVal pszPath As String) As Long
Public Declare Function GetCurrentDirectory Lib "kernel32" Alias "GetCurrentDirectoryA" _
    (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long

' Windows declarations
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (lpVersionInformation As OSVERSIONINFOEX) As Long
Public Declare Function VerifyVersionInfo Lib "kernel32" Alias "VerifyVersionInfoA" _
    (lpVersionInfo As OSVERSIONINFOEX, ByVal dwTypeMask As Long, _
    ByVal dwlConditionMask As Long) As Long

Public Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type

' System tray
Public Declare Function ShellNotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" _
    (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Long
    
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 128
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uTimeoutOrVersion As Long
    szInfoTitle As String * 64
    dwInfoFlags As Long
End Type

' Winsock declarations
Public Declare Function gethostbyaddr Lib "ws2_32.dll" (addr As Long, ByVal addr_len As Long, _
    ByVal addr_type As Long) As Long
Public Declare Function gethostbyname Lib "ws2_32.dll" (ByVal host_name As String) As Long
Public Declare Function gethostname Lib "ws2_32.dll" (ByVal host_name As String, _
    ByVal nameLen As Long) As Long
Public Declare Function getservbyname Lib "ws2_32.dll" (ByVal serv_name As String, _
    ByVal proto As String) As Long
Public Declare Function getprotobynumber Lib "ws2_32.dll" (ByVal proto As Long) As Long
Public Declare Function getprotobyname Lib "ws2_32.dll" (ByVal proto_name As String) As Long
Public Declare Function getservbyport Lib "ws2_32.dll" (ByVal Port As Integer, ByVal proto As Long) As Long
Public Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long
Public Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal inn As Long) As Long
Public Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer
Public Declare Function htonl Lib "ws2_32.dll" (ByVal hostlong As Long) As Long
Public Declare Function ntohl Lib "ws2_32.dll" (ByVal netlong As Long) As Long
Public Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Integer) As Integer
Public Declare Function socket Lib "ws2_32.dll" (ByVal af As Long, _
    ByVal s_type As Long, ByVal Protocol As Long) As Long
Public Declare Function api_closesocket Lib "ws2_32.dll" Alias "closesocket" (ByVal s As Long) As Long
Public Declare Function api_connect Lib "ws2_32.dll" Alias "connect" (ByVal s As _
    Long, ByRef Name As sockaddr_in, ByVal nameLen As Long) As Long
Public Declare Function getsockname Lib "ws2_32.dll" (ByVal s As Long, _
    ByRef Name As sockaddr_in, ByRef nameLen As Long) As Long
Public Declare Function getpeername Lib "ws2_32.dll" (ByVal s As Long, _
    ByRef Name As sockaddr_in, ByRef nameLen As Long) As Long
Public Declare Function api_bind Lib "ws2_32.dll" Alias "bind" (ByVal s As Long, _
    ByRef Name As sockaddr_in, ByRef nameLen As Long) As Long
Public Declare Function api_select Lib "ws2_32.dll" Alias "select" (ByVal nfds As Long, _
    ByRef readfds As Any, ByRef writefds As Any, ByRef exceptfds As Any, ByRef TimeOut As Long) As Long
Public Declare Function recv Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, _
    ByVal buflen As Long, ByVal Flags As Long) As Long
Public Declare Function api_send Lib "ws2_32.dll" Alias "send" (ByVal s As Long, ByRef buf As _
    Any, ByVal buflen As Long, ByVal Flags As Long) As Long
Public Declare Function api_shutdown Lib "ws2_32.dll" Alias "shutdown" (ByVal s As Long, ByVal how As Long) As Long
Public Declare Function api_listen Lib "ws2_32.dll" Alias "listen" (ByVal s As Long, _
    ByVal backlog As Long) As Long
Public Declare Function api_accept Lib "ws2_32.dll" Alias "accept" (ByVal s As Long, _
    ByRef addr As sockaddr_in, ByRef addrlen As Long) As Long
Public Declare Function setsockopt Lib "ws2_32.dll" (ByVal s As Long, _
    ByVal level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Public Declare Function getsockopt Lib "ws2_32.dll" (ByVal s As Long, _
    ByVal level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
Public Declare Function SendTo Lib "ws2_32.dll" Alias "sendto" (ByVal s As _
    Long, ByRef buf As Any, ByVal buflen As Long, ByVal Flags As Long, ByRef toaddr As sockaddr_in, ByVal tolen As Long) As Long
Public Declare Function recvfrom Lib "ws2_32.dll" (ByVal s As Long, _
    ByRef buf As Any, ByVal buflen As Long, ByVal Flags As Long, ByRef from As sockaddr_in, ByRef fromlen As Long) As Long

Public Declare Function WSAAsyncSelect Lib "ws2_32.dll" (ByVal s As Long, _
    ByVal hwnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Public Declare Function WSAAsyncGetHostByAddr Lib "ws2_32.dll" (ByVal hwnd As Long, _
    ByVal wMsg As Long, ByRef lngAddr As Long, ByVal lngLenght As Long, _
    ByVal lngType As Long, buf As Any, ByVal lngBufLen As Long) As Long
Public Declare Function WSAAsyncGetHostByName Lib "ws2_32.dll" (ByVal hwnd As Long, _
    ByVal wMsg As Long, ByVal strHostName As String, buf As Any, ByVal buflen As Long) As Long
Public Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, lpwsad As WSAData) As Long
Public Declare Function WSACleanup Lib "ws2_32.dll" () As Long

Public Type WSAData
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSA_DESCRIPTION_LEN
    szSystemStatus As String * WSA_SYS_STATUS_LEN
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Public Type sockaddr_in
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero As String * 8
End Type
