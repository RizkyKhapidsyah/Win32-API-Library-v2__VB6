VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Network"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mStrings As Strings

'********************************************************************************
'
' FUNCTION: MapNetworkDrive()
'
' RETURN: Long
'
' PARAMS: LocalName As String, RemoteName As String,
'                 Username As String, Password As String,
'                 Flags As ConnectionOptionConstants
'
' USE: Maps a network share to a local path resource, the local resource
'          must be available otherwise you will get a WN_ALREADY_CONNECTED
'          error, if Username and Password as Null then the user will be promted
'          to enter them
'
'********************************************************************************
Public Function MapNetworkDrive(ByVal LocalName As String, ByVal RemoteName As String, ByVal Username As String, ByVal Password As String, Optional ByVal Flags As ConnectOptionConstants = coConnectInteractive) As Long
      Dim lRet As Long
      Dim lpnr As NETRESOURCE
      
      With lpnr
            .dwScope = nrResourceConnected Or nrResourceRemembered
            .dwType = rtResourceTypeDisk
            .dwDisplayType = rdResourceDisplayTypeGeneric
            .dwUsage = ruResourceUsageAll
            .lpLocalName = LocalName
            .lpRemoteName = RemoteName
      End With
      
      lRet = WNetAddConnection(lpnr, Password, Username, Flags)
      MapNetworkDrive = lRet
End Function

'********************************************************************************
'
' FUNCTION: DisconnectNetworkDrive()
'
' RETURN: Long
'
' PARAMS: Name As String, Flags As ConnectionOptionConstants,
'                 Force As Long
'
' USE: Disconnects a local named resource
'
'********************************************************************************
Public Function DisconnectNetworkDrive(ByVal Name As String, ByVal Flags As ConnectOptionConstants, ByVal Force As Long) As Long
      DisconnectNetworkDrive = WNetCancelConnection(Name, Flags, Force)
End Function

'********************************************************************************
'
' FUNCTION: GetNetUser()
'
' RETURN: Long
'
' PARAMS: Name As String, Username As String, Length As Long
'
' USE: Returns the network username for the resource named in Name,
'          if Name is Null then this function returns the active username.
'
'********************************************************************************
Public Function GetNetUser(ByVal Name As String, Username As String, ByVal Length As Long) As Long
      Username = Space$(Length + 1)
      GetNetUser = WNetGetUser(Name, Username, Length)
      Username = mStrings.TrimNull(Username)
End Function

'********************************************************************************
'
' FUNCTION: NetConnectDialog()
'
' RETURN: Long
'
' PARAMS: hwnd As Long, dwType As ResourceTypeConstants
'
' USE: Displays the MapNetworkDrive dialog to the user
'
'********************************************************************************
Public Function NetConnectDialog(ByVal hwnd As Long, ByVal dwType As ResourceTypeConstants) As Long
      NetConnectDialog = WNetConnectDialog(hwnd, dwType)
End Function

'********************************************************************************
'
' FUNCTION: DisconnectConnectDialog()
'
' RETURN: Long
'
' PARAMS: hwnd As Long, dwType As ResourceTypeConstants
'
' USE: Displays the DisconnectNetworkDrive dialog to the user
'
'********************************************************************************
Public Function NetDisconnectDialog(ByVal hwnd As Long, ByVal dwType As ResourceTypeConstants) As Long
      NetDisconnectDialog = WNetDisconnectDialog(hwnd, dwType)
End Function

'********************************************************************************
'
' FUNCTION: GetNetConnection()
'
' RETURN: Long
'
' PARAMS: LocalName As String, RemoteName As String,
'                  cbRemoteName As Long
'
' USE: Returns the Remote name for the resource specified in
'          LocalName, remote name will be a network path to a server share
'          in the format \\server\share
'
'********************************************************************************
Public Function GetNetConnection(ByVal lpszLocalName As String, lpszRemoteName As String, cbRemoteName As Long) As Long
      lpszRemoteName = Space$(cbRemoteName + 1)
      GetNetConnection = WNetGetConnection(lpszLocalName, lpszRemoteName, cbRemoteName)
      lpszRemoteName = mStrings.TrimNull(lpszRemoteName)
End Function

Private Sub Class_Initialize()
    Set mStrings = New Strings
End Sub
