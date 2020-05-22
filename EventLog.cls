VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EventLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mAppName As String
Private mMessageDLL As String
Private mLogType As EventLogConstants
Private mComputerName As String
Private hEventLog As Long
Private mbConnected As Boolean
Private mRegistry As Registry
Private mComputer As Computer

Public Property Let AppName(strVal As String)
    mAppName = strVal
    
    ' WARNING: Do not change anything below this line
    If mLogType = elApplication Then
        mRegistry.SubKey = "CurrentControlSet\Services\EventLog\Application\" & mAppName
    End If
    
    If mLogType = elSecurity Then
        mRegistry.SubKey = "CurrentControlSet\Services\EventLog\Security\" & mAppName
    End If
    
    If mLogType = elSystem Then
        mRegistry.SubKey = "CurrentControlSet\Services\EventLog\System\" & mAppName
    End If
End Property

Public Property Get AppName() As String
    AppName = mAppName
End Property

Public Property Let MessageDLL(vData As String)
    mMessageDLL = vData
End Property

Public Property Get MessageDLL() As String
    MessageDLL = mMessageDLL
End Property

Public Property Let UseLog(vData As EventLogConstants)
    mLogType = UseLog
End Property

Public Property Get UseLog() As EventLogConstants
    UseLog = mLogType
End Property

Public Sub Connect()
    mbConnected = True
    
    ' This routine checks that the key exists
    ' This ensures that everything works ok
    If Not mRegistry.KeyExists Then
        mRegistry.CreateKey
        mRegistry.SetRegistryValue "EventMessageFile", mMessageDLL, dtRegExpandSz
        mRegistry.SetRegistryValue "TypesSupported", 7, dtRegDword
    End If
    
    mComputerName = mComputer.Name
    hEventLog = RegisterEventSource(mComputerName, mAppName)
End Sub

Private Sub Class_Initialize()
    Set mRegistry = New Registry
    Set mComputer = New Computer
    
    mbConnected = False
    mAppName = ""
    mLogType = elApplication
    mRegistry.hKey = hKeyLocalMachine
    mRegistry.KeyRoot = "System"
End Sub

Private Sub Class_Terminate()
    Set mRegistry = Nothing
    Set mComputer = Nothing
    DeregisterEventSource hEventLog
End Sub

Public Sub WriteToLog(ByVal MessageID As Long, ByVal LogType As EventLogTypeConstants, _
                                ByVal Category As EventLogCategoryConstants, _
                                Optional wNumStrings As Long, Optional ByVal lpStrings As Long)
    If Not mbConnected Then
        Connect
    End If
    ReportEvent hEventLog, LogType, Category, MessageID, ByVal 0, wNumStrings, 0, ByVal lpStrings, ByVal 0
End Sub

