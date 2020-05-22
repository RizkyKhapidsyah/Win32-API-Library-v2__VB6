VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "System"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mOS As OS
Private mComputer As Computer
Private mUser As User
Private mProgress As Progress
Private mEventLog As EventLog
Private mRegistry As Registry
Private mNetwork As Network
Private mTray As Tray
Private mSMTP As SMTP

Private Sub Class_Initialize()
    Set mOS = New OS
    Set mComputer = New Computer
    Set mUser = New User
    Set mProgress = New Progress
    Set mEventLog = New EventLog
    Set mRegistry = New Registry
    Set mNetwork = New Network
    Set mTray = New Tray
    Set mSMTP = New SMTP
End Sub

Private Sub Class_Terminate()
    Set mOS = Nothing
    Set mComputer = Nothing
    Set mUser = Nothing
    Set mProgress = Nothing
    Set mEventLog = Nothing
    Set mRegistry = Nothing
    Set mNetwork = Nothing
    Set mTray = Nothing
    Set mSMTP = Nothing
End Sub

Public Property Get SMTP() As SMTP
    Set SMTP = mSMTP
End Property

Public Property Get Tray() As Tray
    Set Tray = mTray
End Property

Public Property Get Network() As Network
    Set Network = mNetwork
End Property

Public Property Get Registry() As Registry
    Set Registry = mRegistry
End Property

Public Property Get EventLog() As EventLog
    Set EventLog = mEventLog
End Property

Public Property Get OS() As OS
    Set OS = mOS
End Property

Public Property Get Computer() As Computer
    Set Computer = mComputer
End Property

Public Property Get User() As User
    Set User = mUser
End Property

Public Property Get Progress() As Progress
      Set Progress = mProgress
End Property
