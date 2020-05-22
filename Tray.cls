VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarKey As String

Public Sub AddNew(hwnd As Long, ByVal uID As Long, _
                            Optional ByVal Icon As stdole.IPictureDisp, _
                            Optional ByVal TrayStyle As TrayStyleConstants = tsSimple, _
                            Optional ByVal Message As String, Optional ByVal Title As String, _
                            Optional ByVal MessageIcon As TrayMessageIconConstants = tmiNone, _
                            Optional ByVal CallbackMessage As MouseMessageConstants = mmMouseOver, _
                            Optional ByVal Delay As Long = &H0)
    TrayIcon tmAdd, hwnd, uID, Icon, TrayStyle, Message, Title, _
        MessageIcon, CallbackMessage, Delay
        
    Key = hwnd
End Sub

Public Sub Update(hwnd As Long, ByVal uID As Long, _
                            Optional ByVal Icon As stdole.IPictureDisp, _
                            Optional ByVal TrayStyle As TrayStyleConstants = tsSimple, _
                            Optional ByVal Message As String, Optional ByVal Title As String, _
                            Optional ByVal MessageIcon As TrayMessageIconConstants = tmiNone, _
                            Optional ByVal CallbackMessage As MouseMessageConstants = mmMouseOver, _
                            Optional ByVal Delay As Long = &H0)
    TrayIcon tmModify, hwnd, uID, Icon, TrayStyle, Message, Title, _
        MessageIcon, CallbackMessage, Delay
    
    Key = hwnd
End Sub

Public Sub Remove(ByVal uID As Long)
    TrayIcon tmDelete, Key, uID
End Sub

Private Sub TrayIcon(ByVal TrayMessage As TrayMessageConstants, _
                            hwnd As Long, ByVal uID As Long, _
                            Optional ByVal Icon As stdole.IPictureDisp, _
                            Optional ByVal TrayStyle As TrayStyleConstants = tsSimple, _
                            Optional ByVal Message As String, Optional ByVal Title As String, _
                            Optional ByVal MessageIcon As TrayMessageIconConstants = tmiNone, _
                            Optional ByVal CallbackMessage As MouseMessageConstants = mmMouseOver, _
                            Optional ByVal Delay As Long = &H0)
    Dim nid As NOTIFYICONDATA
    Dim Result As Long
    
    With nid
        .cbSize = Len(nid)
        .hwnd = hwnd
        .uID = uID
        .uFlags = (tiIcon Or tiMessage)
        If Not Icon Is Nothing Then .hIcon = Icon
        .uCallbackMessage = CallbackMessage
        .dwState = 0
        .dwStateMask = 0
        
        Select Case TrayStyle
            Case tsSimple
                .uFlags = (.uFlags Or tiTip)
                .szTip = Message & vbNullChar
            Case tsBalloon
                .uFlags = (.uFlags Or tiTip Or tiInfo)
                .szInfo = Message & vbNullChar
                .szInfoTitle = Title & vbNullChar
                .uTimeoutOrVersion = (Delay)
                .dwInfoFlags = MessageIcon
        End Select
    End With
    
    Result = ShellNotifyIcon(TrayMessage, nid)
    If TrayMessage = tmModify And Result = 0 Then
        Result = ShellNotifyIcon(tmAdd, nid)
    End If
End Sub

Public Property Get Key() As Variant
    Key = mvarKey
End Property

Public Property Let Key(vData As Variant)
    mvarKey = vData
End Property
