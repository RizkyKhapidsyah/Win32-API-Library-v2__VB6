VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private s As Strings

Private Property Get OSInfo() As OSVERSIONINFOEX
    OSInfo.dwOSVersionInfoSize = Len(OSInfo)
    GetVersionEx OSInfo
End Property

Public Property Get LongName(vData As Long) As String
      Select Case vData
            Case osWindows95: LongName = "Microsoft Windows 95"
            Case osWindows98: LongName = "Microsoft Windows 98"
            Case osWindowsME: LongName = "Microsoft Windows Millenium Edition"
            Case osWindowsNT351: LongName = "Microsoft Windows NT 3.51"
            Case osWindowsNT400: LongName = "Microsoft Windows NT 4.0"
            Case osWindows2000: LongName = "Microsoft Windows 2000"
            Case osWindowsXP: LongName = "Microsoft Windows XP"
            Case Else: LongName = "Operating System Not Known"
      End Select
End Property

Public Property Get Name() As OperatingSystemConstants
    Select Case OSInfo.dwMajorVersion
        Case mnMajorVersionWin95
            Select Case OSInfo.dwMinorVersion
                Case mnMinorVersionWin95: Name = osWindows95
                Case mnMinorVersionWin98: Name = osWindows98
                Case mnMinorVersionWinME: Name = osWindowsME
                Case mnMinorVersionWinNT400: Name = osWindowsNT400
                Case Else: Name = osUnknown
            End Select
        Case mnMajorVersionWinNT351
            Select Case OSInfo.dwMinorVersion
                Case mnMinorVersionWinNT351: Name = osWindowsNT351
                Case Else: Name = osUnknown
            End Select
        Case mnMajorVersionWin2000
            Select Case OSInfo.dwMinorVersion
                Case mnMinorVersionWin2000: Name = osWindows2000
                Case mnMinorVersionWinXP: Name = osWindowsXP
                Case mnMinorVersionWinNetServer: Name = osWindowsNetServer
                Case Else: Name = osUnknown
            End Select
        Case Else
            Name = osUnknown
    End Select
End Property

Public Property Get Build() As Long
    Build = OSInfo.dwBuildNumber
End Property

Public Property Get Major() As Long
    Major = OSInfo.dwMajorVersion
End Property

Public Property Get Minor() As Long
    Minor = OSInfo.dwMinorVersion
End Property

Public Property Get ServicePack() As String
    ServicePack = s.TrimNull(OSInfo.szCSDVersion)
End Property

Public Property Get ServicePackMinor() As Long
    ServicePackMinor = OSInfo.wServicePackMinor
End Property

Public Property Get ServicePackMajor() As Long
    ServicePackMajor = OSInfo.wServicePackMajor
End Property

Public Property Get ServicePackInstalled() As Boolean
    ServicePackInstalled = Not (ServicePack = vbNullString)
End Property

Private Sub Class_Initialize()
    Set s = New Strings
End Sub

Private Sub Class_Terminate()
    Set s = Nothing
End Sub
