VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Computer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private s As Strings

Public Property Get Name() As String
    Dim localName As String
    localName = String$(MAX_COMPUTERNAME_LENGTH, vbNullChar)
    
    GetComputerName localName, MAX_COMPUTERNAME_LENGTH + 1
    
    localName = s.TrimNull(localName)
    Name = localName
End Property

Private Sub Class_Initialize()
    Set s = New Strings
End Sub

Private Sub Class_Terminate()
    Set s = Nothing
End Sub

Private Function GetFolderLocation(ByVal FolderID As CSIDL) As String
    Dim localDir As String, lRet As Long
    localDir = String$(MAX_PATH, vbNullChar)
      
    lRet = SHGetFolderPath(NONESSENTIAL, FolderID, NONESSENTIAL, _
        SHGFP_TYPE_CURRENT, localDir)
      
    localDir = s.TrimNull(localDir)
    GetFolderLocation = localDir
End Function

Public Property Get WindowsDirectory() As String
    WindowsDirectory = GetFolderLocation(CSIDL_WINDOWS)
End Property

Public Property Get CurrentDirectory() As String
    Dim localDir As String
    localDir = String$(MAX_PATH, vbNullChar)
    
    GetCurrentDirectory MAX_PATH + 1, localDir
    
    localDir = s.TrimNull(localDir)
    CurrentDirectory = localDir
End Property

Public Property Get SystemDirectory() As String
    SystemDirectory = GetFolderLocation(CSIDL_SYSTEM)
End Property

Public Property Get ProgramFiles() As String
    ProgramFiles = GetFolderLocation(CSIDL_PROGRAM_FILES)
End Property

Public Property Get RecycleBin() As String
    RecycleBin = GetFolderLocation(CSIDL_BITBUCKET)
End Property

Public Property Get Fonts() As String
    Fonts = GetFolderLocation(CSIDL_FONTS)
End Property

Public Property Get CommonFiles() As String
    CommonFiles = GetFolderLocation(CSIDL_PROGRAM_FILES_COMMON)
End Property

Public Property Get Printers() As String
    Printers = GetFolderLocation(CSIDL_PRINTERS)
End Property
