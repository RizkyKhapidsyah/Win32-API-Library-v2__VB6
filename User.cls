VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private s As Strings

Private Sub Class_Initialize()
    Set s = New Strings
End Sub

Private Sub Class_Terminate()
    Set s = Nothing
End Sub

Public Property Get Name() As String
    Dim Username As String
    Username = String$(MAX_USERNAME_LENGTH, vbNullChar)
    
    GetUserName Username, MAX_USERNAME_LENGTH + 1
    
    Username = s.TrimNull(Username)
    Name = Username
End Property

Public Property Get MyDocuments() As String
    MyDocuments = GetFolderLocation(CSIDL_PERSONAL)
End Property

Public Property Get StartMenu() As String
    StartMenu = GetFolderLocation(CSIDL_STARTMENU)
End Property

Public Property Get ApplicationData() As String
    ApplicationData = GetFolderLocation(CSIDL_APPDATA)
End Property

Public Property Get Startup() As String
    Startup = GetFolderLocation(CSIDL_STARTUP)
End Property

Public Property Get Programs() As String
    Programs = GetFolderLocation(CSIDL_PROGRAMS)
End Property

Public Property Get Recent() As String
    Recent = GetFolderLocation(CSIDL_RECENT)
End Property

Public Property Get Favorites() As String
    Favorites = GetFolderLocation(CSIDL_FAVORITES)
End Property

Public Property Get Desktop() As String
    Desktop = GetFolderLocation(CSIDL_DESKTOP)
End Property

Public Property Get SendTo() As String
    SendTo = GetFolderLocation(CSIDL_SENDTO)
End Property

Public Property Get History() As String
    History = GetFolderLocation(CSIDL_HISTORY)
End Property

Public Property Get Cookies() As String
    Cookies = GetFolderLocation(CSIDL_COOKIES)
End Property

Public Property Get MyPictures() As String
    MyPictures = GetFolderLocation(CSIDL_MYPICTURES)
End Property

Public Property Get TemporaryInternetFiles() As String
    TemporaryInternetFiles = GetFolderLocation(CSIDL_INTERNET_CACHE)
End Property

Private Function GetFolderLocation(ByVal FolderID As CSIDL) As String
    Dim localDir As String
    localDir = String$(MAX_PATH, vbNullChar)
    
    SHGetFolderPath NONESSENTIAL, FolderID, NONESSENTIAL, SHGFP_TYPE_CURRENT, localDir
    
    localDir = s.TrimNull(localDir)
    GetFolderLocation = localDir
End Function

Public Property Get Profile() As String
      Profile = GetFolderLocation(CSIDL_PROFILE)
End Property

Public Property Get NetHood() As String
      NetHood = GetFolderLocation(CSIDL_NETHOOD)
End Property

Public Property Get PrintHood() As String
      PrintHood = GetFolderLocation(CSIDL_PRINTERS)
End Property

Public Property Get Templates() As String
      Templates = GetFolderLocation(CSIDL_TEMPLATES)
End Property
