VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Strings"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function TrimNull(ByVal Item As String) As String
    Dim nPos As Long
    nPos = InStr(Item, Chr$(0))
    If nPos Then Item = Left$(Item, nPos - 1)
    TrimNull = Item
End Function

Public Function StringFromPointer(ByVal lPointer As Long) As String
    Dim strTemp As String
    Dim lRetVal As Long
    
    strTemp = String$(lstrlen(ByVal lPointer), 0)
    lRetVal = lstrcpy(ByVal strTemp, ByVal lPointer)
    
    If lRetVal Then StringFromPointer = strTemp
End Function

Public Function ConCatString(lpString1 As String, lpString2 As String)
    ConCatString = lstrcat(lpString1, lpString2)
End Function

Public Function CompareString(lpString1 As String, lpString2 As String) As Long
    CompareString = lstrcmp(lpString1, lpString2)
End Function

Public Function IsUpper(ByVal cChar As Variant) As Long
    Dim bVal() As Byte
    bVal = cChar
    IsUpper = IsCharUpper(bVal(0))
End Function

Public Function IsLower(ByVal cChar As Variant) As Long
    Dim bVal() As Byte
    bVal = cChar
    IsLower = IsCharLower(bVal(0))
End Function

Public Function nl(n As Long) As String
      Dim a As Long
      Dim sCRLF As String
      
      For a = 1 To n
            sCRLF = sCRLF & vbCrLf
      Next a
      nl = sCRLF
End Function
