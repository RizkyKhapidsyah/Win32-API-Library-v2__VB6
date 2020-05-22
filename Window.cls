VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Window"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function GetWndPlacement(ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
      lpwndpl.Length = Len(lpwndpl)
      GetWndPlacement = GetWindowPlacement(hwnd, lpwndpl)
End Function

Public Function SetWndPlacement(ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
      SetWndPlacement = SetWindowPlacement(hwnd, lpwndpl)
End Function

Public Sub Center(frm As Object, Parent As Object)
      If Parent Is Nothing Then
            frm.Top = (Screen.Height / 2) - frm.Height / 2
            frm.Left = (Screen.Width / 2) - frm.Width / 2
      Else
            frm.Top = (Parent.Top + (Parent.Height / 2)) - frm.Height / 2
            frm.Left = (Parent.Left + (Parent.Width / 2)) - frm.Width / 2
      End If
End Sub

Public Function TopMost(ByVal hwnd As Long) As Long
      TopMost = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS)
End Function

Public Function MakeNormal(ByVal hwnd As Long) As Long
      MakeNormal = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS)
End Function
