VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Memory"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' THIS WORKS FINE
Public Sub ZeroMemory(Destination, ByVal Length As Long)
    RtlZeroMemory Destination, Length
End Sub

' TODO: GET THESE TWO TO WORK
Public Sub MoveMemory(Destination, source, ByVal Length As Long)
    RtlMoveMemory Destination, source, Length
End Sub

Public Sub CopyMemory(Destination, source, ByVal Length As Long)
    RtlMoveMemory Destination, source, Length
End Sub
