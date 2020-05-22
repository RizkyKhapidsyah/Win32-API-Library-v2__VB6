VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Math"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647
Private Const OFFSET_2 = 65536
Private Const MAXINT_2 = 32767

Public Function UnsignedToLong(Value As Double) As Long
    If Value < 0 Or Value >= OFFSET_4 Then Error 6
    
    If Value <= MAXINT_4 Then
        UnsignedToLong = Value
    Else
        UnsignedToLong = Value - OFFSET_4
    End If
End Function

Public Function LongToUnsigned(Value As Long) As Double
    If Value < 0 Then
        LongToUnsigned = Value + OFFSET_4
    Else
        LongToUnsigned = Value
    End If
End Function

Public Function UnsignedToInteger(Value As Long) As Integer
    If Value < 0 Or Value >= OFFSET_2 Then Error 6
    
    If Value <= MAXINT_2 Then
        UnsignedToInteger = Value
    Else
        UnsignedToInteger = Value - OFFSET_2
    End If
End Function

Public Function IntegerToUnsigned(Value As Integer) As Long
    If Value < 0 Then
        IntegerToUnsigned = Value + OFFSET_2
    Else
        IntegerToUnsigned = Value
    End If
End Function

Public Function Mod7(ByVal Value, Optional ByVal Check As Long = 7) As Long
      Dim RetVal
      Dim saryResult
      Dim i As Long
      
      For i = 1 To Len(Value)
            If IsNumeric(Mid(Value, i, 1)) Then
                  RetVal = RetVal & Mid(Value, i, 1)
            End If
      Next i
      
      RetVal = RetVal / Check
      saryResult = Split(RetVal, ".")
      RetVal = "." & saryResult(1)
      
      If RetVal = "." Then RetVal = 0
      RetVal = RetVal * Check
      Mod7 = RetVal
End Function
