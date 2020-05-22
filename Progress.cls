VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mProgressCtrl As Object
Private mProgressStatus As Object
Private mProgressStatusCaption As String

Private mlMinValue As Long
Private mlMaxValue As Long
Private mlStep As Long

Private mlPrevPercent As Long
Private mlPrevPosition As Long

Public Property Set ProgressControl(vData As Object)
      Set mProgressCtrl = vData
End Property

Public Property Set ProgressStatusLabel(vData As Object)
      Set mProgressStatus = vData
End Property

Public Property Let ProgressStatusCaption(vData As String)
      mProgressStatusCaption = vData
End Property

Public Function SetRange(ByVal lMin As Long, ByVal lMax As Long, Optional ByVal Step As Long = 1) As Long
      mlMinValue = lMin
      mlMaxValue = lMax
      mlStep = Step
      
      If mProgressCtrl Is Nothing Then SetRange = -1: Exit Function
      mProgressCtrl.Min = mlMinValue
      mProgressCtrl.Max = mlMaxValue
      mProgressCtrl.Value = 0
      
      If Not mProgressStatus Is Nothing Then
            mProgressStatus.Caption = mProgressStatusCaption & " [0%]"
      End If
End Function

Public Function StepIt() As Long
      StepIt = SetPos(mlPrevPosition + mlStep)
End Function

Private Function SetPos(ByVal nPos As Long) As Long
      Dim strTitle As String
      Dim lPercentage As Long
      
      DoEvents
      If mProgressCtrl Is Nothing Then SetPos = -1: Exit Function
      
      mlPrevPosition = nPos
      
      If mlMaxValue > mlMinValue Then
            lPercentage = (nPos * 100) / (mlMaxValue - mlMinValue)
      Else
            lPercentage = 0
      End If
      
      If mlPrevPercent <> lPercentage Then
            mlPrevPercent = lPercentage
            strTitle = " [" & lPercentage & "%]"
            If Not mProgressStatus Is Nothing Then
                  mProgressStatus.Caption = mProgressStatusCaption & strTitle
            End If
      End If
      
      mProgressCtrl.Value = nPos
      SetPos = nPos
End Function

Public Sub Reset()
      SetRange mlMinValue, mlMaxValue, mlStep
      
      mlPrevPercent = 0
      mlPrevPosition = 0
End Sub

Private Sub Class_Initialize()
      Set mProgressCtrl = Nothing
      Set mProgressStatus = Nothing
      mProgressStatusCaption = "Please Wait, Processing..."
      
      mlMinValue = 0
      mlMaxValue = 100
      mlStep = 1
End Sub

Private Sub Class_Terminate()
      Set mProgressCtrl = Nothing
      Set mProgressStatus = Nothing
End Sub
