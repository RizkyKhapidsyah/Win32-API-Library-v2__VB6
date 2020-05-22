VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Error"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private s As Strings

Public Function ErrorHandler(err As ErrObject, Optional ByVal Buttons As VbMsgBoxStyle = vbCritical, _
      Optional ByVal Title, Optional ByVal source, Optional ByVal HelpFile, Optional ByVal Context, _
      Optional ByVal LineNumber As Long) As VbMsgBoxResult
      Dim Msg As String
      
      If IsMissing(source) Then source = err.source
      If IsMissing(LineNumber) Then LineNumber = 0
      If IsMissing(Title) Then Title = IIf(err.Number > 30000, "Application Error!", "Visual Basic Error!")
      
      Msg = "The following " & IIf(err.Number > 30000, "Application ", "Visual Basic ")
      Msg = Msg & "error has occurred:" & s.nl(2)
      
      Msg = Msg & "Error Number:" & vbTab & err.Number & s.nl(1)
      Msg = Msg & "Description:" & vbTab & err.Description & s.nl(1)
      Msg = Msg & "Source:" & vbTab & vbTab & source & s.nl(1)
      Msg = Msg & "Line Number:" & vbTab & LineNumber & s.nl(1)
      
      Msg = Msg & s.nl(1)
      Msg = Msg & "Contact your system administrator or your local support desk, and quote the message above." & s.nl(1)
      
      MsgBox Msg, Buttons, Title
      err.Clear
End Function

Private Sub Class_Initialize()
      Set s = New Strings
End Sub

Private Sub Class_Terminate()
      Set s = Nothing
End Sub
