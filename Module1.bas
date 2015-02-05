Attribute VB_Name = "Module1"
Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
Public strComputerID As String


Public Function GetComputerName() As String
Dim sResult As String * 255
    GetComputerNameA sResult, 255
    GetComputerName = Left$(sResult, InStr(sResult, Chr$(0)) - 1)
End Function

Public Sub Logger(Message As String)
Dim tmpMsg As String
tmpMsg = DateTime.Date & " " & DateTime.Time & ": " & Message
Sender.lstLog.AddItem tmpMsg, Sender.lstLog.ListCount
Sender.lstLog.ListIndex = Sender.lstLog.NewIndex
End Sub
