Attribute VB_Name = "Module1"
Declare Function GetComputerNameA _
        Lib "kernel32" (ByVal lpBuffer As String, _
                        nSize As Long) As Long
Public strComputerID As String
Public Const LB_SETHORIZONTALEXTENT = &H194
Public Declare Function SendMessageByNum Lib "user32" _
        Alias "SendMessageA" (ByVal hwnd As Long, ByVal _
        wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
 
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

          SendMessageByNum Sender.lstLog.hwnd, LB_SETHORIZONTALEXTENT, 500, 0
End Sub
