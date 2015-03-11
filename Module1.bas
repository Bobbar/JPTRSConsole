Attribute VB_Name = "Module1"
Declare Function GetComputerNameA _
        Lib "kernel32" (ByVal lpBuffer As String, _
                        nSize As Long) As Long
Public strComputerID As String
Public Const LB_SETHORIZONTALEXTENT = &H194
Public Declare Function SendMessageByNum _
               Lib "user32" _
               Alias "SendMessageA" (ByVal hwnd As Long, _
                                     ByVal wMsg As Long, _
                                     ByVal wParam As Long, _
                                     ByVal lParam As Long) As Long
Public strBuffer As String

                                     
Public Function GetComputerName() As String
    Dim sResult As String * 255
    GetComputerNameA sResult, 255
    GetComputerName = Left$(sResult, InStr(sResult, Chr$(0)) - 1)
End Function
Public Sub Logger(Message As String, Optional Color As Long)
    Dim tmpMsg As String
    tmpMsg = Message
   
   Sender.rtbLog.Visible = False

    Sender.rtbLog.SelStart = Len(Sender.rtbLog.Text)
    If Color = 0 Then
        Sender.rtbLog.SelColor = vbWhite
    Else
        Sender.rtbLog.SelColor = Color
    End If
    Sender.rtbLog.SelText = tmpMsg & vbNewLine
      Sender.rtbLog.Visible = True
strBuffer = strBuffer + tmpMsg


   
End Sub
