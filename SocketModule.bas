Attribute VB_Name = "SocketModule"
Public Const strRemoteComputer As String = "localhost" '"ohbre-pwadmin01"
Public Const strPort           As String = "1001"
Public strSocketData           As String
Public strSocketRequestID      As String
Public strSocketAcceptedID     As String
Private Type PacketType
    ID As String
    Type As String
    DataString As String
End Type
Private PacketData           As PacketType
Public Const CommandPacket   As String = "COM"
Public Const RequestPacket   As String = "REQ"
Public Const TerminatePacket As String = "TERM"
Public Const PasswordPacket  As String = "PWD"
Public Const LogPacket       As String = "LOG"
Public Const NamePacket      As String = "NAME"
Public Const ResponsePacket  As String = "RESP"
Public bolWaitingForPass     As Boolean
Private strPacket            As String
Private bolSplitPacket       As Boolean
Public Sub SendCommand(Command As String)
    On Error Resume Next
    Dim tmpString As String
    Logger ">> " & Command, vbGreen
    If bolWaitingForPass Then
        tmpString = strComputerID & "," & PasswordPacket & "," & Command
    Else
        tmpString = strComputerID & "," & CommandPacket & "," & Command
    End If
    Sender.TCPClient.SendData tmpString
End Sub
Public Sub BuildPacket(Data As String) 'compile and separate packet chunks
    Dim tmpData    As String
    Dim lngDataLen As Long, lngCurPos As Long
    tmpData = Data
    lngDataLen = Len(tmpData)
    lngCurPos = 0
    Do
        If InStr(1, tmpData, Chr$(1)) > 0 And InStr(1, tmpData, Chr$(4)) > 0 And Not bolSplitPacket Then 'if there is a start and end marker (Complete Packet)
            strPacket = Replace(Replace(Left$(tmpData, InStr(1, tmpData, Chr$(4))), Chr$(1), ""), Chr$(4), "") 'get data between markers
            ParsePacket strPacket 'send data to be parsed
            tmpData = Mid$(tmpData, InStr(1, tmpData, Chr$(4)) + 1, Len(tmpData)) 'trim out used data
        ElseIf InStr(1, tmpData, Chr$(1)) > 0 And InStr(1, tmpData, Chr$(4)) = 0 Then 'if there is a start but no end marker  (Start Packet)
            strPacket = Replace$(tmpData, Chr$(1), "") 'take all the data
            tmpData = ""
            bolSplitPacket = True
        ElseIf InStr(1, tmpData, Chr$(1)) = 0 And InStr(1, tmpData, Chr$(4)) > 0 Or bolSplitPacket Then  'if there is and end but no start  (End Packet)
            strPacket = strPacket + Replace(Left$(tmpData, InStr(1, tmpData, Chr$(4))), Chr$(4), "") 'take all data until end marker
            ParsePacket strPacket 'send data to be parsed
            tmpData = Mid$(tmpData, InStr(1, tmpData, Chr$(4)) + 1, Len(tmpData))
            bolSplitPacket = False
        ElseIf InStr(1, tmpData, Chr$(1)) = 0 And InStr(1, tmpData, Chr$(4)) = 0 Then 'if no end and no start  (Tweener Packet)
            strPacket = strPacket + tmpData 'take all the data
            tmpData = ""
            bolSplitPacket = True
        End If
    Loop Until Len(tmpData) = 0
End Sub
Public Sub ParsePacket(Data As String)
    On Error GoTo errs
    Dim SplitData
    Dim SplitPackets
    Dim i As Integer
    SplitData = Split(Data, ",", 3)
    PacketData.ID = SplitData(0)
    PacketData.Type = SplitData(1)
    PacketData.DataString = SplitData(2)
    HandlePacket PacketData
    Exit Sub
errs:
    Logger "Parser Error! Raw Data: " & Data, vbRed
    Resume Next
End Sub
Public Function AuthPacket(Packet As PacketType) As Boolean
    AuthPacket = False
    If Packet.ID = strSocketAcceptedID Then AuthPacket = True
End Function
Public Sub HandlePacket(Packet As PacketType)
    Select Case Packet.Type
        Case CommandPacket
            PacketCommand Packet.DataString
        Case RequestPacket
            Select Case Packet.DataString
                Case "GIVENAME"
            End Select
        Case TerminatePacket
        Case PasswordPacket
            If Packet.DataString = "GIVEPASS" Then
                bolWaitingForPass = True
            ElseIf Packet.DataString = "GOODPASS" Then
                bolWaitingForPass = False
                GiveName
            End If
        Case LogPacket
            Logger "<< " & Packet.DataString
        Case ResponsePacket
            Logger "<< " & Packet.DataString, vbYellow
    End Select
End Sub
Public Sub GiveName()
    Dim tmpString As String
    Logger ">> Sending Computer Name...", vbGreen
    tmpString = strComputerID & "," & NamePacket & "," & strComputerID
    Sender.TCPClient.SendData tmpString
End Sub
Public Sub PacketCommand(Command As String)
    ' Logger "Remote Command From " & strSocketAcceptedID & ": " & Command
    Select Case Command
        Case "UPDATEUSERLIST"
            ' Logger "Updating user list..."
            'RefreshUserList
        Case "CLEARQUEUE"
            'ClearEmailQueueAll
        Case "UPTIME"
        Case "STARTREPORT DAILY"
        Case "STARTREPORT WEEKLY"
        Case "PAUSE"
        Case "RESUME"
        Case "ENDPROGRAM"
        Case "STATUS"
        Case "PASSWORD"
        Case "BADPASS"
            Logger "Invalid password!"
        Case Else
            Logger "'" & Command & "' is not a recognized command."
    End Select
End Sub
