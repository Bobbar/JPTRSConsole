Attribute VB_Name = "SocketModule"
Public Const strRemoteComputer As String = "ohbre-pwadmin01"
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
Public bolWaitingForPass     As Boolean
Public Sub SendCommand(Command As String)
    Dim tmpString As String
    Logger ">> " & Command
    If bolWaitingForPass Then
        tmpString = strComputerID & "," & PasswordPacket & "," & Command
    Else
        tmpString = strComputerID & "," & CommandPacket & "," & Command
    End If
    Sender.TCPClient.SendData tmpString
End Sub
Public Sub ParsePacket(Data As String)
    On Error GoTo errs
    Dim SplitData
    Dim SplitPackets
    Dim i As Integer
    'Dim PacketData As PacketType
    SplitPackets = Split(Data, Chr$(1))
    For i = 1 To UBound(SplitPackets)
        SplitData = Split(SplitPackets(i), ",", 3)
        PacketData.ID = SplitData(0)
        PacketData.Type = SplitData(1)
        PacketData.DataString = SplitData(2)
        HandlePacket PacketData
    Next i
    ' If PacketData.ID = strComputerID Then HandlePacket PacketData
    Exit Sub
errs:
    Logger "Parser Error! Raw Data: " & Data
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
            Logger Packet.DataString
    End Select
End Sub
Public Sub GiveName()
    Dim tmpString As String
    Logger "Sending Computer Name..."
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
