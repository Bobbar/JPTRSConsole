Attribute VB_Name = "SocketModule"
Public strSocketData As String
Public strSocketRequestID As String
Public strSocketAcceptedID As String
Private Type PacketType
ID As String
Type As String
DataString As String

    
End Type
Private PacketData As PacketType
Public Const CommandPacket As String = "COM"
Public Const RequestPacket As String = "REQ"
Public Const TerminatePacket As String = "TERM"
Public Const PasswordPacket As String = "PWD"
Public Const LogPacket As String = "LOG"
Public Const NamePacket As String = "NAME"

Public bolWaitingForPass As Boolean
Public Sub SendCommand(Command As String)
    Dim tmpString As String
    Logger Command
    
    If bolWaitingForPass Then
        tmpString = strComputerID & "," & PasswordPacket & "," & Command
    Else
        tmpString = strComputerID & "," & CommandPacket & "," & Command
    End If
    Sender.TCPClient.SendData tmpString
End Sub
Public Sub ParsePacket(Data As String)
    Dim SplitData
    Dim SplitPackets
    Dim i As Integer
    'Dim PacketData As PacketType
    
    
    SplitPackets = Split(Data, Chr$(1))

    
    For i = 1 To UBound(SplitPackets)
    
    
    SplitData = Split(SplitPackets(i), ",")
    PacketData.ID = SplitData(0)
    PacketData.Type = SplitData(1)
    PacketData.DataString = SplitData(2)
    HandlePacket PacketData
    Next i
    
    
    
   ' If PacketData.ID = strComputerID Then HandlePacket PacketData
   
    
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
    Logger "Remote Command From " & strSocketAcceptedID & ": " & Command
    
    Select Case Command
        Case "UPDATEUSERLIST"
            Logger "Updating user list..."
            RefreshUserList
        Case "CLEARQUEUE"
            ClearEmailQueueAll
        Case "UPTIME"
        
        Case "STARTREPORT DAILY"
        
        Case "STARTREPORT WEEKLY"
        
        Case "PAUSE"
        
        Case "RESUME"
        
        Case "ENDPROGRAM"
        
        Case "STATUS"
        
        Case "PASSWORD"
            CheckPassword
        Case Else
            Logger "'" & Command & "' is not a recognized command."
    End Select
End Sub
Public Sub CheckPassword(Password As String)
    On Error GoTo errs
    Dim rs          As New ADODB.Recordset
    Dim strSQL1     As String
    Dim strPassword As String
    strSQL1 = "SELECT * FROM socketvars"
    cn_global.CursorLocation = adUseClient
    rs.Open strSQL1, cn_global, adOpenKeyset
    With rs
        strPassword = !idPassword
    End With
    If Password = strPassword Then
        AcceptPassword
        SocketLog "Password accepted!"
        Logger "TCP Socket: Password accepted!"
        strSocketAcceptedID = PacketData.ID
    Else
        RejectPassword
        SocketLog "Password rejected!"
        Logger "TCP Socket: Password rejected!"
        strSocketAcceptedID = vbNullString
    End If
    Exit Sub
errs:
    ErrHandle Err.Number, Err.Description, "CheckPassword"
End Sub
Public Sub AcceptPassword()
 With JPTRS
        If .TCPServer.State <> sckClosed Then
            .TCPServer.SendData strSocketRequestID & "," & CommandPacket & ",GOODPASS"
            bolWaitingForPass = False
            
        End If
    End With
End Sub
Public Sub RejectPassword()
 With JPTRS
        If .TCPServer.State <> sckClosed Then
            .TCPServer.SendData strSocketRequestID & "," & CommandPacket & ",BADPASS"
            bolWaitingForPass = False
            
        End If
    End With
End Sub
Public Sub RequestPass()
    With JPTRS
        If .TCPServer.State <> sckClosed Then
            SocketLog "Password?"
            .TCPServer.SendData strSocketRequestID & "," & RequestPacket & ",GIVEPASS"
            bolWaitingForPass = True
            
        End If
    End With
End Sub

Public Sub PacketRequest(Request As String)

Select Case Request

    Case "BLAH"

End Select



End Sub
Public Sub SocketLog(strLog As String)
    If .TCPServer.State <> sckClosed Then
        .TCPServer.SendData strSocketRequestID & "," & LogPacket & "," & strLog
    End If
    'TCPServer.SendData strLog
End Sub
Public Sub StartTCPServer()
    On Error GoTo errs
    TCPServer.LocalPort = strListenPort
    TCPServer.Listen
    Logger "Listening on port " & strListenPort
    Exit Sub
errs:
    Logger "***** Error Starting TCP Server! *****"
    ErrHandle Err.Number, Err.Description, "StartTCPServer"
End Sub

