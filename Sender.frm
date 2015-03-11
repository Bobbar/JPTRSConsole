VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Sender 
   Caption         =   "JPTRS Console"
   ClientHeight    =   6600
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12915
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Sender.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   12915
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rtbLog 
      Height          =   5535
      Left            =   60
      TabIndex        =   4
      Top             =   600
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   9763
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      TextRTF         =   $"Sender.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   360
      Left            =   1320
      TabIndex        =   3
      Top             =   120
      Width           =   990
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   360
      Left            =   180
      TabIndex        =   2
      Top             =   120
      Width           =   990
   End
   Begin MSWinsockLib.Winsock TCPClient 
      Left            =   3420
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSendCommand 
      Caption         =   "Send"
      Height          =   360
      Left            =   7320
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.TextBox txtCommand 
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   6300
      Width           =   12795
   End
End
Attribute VB_Name = "Sender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetWindow _
                Lib "user32" (ByVal hwnd As Long, _
                              ByVal wCmd As Long) As Long
Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long
Private Declare Function GetClassName _
                Lib "user32" _
                Alias "GetClassNameA" (ByVal hwnd As Long, _
                                       ByVal lpClassName As String, _
                                       ByVal nMaxCount As Long) As Long
Private Declare Function FindWindow _
                Lib "user32" _
                Alias "FindWindowA" (ByVal lpClassName As String, _
                                     ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx _
                Lib "user32" _
                Alias "FindWindowExA" (ByVal hWnd1 As Long, _
                                       ByVal hWnd2 As Long, _
                                       ByVal lpsz1 As String, _
                                       ByVal lpsz2 As String) As Long
Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Private Const WM_SETTEXT = &HC
Private Const BM_CLICK As Long = &HF5
Private Sub cmdClear_Click()
   rtbLog.Text = ""
End Sub
Private Sub cmdConnect_Click()
    ' Invoke the Connect method to initiate a
    ' connection.
    Logger ">> Connecting to " & strRemoteComputer, vbGreen
    TCPClient.Close
    TCPClient.RemoteHost = strRemoteComputer '"RemoteComputerName"
    TCPClient.RemotePort = strPort
    TCPClient.Connect
End Sub
Private Sub cmdSendCommand_Click()
    SendCommand Trim$(txtCommand.Text)
    txtCommand.Text = ""
End Sub
Private Function WindowTextGet(ByVal hwnd As Long) As String
    Dim strBuff As String, lngLen As Long
    lngLen = SendMessage(hwnd, WM_GETTEXTLENGTH, 0, 0)
    If lngLen > 0 Then
        lngLen = lngLen + 1
        strBuff = String(lngLen, vbNullChar)
        lngLen = SendMessage(hwnd, WM_GETTEXT, lngLen, ByVal strBuff)
        WindowTextGet = Left(strBuff, lngLen)
    End If
End Function
Private Function WindowTextSet(ByVal hwnd As Long, ByVal strText As String) As Boolean
    WindowTextSet = (SendMessage(hwnd, WM_SETTEXT, Len(strText), ByVal strText) <> 0)
End Function
Private Sub Form_Load()
    ' The name of the Winsock control is tcpClient.
    ' Note: to specify a remote host, you can use
    ' either the IP address (ex: "121.111.1.1") or
    ' the computer's "friendly" name, as shown here.
    strComputerID = GetComputerName
    TCPClient.RemoteHost = strRemoteComputer '"RemoteComputerName"
    TCPClient.RemotePort = strPort
    Logger "Ready..."
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    TCPClient.Close
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    rtbLog.Height = Sender.Height - 1550
    rtbLog.Width = Sender.Width - 365
    txtCommand.Top = Sender.Height - 990
    txtCommand.Width = Sender.Width - 365
End Sub
Private Sub rtbLog_KeyDown(KeyCode As Integer, Shift As Integer)
    rtbLog.SelColor = vbWhite
End Sub

Private Sub TCPClient_Close()
    TCPClient.Close
    TCPClient.RemoteHost = strRemoteComputer '"RemoteComputerName"
    TCPClient.RemotePort = strPort
End Sub
Private Sub TCPClient_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    TCPClient.GetData strData
    BuildPacket strData
End Sub

Private Sub txtCommand_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendCommand Trim$(txtCommand.Text)
        txtCommand.Text = ""
    End If
End Sub
