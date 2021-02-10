VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form Form1 
   Caption         =   "Server Socket"
   ClientHeight    =   4785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5760
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Conection"
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   6255
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Text            =   "Text3"
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Start Server"
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1200
         Width           =   4695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Send"
         Height          =   375
         Left            =   5040
         TabIndex        =   3
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Port:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   5040
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iSockets As Integer
Dim sServerMsg As String
Dim sRequestID As String
   
Private Sub Command1_Click()
    If Socket(iSockets).State = sckConnected Then
        Socket(iSockets).SendData Text1.Text
    Else
        MsgBox "Sem conexão!!", vbOKOnly
    End If
End Sub

Private Sub Command2_Click()
    If Socket(iSockets).State <> 0 Then
        Socket(iSockets).Close
    End If
    If Trim(Text3.Text) <> "" And IsNumeric(Text3.Text) Then
        'Socket(0).LocalPort = 1007
        Socket(0).LocalPort = Trim(Text3.Text)
        sServerMsg = "Listening to port: " & Socket(0).LocalPort
        List1.AddItem (sServerMsg)
        Socket(0).Listen
        Label3.Caption = Socket(iSockets).State
    Else
        MsgBox "Porta invalida", vbOKOnly
        Text3.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Form1.Show
    'Text2.Text = Socket(0).LocalIP
'    Socket(0).LocalPort = 1007
'    sServerMsg = "Listening to port: " & Socket(0).LocalPort
'    List1.AddItem (sServerMsg)
'    Socket(0).Listen
'    Label1.Caption = Socket(iSockets).State
End Sub

Private Sub socket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    sServerMsg = "Connection request id " & requestID & " from " & Socket(Index).RemoteHostIP
    'If Index = 0 Then
        List1.AddItem (sServerMsg)
        sRequestID = requestID
        iSockets = iSockets + 1
        Load Socket(iSockets)
        Socket(iSockets).LocalPort = 1007
        Socket(iSockets).Accept requestID
'        Label2.Caption = Socket(iSockets).State
'        Label1.Caption = Socket(0).State
'        Label3.Caption = iSockets
        DoEvents
    'End If
End Sub

Private Sub socket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim sItemData As String
Dim strOutData As String
        
    ' get data from client
    Socket(Index).GetData sItemData, vbString
    sServerMsg = "Received: " & sItemData & " from " & Socket(Index).RemoteHostIP & "(" & sRequestID & ")"
    List1.AddItem (sServerMsg)
   
    'send data to client
    sServerMsg = "Sending: " & strOutData & " to " & Socket(Index).RemoteHostIP
    List1.AddItem (sServerMsg)
    'strOutData = "Recebido a mensagem: " sServerMsg
    Socket(Index).SendData sServerMsg
End Sub

