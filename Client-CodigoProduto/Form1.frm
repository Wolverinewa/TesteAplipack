VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form Form1 
   Caption         =   "Dados Codigo produto"
   ClientHeight    =   4515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   ScaleHeight     =   10215
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCod 
      Height          =   405
      Left            =   1680
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txtIP 
      Height          =   405
      Left            =   2040
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtPort 
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   840
      Width           =   855
   End
   Begin VB.ListBox List2 
      Height          =   2010
      Left            =   4560
      TabIndex        =   5
      Top             =   2400
      Width           =   4095
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Conectar"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send data"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3960
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label7 
      Caption         =   "Código do Produto"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "IP:"
      Height          =   255
      Left            =   1800
      TabIndex        =   9
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Port:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim processando As Boolean
Dim idx As Long
Dim sClientMsg As String

Private Sub cmdClose_Click()
    Winsock1.Close
End Sub

Private Sub Command1_Click()
Dim i As Long

    i = 0
    List1.Clear
    
    Do While processando
        If Winsock1.State = sckConnected Then
            Winsock1.SendData txtCod.Text
            List1.AddItem "Sending Data: " & txtCod.Text
            DoEvents
        Else
            Label3.Caption = "Not currently connected to host"
        End If
        If i > 10 Then
            Exit Do
        End If
    Loop
    
End Sub

Private Sub Command2_Click()
    Winsock1.RemoteHost = Trim(txtIP.Text)
    Winsock1.RemotePort = Trim(txtPort.Text)
    Winsock1.Connect
    Label3.Caption = "Status: " & Winsock1.State
    DoEvents
End Sub

Private Sub Form_Load()
    List1.Clear
    processando = True
    idx = 0
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim sData As String
    Winsock1.GetData sData, vbString
    Label1.Caption = "Received Data: " & sData
    List2.AddItem "Received Data: " & sData
End Sub

Private Sub Winsock1_SendComplete()

    Label3.Caption = "Completed Data Transmission"

End Sub

