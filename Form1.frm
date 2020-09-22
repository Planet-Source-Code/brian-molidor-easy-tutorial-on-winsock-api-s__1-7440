VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Example using winsock api's"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   9390
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtStatus 
      Height          =   4215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   960
      Width           =   9135
   End
   Begin VB.Frame Frame1 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9375
      Begin VB.CommandButton Command3 
         Caption         =   "&Send"
         Height          =   315
         Left            =   8040
         TabIndex        =   10
         Top             =   5280
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   5280
         Width           =   7815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Dissconnect"
         Height          =   375
         Left            =   6720
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   2280
         TabIndex        =   6
         Text            =   "6666"
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Connect"
         Height          =   375
         Left            =   8040
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Text            =   "irc.dalnet.com"
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Port:"
         Height          =   255
         Left            =   2280
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "&IRC Server:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   2175
      Left            =   3480
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This example was done by Brian Molidor
'PulseWave@aol.com
'url: coming soon!

Private Sub Command1_Click()
    'make sure the port is closed!
    If mysock <> 0 Then Call closesocket(mysock)
    'let's connect!!!       host            port       handle
    Call ConnectSock(CStr(Text1.Text), CLng(Text2.Text), 0, Form1.hWnd, True)
End Sub

Private Sub Command2_Click()
    'lets close the connection
    Call closesocket(mysock)
    'set mysock = to 0
    mysock = 0
End Sub

Private Sub Command3_Click()
    If mysock <> 0 Then 'make sure we are connected
        Call SendData(mysock, Text3.Text & vbCrLf) 'send the data
        Text3.Text = ""
    End If
End Sub

Private Sub Form_Load()
    'ok, we have to start winsock, DUH!
    Call StartWinsock("")
    'lets subclassing the handle
    'for the connection we are going to make
    Call Hook(Form1.hWnd)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'lets close the connection
    Call closesocket(mysock)
    'lets unhook the hwnd so we dont
    'get an error
    Call UnHook(Form1.hWnd)
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    'check to see if enter was pushed, if so
    'call command3_click, which sends the data
    If KeyAscii = 13 Then Command3_Click
End Sub

Private Sub txtStatus_Change()
    'keep the txtbox at the very bottom at all times
    txtStatus.SelStart = Len(txtStatus)
End Sub
