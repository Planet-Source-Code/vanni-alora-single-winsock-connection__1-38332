VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client - Winsock"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2790
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   2790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox txtPort 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox txtIP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin MSWinsockLib.Winsock wskClient 
      Left            =   2040
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblStatus 
      Caption         =   "Ready for Action..."
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblPort 
      Caption         =   "Server Port:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lblIP 
      Caption         =   "Server IP:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' author: Vanni Alora
' email:  vanjo08@msn.com
' url:    not yet available
' this is my first submission to PSC, im a newbie in winsock and hope
' you may find it usefull and help you specially for beginners like me.
' i need comments, suggestions, or criticism about my little app...
' i dont expect your votes... thanx and God Bless...
' my credits goes to my friends Cris "Coding Genius" Waddell and...
' Yariv Sarafraz for helping me learn about Winsock, thanx guys.

Option Explicit

' define global variables...
Dim Port As String
Dim svrIP As String

Private Sub cmdConnect_Click()
    ' get the variable...
    svrIP = txtIP.Text
    Port = txtPort.Text
    
    ' close any current and used connections...
    wskClient.Close
    
    ' set the IP and Port to connect...
    wskClient.RemoteHost = svrIP
    wskClient.RemotePort = Port
    ' start to connect...
    wskClient.Connect
    
    ' disable the "Connect" Button and enable the "Disconnect" Button...
    cmdConnect.Enabled = False
    cmdDisconnect.Enabled = True
    ' display the status...
    lblStatus.Caption = "Searching for Server..."
    ' locked the texboxes for any changes...
    txtIP.Enabled = False
    txtPort.Enabled = False
End Sub

Private Sub cmdDisconnect_Click()
    ' close the current connection...
    wskClient.Close
    
    ' disable the "Connect" Button and enable the "Disconnect" Button...
    cmdConnect.Enabled = True
    cmdDisconnect.Enabled = False
    ' display the status...
    lblStatus.Caption = "Disconnected to Server..."
    ' unlock the textboxes...
    txtIP.Enabled = True
    txtPort.Enabled = True
End Sub

Private Sub Form_Load()
    ' disable the "Disconnect" Button...
    cmdDisconnect.Enabled = False
    
    ' display the default IP and Port in the textbox...
    txtIP.Text = "127.0.0.1"
    txtPort.Text = "11898"
End Sub

Private Sub wskClient_Close()
    ' if client is disconnected and try to connect again...
    If wskClient.State <> sckClosed Then wskClient.Close
    MsgBox "Connection to Server lost..."
    
    ' call the "cmdDisconnect_Click" Event...
    cmdDisconnect_Click
End Sub

Private Sub wskClient_Connect()
    ' if Client is successfully connected, then display the status...
     If wskClient.State <> sckClosed Then
        lblStatus.Caption = "Connected to Server..."
    End If
End Sub

Private Sub wskClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ' if any error occurs, then display a message...
    MsgBox "Cannot connect to server..."
    
    ' call the "cmdDisconnect_Click" Event...
    cmdDisconnect_Click
End Sub
