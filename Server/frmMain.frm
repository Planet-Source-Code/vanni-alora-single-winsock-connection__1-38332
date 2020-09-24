VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server - Winsock"
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
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop Listen"
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
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Listen"
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
   Begin VB.TextBox txtServerPort 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox txtServerIP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin MSWinsockLib.Winsock wskServer 
      Left            =   2040
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblStatus 
      Caption         =   "Server Closed..."
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblServerPort 
      Caption         =   "Server Port:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lblServerIP 
      Caption         =   "Server IP Address:"
      Height          =   255
      Left            =   360
      TabIndex        =   1
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

' define global variable...
Dim svrPort As String

Private Sub cmdStart_Click()
    ' get the variables...
    svrPort = txtServerPort.Text
    
    ' close all current open connection...
    wskServer.Close
    ' set the Port for listening...
    wskServer.LocalPort = svrPort
    ' start the server for listening...
    wskServer.Listen
    
    ' disable the "Start Listen" Button and enable the "Stop Listen" Button...
    cmdStart.Enabled = False
    cmdStop.Enabled = True
    ' disable the Port textbox for any changes...
    txtServerPort.Enabled = False
    ' set the status of the server...
    lblStatus.Caption = "Waiting for connections..."
End Sub

Private Sub cmdStop_Click()
    ' close the current connection...
    wskServer.Close
    
    ' display the status...
    lblStatus.Caption = "Server Closed..."
    
    ' enable and disable the buttons and textbox...
    cmdStart.Enabled = True
    cmdStop.Enabled = False
    txtServerPort.Enabled = True
End Sub

Private Sub Form_Load()
    ' set "Stop Listen" Button to disable...
    cmdStop.Enabled = False
    
    ' display Local IP Address in textbox...
    txtServerIP.Text = wskServer.LocalIP
    ' display the default listening port...
    txtServerPort.Text = "11898"
End Sub

Private Sub wskServer_Close()
    ' if the client is disconnected and try to connect again, then do this...
    If wskServer.State <> sckClosed Then wskServer.Close
    
    ' call the "cmdStart_Click" Event...
    cmdStart_Click
End Sub

Private Sub wskServer_ConnectionRequest(ByVal requestID As Long)
    ' close any used control...
    wskServer.Close
    
    ' accept the incoming connection...
    wskServer.Accept requestID
    ' display the status...
    lblStatus.Caption = "Connection Successfull..."
End Sub

Private Sub wskServer_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ' if any error occured, then display a message...
    MsgBox "Unexpected error occured...", vbCritical + vbOKOnly, "Error"
End Sub
