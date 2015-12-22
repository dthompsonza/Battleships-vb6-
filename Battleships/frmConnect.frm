VERSION 5.00
Begin VB.Form frmConnect 
   BorderStyle     =   0  'None
   Caption         =   "BattleShips - Connect"
   ClientHeight    =   4125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8250
   Icon            =   "frmConnect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmConnect.frx":08CA
   ScaleHeight     =   4125
   ScaleWidth      =   8250
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   1440
      Width           =   1575
   End
   Begin VB.OptionButton optServer 
      BackColor       =   &H00D8F4E0&
      Caption         =   "I want the other player to connect to me"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Value           =   -1  'True
      Width           =   220
   End
   Begin VB.OptionButton optClient 
      BackColor       =   &H00D8F4E0&
      Caption         =   "I want to connect to another player"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   230
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   7320
      TabIndex        =   5
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "version"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   210
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Written by David Thompson  February 2003"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   360
      TabIndex        =   13
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label lblMain 
      BackStyle       =   0  'Transparent
      Caption         =   "BattleShips"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   220
      TabIndex        =   12
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblClient 
      BackStyle       =   0  'Transparent
      Caption         =   "   I want to connect to another player"
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label lblServer 
      BackStyle       =   0  'Transparent
      Caption         =   "   I want the other player to connect to me"
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Your name :"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblMyComputer 
      BackStyle       =   0  'Transparent
      Caption         =   "MyComputer"
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   1815
      Width           =   975
   End
   Begin VB.Label lblHost 
      BackStyle       =   0  'Transparent
      Caption         =   "Host Computer :"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblThis 
      BackStyle       =   0  'Transparent
      Caption         =   "This Computer :"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   1215
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Code written by David C. Thompson                       *
'*                                                         *
'* Copyright Restrictions:                                 *
'*     This code is here for reference, and may NOT be     *
'*     sold or leased under any circumstances. Changes     *
'*     will be allowed to be made so long as the original  *
'*     author's name is retained in the credits.           *
'*                                                         *
'* February 2003                                           *
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * *



Private Sub cmdOk_Click()
  Dim i%
   
  If txtName.Text = "" Then
    MsgBox "Please enter a player name!!", vbOKOnly
    Exit Sub
  End If
  
  Player = txtName.Text
  Host = False
  If optServer.Value Then
    Host = True
  End If
  If Not Host And txtHost.Text = "" Then
    MsgBox "Please enter a host computer name!!", vbOKOnly
    Exit Sub
  End If
  HostComputer = txtHost.Text
  
  SaveSetting App.Title, App.Title, "player", Player
  i = 1
  If Host Then i = 0
  SaveSetting App.Title, App.Title, "option", Str(i)
  If HostComputer <> "" Then
    SaveSetting App.Title, App.Title, "host", HostComputer
  End If
  
  FirstRun = True
  Me.Hide
  frmMain.Show
  
  Unload Me
  
End Sub

Private Sub cmdQuit_Click()
  End
End Sub


Private Sub Form_Load()
  Dim s$
  
  lblVersion = "Version " & GetVersion
  s = GetSetting(App.Title, App.Title, "player", GetUserName)
  txtName.Text = s
  s = GetSetting(App.Title, App.Title, "host", "")
  txtHost.Text = s
  s = GetSetting(App.Title, App.Title, "option", "0")
  If Val(s) = 0 Then
    optServer.Value = True
  Else
    optClient.Value = True
  End If
  
  lblMyComputer.Caption = GetTheComputerName
  NR = False
  AutoFire = False
  
  BroadcastTo = "DAA-013920"
  Broadcast = False
  
  'SendKeys "+{End}"
End Sub

Private Sub lblClient_Click()
  'optClient_Click
  optClient.Value = True
End Sub

Private Sub lblServer_Click()
  'optServer_Click
  optServer.Value = True
End Sub

Private Sub optClient_Click()
  lblHost.Visible = True
  txtHost.Visible = True
  lblThis.Visible = False
End Sub

Private Sub optServer_Click()
  lblHost.Visible = False
  txtHost.Visible = False
  lblThis.Visible = True
End Sub

Private Sub txtHost_Change()
  If Len(txtHost) > 0 And txtHost.Visible Then
    txtHost = UCase(txtHost)
    SendKeys "{End}"
  End If
End Sub
