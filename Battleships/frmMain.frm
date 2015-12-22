VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C09440&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BattleShips"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10020
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   386
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   668
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSendW2K 
      Caption         =   "Send W2K"
      Height          =   375
      Left            =   7080
      TabIndex        =   32
      ToolTipText     =   "Sends a messages using the command ""NET SEND"". For NT based systems only!"
      Top             =   4560
      Width           =   975
   End
   Begin VB.Timer timBroadcast 
      Left            =   3600
      Top             =   0
   End
   Begin VB.CommandButton cmdSurrender 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Surrender"
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdAutoFire 
      Caption         =   "auto-fire"
      Height          =   375
      Left            =   600
      TabIndex        =   24
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Timer timAutoFire 
      Interval        =   500
      Left            =   2640
      Top             =   0
   End
   Begin VB.CommandButton cmdFire 
      Caption         =   "Fire"
      Height          =   375
      Left            =   7080
      TabIndex        =   21
      ToolTipText     =   "Enter coordinates here or click on the grid to fire!"
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox txtOppChat 
      BackColor       =   &H00B05D04&
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4200
      MaxLength       =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Top             =   3720
      Width           =   3855
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "hide grid"
      Height          =   375
      Left            =   1800
      TabIndex        =   19
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdReady 
      BackColor       =   &H00C0FFC0&
      Caption         =   "ready"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdLayout 
      Caption         =   "place ships"
      Height          =   375
      Left            =   600
      TabIndex        =   18
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdSendChat 
      Caption         =   "Send"
      Height          =   375
      Left            =   6360
      TabIndex        =   17
      ToolTipText     =   "Sends a message via the games networking"
      Top             =   4560
      Width           =   615
   End
   Begin VB.TextBox txtChat 
      BackColor       =   &H00B05D04&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4200
      MaxLength       =   120
      TabIndex        =   16
      Top             =   4560
      Width           =   2055
   End
   Begin MSComctlLib.ImageList ilsPics 
      Left            =   1560
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   18
      ImageHeight     =   18
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   75
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   "1_water"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D0C
            Key             =   "5_water"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":114E
            Key             =   "5_afthHIT"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1590
            Key             =   "5_aftv"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19D2
            Key             =   "5_aftvHIT"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E14
            Key             =   "5_bowh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2256
            Key             =   "5_bowhHIT"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2698
            Key             =   "5_bowv"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2ADA
            Key             =   "5_bowvHIT"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F1C
            Key             =   "5_hit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":335E
            Key             =   "5_midh"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":37A0
            Key             =   "5_midhHIT"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3BE2
            Key             =   "5_midv"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4024
            Key             =   "5_midvHIT"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4466
            Key             =   "5_miss"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":48A8
            Key             =   "5_afth"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4CEA
            Key             =   "4_water"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":512C
            Key             =   "4_afthHIT"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":556E
            Key             =   "4_aftv"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":59B0
            Key             =   "4_aftvHIT"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5DF2
            Key             =   "4_bowh"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6234
            Key             =   "4_bowhHIT"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6676
            Key             =   "4_bowv"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6AB8
            Key             =   "4_bowvHIT"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6EFA
            Key             =   "4_hit"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":733C
            Key             =   "4_midh"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":777E
            Key             =   "4_midhHIT"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7BC0
            Key             =   "4_midv"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8002
            Key             =   "4_midvHIT"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8444
            Key             =   "4_miss"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8886
            Key             =   "4_afth"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8CC8
            Key             =   "3_water"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":910A
            Key             =   "3_afthHIT"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":954C
            Key             =   "3_aftv"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":998E
            Key             =   "3_aftvHIT"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9DD0
            Key             =   "3_bowh"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A212
            Key             =   "3_bowhHIT"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A654
            Key             =   "3_bowv"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AA96
            Key             =   "3_bowvHIT"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AED8
            Key             =   "3_hit"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B31A
            Key             =   "3_midh"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B75C
            Key             =   "3_midhHIT"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BB9E
            Key             =   "3_midv"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BFE0
            Key             =   "3_midvHIT"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C422
            Key             =   "3_miss"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C864
            Key             =   "3_afth"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CCA6
            Key             =   "2_water"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D0E8
            Key             =   "2_afthHIT"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D52A
            Key             =   "2_aftv"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D96C
            Key             =   "2_aftvHIT"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DDAE
            Key             =   "2_bowh"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E1F0
            Key             =   "2_bowhHIT"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E632
            Key             =   "2_bowv"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EA74
            Key             =   "2_bowvHIT"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EEB6
            Key             =   "2_hit"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F2F8
            Key             =   "2_midh"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F73A
            Key             =   "2_midhHIT"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FB7C
            Key             =   "2_midv"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FFBE
            Key             =   "2_midvHIT"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10400
            Key             =   "2_miss"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10842
            Key             =   "2_afth"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10C84
            Key             =   "1_afthHIT"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":110C6
            Key             =   "1_aftv"
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11508
            Key             =   "1_aftvHIT"
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1194A
            Key             =   "1_bowh"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11D8C
            Key             =   "1_bowhHIT"
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":121CE
            Key             =   "1_bowv"
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12610
            Key             =   "1_bowvHIT"
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12A52
            Key             =   "1_hit"
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12E94
            Key             =   "1_midh"
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":132D6
            Key             =   "1_midhHIT"
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13718
            Key             =   "1_midv"
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13B5A
            Key             =   "1_midvHIT"
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13F9C
            Key             =   "1_miss"
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":143DE
            Key             =   "1_afth"
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstHits 
      BackColor       =   &H00B05D04&
      ForeColor       =   &H00FFFFFF&
      Height          =   1620
      Left            =   9000
      TabIndex        =   9
      Top             =   2400
      Width           =   735
   End
   Begin VB.ListBox lstShots 
      BackColor       =   &H00B05D04&
      ForeColor       =   &H00FFFFFF&
      Height          =   1620
      Left            =   8160
      TabIndex        =   8
      Top             =   2400
      Width           =   735
   End
   Begin VB.Timer timUpdate 
      Interval        =   50
      Left            =   3120
      Top             =   0
   End
   Begin VB.CommandButton cmdChangeShipGrid 
      Caption         =   "random ships"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   3720
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock wsNet 
      Left            =   2160
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.TextBox txtFire 
      Height          =   375
      Left            =   7080
      MaxLength       =   2
      TabIndex        =   0
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label lblShotHitRatio 
      BackColor       =   &H00C09440&
      Caption         =   "-"
      Height          =   255
      Left            =   8400
      TabIndex        =   31
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C09440&
      Caption         =   "Shot/Hit Ratio :"
      Height          =   255
      Left            =   7200
      TabIndex        =   30
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblStatus2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00B05D04&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4200
      TabIndex        =   29
      Top             =   5280
      Width           =   5655
   End
   Begin VB.Shape shpTurn 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   3900
      Shape           =   3  'Circle
      Top             =   450
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C09440&
      Caption         =   "Firing Grid"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4080
      TabIndex        =   27
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C09440&
      Caption         =   "ShipGrid"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   360
      TabIndex        =   26
      Top             =   120
      Width           =   1455
   End
   Begin VB.Image Image3 
      Height          =   270
      Left            =   4200
      Picture         =   "frmMain.frx":14820
      Top             =   435
      Width           =   2835
   End
   Begin VB.Image Image2 
      Height          =   2835
      Left            =   3915
      Picture         =   "frmMain.frx":17052
      Top             =   720
      Width           =   270
   End
   Begin VB.Image Image1 
      Height          =   2835
      Left            =   195
      Picture         =   "frmMain.frx":199EC
      Top             =   720
      Width           =   270
   End
   Begin VB.Image Image4 
      Height          =   270
      Left            =   480
      Picture         =   "frmMain.frx":1C386
      Top             =   435
      Width           =   2835
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   100
      Left            =   6765
      Top             =   3285
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   99
      Left            =   6480
      Top             =   3285
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   98
      Left            =   6195
      Top             =   3285
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   97
      Left            =   5910
      Top             =   3285
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   96
      Left            =   5625
      Top             =   3285
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   95
      Left            =   5340
      Top             =   3285
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   94
      Left            =   5055
      Top             =   3285
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   93
      Left            =   4770
      Top             =   3285
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   92
      Left            =   4485
      Top             =   3285
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   91
      Left            =   4200
      Top             =   3285
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   90
      Left            =   6765
      Top             =   3000
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   89
      Left            =   6480
      Top             =   3000
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   88
      Left            =   6195
      Top             =   3000
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   87
      Left            =   5910
      Top             =   3000
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   86
      Left            =   5625
      Top             =   3000
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   85
      Left            =   5340
      Top             =   3000
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   84
      Left            =   5055
      Top             =   3000
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   83
      Left            =   4770
      Top             =   3000
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   82
      Left            =   4485
      Top             =   3000
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   81
      Left            =   4200
      Top             =   3000
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   80
      Left            =   6765
      Top             =   2715
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   79
      Left            =   6480
      Top             =   2715
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   78
      Left            =   6195
      Top             =   2715
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   77
      Left            =   5910
      Top             =   2715
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   76
      Left            =   5625
      Top             =   2715
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   75
      Left            =   5340
      Top             =   2715
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   74
      Left            =   5055
      Top             =   2715
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   73
      Left            =   4770
      Top             =   2715
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   72
      Left            =   4485
      Top             =   2715
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   71
      Left            =   4200
      Top             =   2715
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   70
      Left            =   6765
      Top             =   2430
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   69
      Left            =   6480
      Top             =   2430
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   68
      Left            =   6195
      Top             =   2430
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   67
      Left            =   5910
      Top             =   2430
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   66
      Left            =   5625
      Top             =   2430
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   65
      Left            =   5340
      Top             =   2430
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   64
      Left            =   5055
      Top             =   2430
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   63
      Left            =   4770
      Top             =   2430
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   62
      Left            =   4485
      Top             =   2430
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   61
      Left            =   4200
      Top             =   2430
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   60
      Left            =   6765
      Top             =   2145
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   59
      Left            =   6480
      Top             =   2145
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   58
      Left            =   6195
      Top             =   2145
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   57
      Left            =   5910
      Top             =   2145
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   56
      Left            =   5625
      Top             =   2145
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   55
      Left            =   5340
      Top             =   2145
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   54
      Left            =   5055
      Top             =   2145
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   53
      Left            =   4770
      Top             =   2145
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   52
      Left            =   4485
      Top             =   2145
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   51
      Left            =   4200
      Top             =   2145
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   50
      Left            =   6765
      Top             =   1860
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   49
      Left            =   6480
      Top             =   1860
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   48
      Left            =   6195
      Top             =   1860
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   47
      Left            =   5910
      Top             =   1860
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   46
      Left            =   5625
      Top             =   1860
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   45
      Left            =   5340
      Top             =   1860
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   44
      Left            =   5055
      Top             =   1860
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   43
      Left            =   4770
      Top             =   1860
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   42
      Left            =   4485
      Top             =   1860
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   41
      Left            =   4200
      Top             =   1860
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   40
      Left            =   6765
      Top             =   1575
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   39
      Left            =   6480
      Top             =   1575
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   38
      Left            =   6195
      Top             =   1575
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   37
      Left            =   5910
      Top             =   1575
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   36
      Left            =   5625
      Top             =   1575
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   35
      Left            =   5340
      Top             =   1575
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   34
      Left            =   5055
      Top             =   1575
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   33
      Left            =   4770
      Top             =   1575
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   32
      Left            =   4485
      Top             =   1575
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   31
      Left            =   4200
      Top             =   1575
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   30
      Left            =   6765
      Top             =   1290
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   29
      Left            =   6480
      Top             =   1290
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   28
      Left            =   6195
      Top             =   1290
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   27
      Left            =   5910
      Top             =   1290
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   26
      Left            =   5625
      Top             =   1290
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   25
      Left            =   5340
      Top             =   1290
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   24
      Left            =   5055
      Top             =   1290
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   23
      Left            =   4770
      Top             =   1290
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   22
      Left            =   4485
      Top             =   1290
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   21
      Left            =   4200
      Top             =   1290
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   20
      Left            =   6765
      Top             =   1005
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   19
      Left            =   6480
      Top             =   1005
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   18
      Left            =   6195
      Top             =   1005
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   17
      Left            =   5910
      Top             =   1005
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   16
      Left            =   5625
      Top             =   1005
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   15
      Left            =   5340
      Top             =   1005
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   14
      Left            =   5055
      Top             =   1005
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   13
      Left            =   4770
      Top             =   1005
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   12
      Left            =   4485
      Top             =   1005
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   11
      Left            =   4200
      Top             =   1005
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   10
      Left            =   6765
      Top             =   720
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   9
      Left            =   6480
      Top             =   720
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   8
      Left            =   6195
      Top             =   720
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   7
      Left            =   5910
      Top             =   720
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   6
      Left            =   5625
      Top             =   720
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   5
      Left            =   5340
      Top             =   720
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   4
      Left            =   5055
      Top             =   720
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   3
      Left            =   4770
      Top             =   720
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   2
      Left            =   4485
      Top             =   720
      Width           =   270
   End
   Begin VB.Image imgFireGrid 
      Height          =   270
      Index           =   1
      Left            =   4200
      Top             =   720
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   1
      Left            =   480
      Top             =   720
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   100
      Left            =   3045
      Top             =   3285
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   99
      Left            =   2760
      Top             =   3285
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   98
      Left            =   2475
      Top             =   3285
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   97
      Left            =   2190
      Top             =   3285
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   96
      Left            =   1905
      Top             =   3285
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   95
      Left            =   1620
      Top             =   3285
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   94
      Left            =   1335
      Top             =   3285
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   93
      Left            =   1050
      Top             =   3285
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   92
      Left            =   765
      Top             =   3285
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   91
      Left            =   480
      Top             =   3285
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   90
      Left            =   3045
      Top             =   3000
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   89
      Left            =   2760
      Top             =   3000
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   88
      Left            =   2475
      Top             =   3000
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   87
      Left            =   2190
      Top             =   3000
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   86
      Left            =   1905
      Top             =   3000
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   85
      Left            =   1620
      Top             =   3000
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   84
      Left            =   1335
      Top             =   3000
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   83
      Left            =   1050
      Top             =   3000
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   82
      Left            =   765
      Top             =   3000
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   81
      Left            =   480
      Top             =   3000
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   80
      Left            =   3045
      Top             =   2715
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   79
      Left            =   2760
      Top             =   2715
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   78
      Left            =   2475
      Top             =   2715
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   77
      Left            =   2190
      Top             =   2715
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   76
      Left            =   1905
      Top             =   2715
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   75
      Left            =   1620
      Top             =   2715
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   74
      Left            =   1335
      Top             =   2715
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   73
      Left            =   1050
      Top             =   2715
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   72
      Left            =   765
      Top             =   2715
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   71
      Left            =   480
      Top             =   2715
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   70
      Left            =   3045
      Top             =   2430
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   69
      Left            =   2760
      Top             =   2430
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   68
      Left            =   2475
      Top             =   2430
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   67
      Left            =   2190
      Top             =   2430
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   66
      Left            =   1905
      Top             =   2430
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   65
      Left            =   1620
      Top             =   2430
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   64
      Left            =   1335
      Top             =   2430
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   63
      Left            =   1050
      Top             =   2430
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   62
      Left            =   765
      Top             =   2430
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   61
      Left            =   480
      Top             =   2430
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   60
      Left            =   3045
      Top             =   2145
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   59
      Left            =   2760
      Top             =   2145
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   58
      Left            =   2475
      Top             =   2145
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   57
      Left            =   2190
      Top             =   2145
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   56
      Left            =   1905
      Top             =   2145
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   55
      Left            =   1620
      Top             =   2145
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   54
      Left            =   1335
      Top             =   2145
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   53
      Left            =   1050
      Top             =   2145
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   52
      Left            =   765
      Top             =   2145
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   51
      Left            =   480
      Top             =   2145
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   50
      Left            =   3045
      Top             =   1860
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   49
      Left            =   2760
      Top             =   1860
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   48
      Left            =   2475
      Top             =   1860
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   47
      Left            =   2190
      Top             =   1860
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   46
      Left            =   1905
      Top             =   1860
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   45
      Left            =   1620
      Top             =   1860
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   44
      Left            =   1335
      Top             =   1860
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   43
      Left            =   1050
      Top             =   1860
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   42
      Left            =   765
      Top             =   1860
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   41
      Left            =   480
      Top             =   1860
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   40
      Left            =   3045
      Top             =   1575
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   39
      Left            =   2760
      Top             =   1575
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   38
      Left            =   2475
      Top             =   1575
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   37
      Left            =   2190
      Top             =   1575
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   36
      Left            =   1905
      Top             =   1575
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   35
      Left            =   1620
      Top             =   1575
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   34
      Left            =   1335
      Top             =   1575
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   33
      Left            =   1050
      Top             =   1575
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   32
      Left            =   765
      Top             =   1575
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   31
      Left            =   480
      Top             =   1575
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   30
      Left            =   3045
      Top             =   1290
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   29
      Left            =   2760
      Top             =   1290
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   28
      Left            =   2475
      Top             =   1290
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   27
      Left            =   2190
      Top             =   1290
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   26
      Left            =   1905
      Top             =   1290
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   25
      Left            =   1620
      Top             =   1290
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   24
      Left            =   1335
      Top             =   1290
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   23
      Left            =   1050
      Top             =   1290
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   22
      Left            =   765
      Top             =   1290
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   21
      Left            =   480
      Top             =   1290
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   20
      Left            =   3045
      Top             =   1005
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   19
      Left            =   2760
      Top             =   1005
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   18
      Left            =   2475
      Top             =   1005
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   17
      Left            =   2190
      Top             =   1005
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   16
      Left            =   1905
      Top             =   1005
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   15
      Left            =   1620
      Top             =   1005
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   14
      Left            =   1335
      Top             =   1005
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   13
      Left            =   1050
      Top             =   1005
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   12
      Left            =   765
      Top             =   1005
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   11
      Left            =   480
      Top             =   1005
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   10
      Left            =   3045
      Top             =   720
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   9
      Left            =   2760
      Top             =   720
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   8
      Left            =   2475
      Top             =   720
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   7
      Left            =   2190
      Top             =   720
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   6
      Left            =   1905
      Top             =   720
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   5
      Left            =   1620
      Top             =   720
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   4
      Left            =   1335
      Top             =   720
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   3
      Left            =   1050
      Top             =   720
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   2
      Left            =   765
      Top             =   720
      Width           =   270
   End
   Begin VB.Label lblAutoFire 
      BackColor       =   &H00C09440&
      Caption         =   "--"
      Height          =   255
      Left            =   1800
      TabIndex        =   23
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C09440&
      Caption         =   "Your Message :"
      Height          =   495
      Left            =   3240
      TabIndex        =   15
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C09440&
      Caption         =   "Opponent Chat :"
      Height          =   375
      Left            =   3240
      TabIndex        =   14
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C09440&
      Caption         =   "Enemy :"
      Height          =   255
      Left            =   7200
      TabIndex        =   13
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C09440&
      Caption         =   "Ships :"
      Height          =   255
      Left            =   7200
      TabIndex        =   12
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C09440&
      Caption         =   "Shot/Hit :"
      Height          =   255
      Left            =   7200
      TabIndex        =   11
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblHitShotStatus 
      BackColor       =   &H00C09440&
      Caption         =   "-"
      Height          =   255
      Left            =   8160
      TabIndex        =   10
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C09440&
      Caption         =   "Hits"
      Height          =   255
      Left            =   8880
      TabIndex        =   7
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C09440&
      Caption         =   "Shots"
      Height          =   255
      Left            =   8160
      TabIndex        =   6
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label lblNumOppShips 
      BackColor       =   &H00C09440&
      Caption         =   "0"
      Height          =   255
      Left            =   8160
      TabIndex        =   5
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblNumPlayShips 
      BackColor       =   &H00C09440&
      Caption         =   "0"
      Height          =   255
      Left            =   8160
      TabIndex        =   4
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00C09440&
      Height          =   1215
      Left            =   8160
      TabIndex        =   3
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label lblShipGridCover 
      BackColor       =   &H00400000&
      Height          =   2865
      Left            =   465
      TabIndex        =   25
      Top             =   705
      Width           =   2865
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Height          =   2865
      Left            =   4200
      TabIndex        =   28
      Top             =   720
      Width           =   2865
   End
End
Attribute VB_Name = "frmMain"
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


Private Sub cmdAutoFire_Click()
  Dim s$
  
  If Not NR Or Not CR Or Not PR Then Exit Sub
  If AutoFire Then
    s = "OFF"
  Else
    s = "ON"
  End If
  RawData = Pack("AFREQ", s, ".")
  wsNet.SendData RawData
End Sub

Private Sub cmdChangeShipGrid_Click()
  If PR Then Exit Sub
  SetMyShips
End Sub

Private Sub cmdFire_Click()
  Dim RawData$, s$
  Dim x%, y%
  
  If Not NR Then Exit Sub
  If Not PR Or Not CR Then
    'MsgBox "Both players are not ready"
    SetStatus "Both players are not ready"
    Exit Sub
  End If
  If Not MyTurn Then
    'MsgBox "Its not your turn to fire!"
    SetStatus "Its not your turn to fire!"
    Exit Sub
  End If
  If Len(txtFire) <> 2 Then
    'MsgBox "Invalid length for firing coordinates"
    SetStatus "Invalid length for firing coordinates"
    Exit Sub
  End If
  txtFire = UCase(txtFire)
  If Not CoordOk(txtFire) Then
    'MsgBox "Coordinates were invalid. Try entering like this : H3"
    SetStatus "Coordinates were invalid. Try entering like this : H3"
    Exit Sub
  End If
  x = CoordX(txtFire)
  y = CoordY(txtFire)
  If y = 0 Then y = 10
  If (x < 1 Or x > 10) Or (y < 1 Or y > 10) Then
    'MsgBox "Invalid firing coordinates"
    SetStatus "Invalid firing coordinates (" & txtFire & ")"
    Exit Sub
  End If
  
  If GetGridVal(FireGrid, x, y) <> "=" Then
    'MsgBox "You have already fired there!"
    SetStatus "You have already fired there! (" & txtFire & ")"
    Exit Sub
  End If
  
  If Host Then
    s = "H"
  Else
    s = "C"
  End If
  LastFire = Now
  'lstShots.AddItem txtFire, 0
  RawData = Pack("MYGRID", PackGrid(ShipGrid), ".")
  wsNet.SendData RawData
  RawData = Pack("FIRE", txtFire, s)
  wsNet.SendData RawData
  MyTurn = False
  'txtFire.SetFocus
  txtFire = ""
End Sub

Private Sub cmdHide_Click()
  Dim i%
  
  If PR Or Not imgShipGrid(1).Visible Then
    For i = 1 To 100
      imgShipGrid(i).Visible = Not imgShipGrid(i).Visible
    Next
  End If
End Sub

Private Sub cmdLayout_Click()
  If Not PR Then
    frmLayout.Show
  End If
End Sub

Private Sub cmdReady_Click()
  Dim RawData$
  
  If Not NR Then Exit Sub
  If CR And Not PR Then
    MyTurn = False
  Else
    MyTurn = True
  End If
    
  LastFire = Now
  lblHitShotStatus = "-"
  PR = True
  RawData = Pack("PLAYREADY", PackGrid(ShipGrid), PackGrid(ShipGrid2))
  wsNet.SendData RawData
  CountShips
  UpdateGrids
  RawData = Pack("NUMSHIPS", Trim(Str(NumPlayShips)), ".")
  wsNet.SendData RawData
End Sub


Private Sub cmdSendChat_Click()
  'On Error Resume Next
  Dim RawData$
  
  If Not NR Then Exit Sub
  If txtChat = "" Then Exit Sub
  
  RawData = Pack("CHATMSG", txtChat, ".")
  wsNet.SendData RawData
  txtChat = ""
  'txtFire.SetFocus
End Sub


Private Sub cmdSendW2K_Click()
  On Error Resume Next
  Dim i%
  
  If Not NR Then Exit Sub
  If txtChat = "" Then Exit Sub
  
  i = Shell("net send " & Computer & " " & txtChat, vbHide)
  
End Sub

Private Sub cmdSurrender_Click()
  If NR And PR And CR Then
    RawData = Pack("YOUWIN", PackGrid(ShipGrid), ".")
    wsNet.SendData RawData
    'PR = False
    'CR = False
    'MsgBox "You lose"
    PlayerWon = False
    frmOutcome.Show
    ResetGame
  End If
End Sub



Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command4_Click()
  
  BackToConnect
End Sub

Private Sub Form_Load()
  Dim c$, RawData$
  Dim i%
  Dim TimeOut As Date
  
  If FirstRun Then
    FirstRun = False
    If Host Then
      wsNet.RemotePort = 4101
      wsNet.Bind 4100
    Else
      wsNet.RemotePort = 4100
      wsNet.Bind 4101
      wsNet.RemoteHost = HostComputer
    End If
  End If
  
  MyTurn = False
  NumPlayShips = 0
  NumOppShips = 0
  ShotsFired = 0
  ShotsHit = 0
  SetMyShips
  ClearFireGrid
  UpdateGrids
  PR = False
  CR = False
  timUpdate_Timer
  
  If Not Host Then
    RawData = Pack("CONNREQ", VERID, ".")
    wsNet.SendData RawData
  End If
  
End Sub

Function SetMyShips()
  Dim x%, y%, i%, s%, hv%, xp%, yp%, j%
  Dim Ships(5) As String
  Dim sss$, c$
  
  Ships(1) = "CCCCC"    'Carrier
  Ships(2) = "BBBB"     'Battleship
  Ships(3) = "DDD"      'Destroyer
  Ships(4) = "SSS"      'Submarine
  Ships(5) = "TT"       'Scout
  
  Randomize
  For i = 1 To 100
    ShipGrid(i) = "="
    ShipGrid2(i) = ""
    j = 0
    Do While j < 1 Or j > 5
      j = (Rnd * 5) + 1
    Loop
    SeaTypeGrid(i) = j
    j = 0
    Do While j < 1 Or j > 5
      j = (Rnd * 5) + 1
    Loop
    FireTypeGrid(i) = j
  Next
  
  Randomize
  For s = 1 To 5
FindNewPos:
    'horizontal or vertical
    hv = (Rnd * 100) + 1
    If hv < 51 Then
      hv = 0
    Else
      hv = 1
    End If
    'find starting point
    If hv = 0 Then  'horizontal
RedoHz:
      xp = ((10 - Len(Ships(s))) * Rnd) + 1
      yp = (10 * Rnd) + 1
      If xp = 0 Or xp = 11 Or yp = 0 Or yp = 11 Then GoTo RedoHz
    Else            'vertical
RedoVt:
      xp = (10 * Rnd) + 1
      yp = ((10 - Len(Ships(s))) * Rnd) + 1
      If xp = 0 Or xp = 11 Or yp = 0 Or yp = 11 Then GoTo RedoVt
    End If
    'check if ship will be over another ship
    x = xp
    y = yp
    For i = 0 To Len(Ships(s)) - 1
      If hv = 0 Then
        If GetGridVal(ShipGrid, x + i, y) <> "=" Then GoTo FindNewPos
      Else
        If GetGridVal(ShipGrid, x, y + i) <> "=" Then GoTo FindNewPos
      End If
    Next
    'drawship in grid
    For i = 0 To Len(Ships(s)) - 1
      c = "M"                               'Midships
      If i = 0 Then c = "B"                 'Bow
      If i = Len(Ships(s)) - 1 Then c = "S" 'Stern
      If hv = 0 Then
        SetGridVal ShipGrid, x + i, y, Mid(Ships(s), 1, 1)
        SetGridVal ShipGrid2, x + i, y, UCase(c)
      Else
        SetGridVal ShipGrid, x, y + i, Mid(Ships(s), 1, 1)
        SetGridVal ShipGrid2, x, y + i, LCase(c)
      End If
    Next
    
  Next
    
  UpdateGrids
End Function

Function UpdateGrids()
  Dim x%, y%, i%, j%
  Dim c$, hv$, Key$
  
  For i = 1 To 100
    c = ShipGrid2(i)
    
    Key = "water"
    If c <> "" Then     'if in lcase then verti
      If c = "b" Or c = "s" Or c = "m" Then
        hv = "v"
      Else              'if in ucase then horiz
        hv = "h"
      End If
    
      c = LCase(c)
      If c = "b" Then Key = "bow" & hv
      If c = "m" Then Key = "mid" & hv
      If c = "s" Then Key = "aft" & hv
      
      If ShipGrid(i) = LCase(ShipGrid(i)) Then 'ship is hit
        Key = Key & "HIT"
      End If
    End If
    
    If ShipGrid(i) = "-" Then
      Key = "miss"
    End If
    
    Key = SeaTypeGrid(i) & "_" & Key
    
    'If ShipGrid(i) = "=" Then Key = "1_water"
    'If ShipGrid(i) = "-" Then Key = "1_miss"
    
    Set imgShipGrid(i).Picture = ilsPics.ListImages(Key).Picture
    'Set picShipGrid(i).Picture = ilsPics.ListImages(Key).Picture
    
    c = FireGrid(i)
    If c = "-" Then Key = "miss"
    If c = "X" Then Key = "hit"
    If c = "=" Or c = "" Then Key = "water"

    Key = FireTypeGrid(i) & "_" & Key
    Set imgFireGrid(i).Picture = ilsPics.ListImages(Key).Picture
  Next
  
  'lblShipGrid = ""
  'For y = 1 To 10
  '  For x = 1 To 10
  '    lblShipGrid = lblShipGrid & GetGridVal(ShipGrid, x, y)
  '  Next
  '  lblShipGrid = lblShipGrid & vbCr
  'Next
 
  'lblFireGrid = ""
  'For y = 1 To 10
  '  For x = 1 To 10
  '    lblFireGrid = lblFireGrid & GetGridVal(FireGrid, x, y)
  '  Next
  '  lblFireGrid = lblFireGrid & vbCr
  'Next
  
  lblNumPlayShips = NumPlayShips
  lblNumOppShips = NumOppShips
End Function

Function ClearFireGrid()
  Dim i%, x%, y%
  
  For i = 1 To 100
    FireGrid(i) = "="
  Next
  
  UpdateGrids
End Function

Private Sub Form_Unload(Cancel As Integer)
  If NR Then
    RawData = Pack("QUITGAME", ".", ".")
    wsNet.SendData RawData
  End If
End Sub

Private Sub imgFireGrid_Click(Index As Integer)
  Dim x%, y%
  Dim c$
  
  x = Index Mod 10
  If x = 0 Then x = 10
  y = ((Index - x) / 10) + 1
  If y = 10 Then y = 0
  c = Chr(64 + x)
  
  txtFire = c & y
  cmdFire_Click
End Sub

Private Sub timAutoFire_Timer()
  Dim x%, y%
  Dim i%, s$
  
  If Not NR Or Not PR Then
    lblAutoFire = "--"
    Exit Sub
  End If
  If Not AutoFire Then Exit Sub
  i = 0
  Randomize
  Do
    i = i + 1
    Do While (x < 1 Or x > 10)
      x = (Rnd * 10) + 1
    Loop
    Do While (y < 1 Or y > 10)
      y = (Rnd * 10) + 1
    Loop
    If GetGridVal(FireGrid, x, y) = "=" Then Exit Do
    If i > 20 Then Exit Sub
  Loop
  
  s = Chr(x + 64)
  If y = 10 Then
    s = s & "0"
  Else
    s = s & y
  End If
  lblAutoFire = s
  
  If Not MyTurn Or Not CR Then Exit Sub
  
  i = DateDiff("s", LastFire, Now)
  s = AUTOFIRETIME - i
  lblAutoFire = lblAutoFire & " (" & s & ")"
  If i >= AUTOFIRETIME Then
    txtFire = lblAutoFire
    cmdFire_Click
  End If
End Sub

Private Sub timUpdate_Timer()
  Dim s$
  
  If NR Then
    s = "NetRdy"
  Else
    s = "NetNotRdy"
  End If
  s = s & vbCr
  If PR Then
    s = s & "PlayerRdy"
  Else
    s = s & "PlayerNotRdy"
  End If
  s = s & vbCr
  If CR Then
    s = s & "OpponentRdy"
  Else
    s = s & "OpponentNotRdy"
  End If
  s = s & vbCr
  If MyTurn Then
    s = s & "YourTurn"
  Else
    s = s & "NotYourTurn"
  End If
  
  If MyTurn And NR And PR And CR Then
    shpTurn.FillColor = &HFF00&
  Else
    shpTurn.FillColor = &HFF&
  End If
  
  
  lblStatus = s
  
End Sub

Private Sub wsNet_DataArrival(ByVal bytesTotal As Long)
  On Error Resume Next
  Dim RawData$, Command$, Stamp$, Extra$, PName$, Extra2$
  Dim c$
  
  wsNet.GetData RawData
  If Err.Number = 10054 Then
    Err.Clear
    MsgBox "Connection Failure", vbOKOnly + vbCritical
    End
  End If
  Unpack RawData, Computer, Stamp, Command, PName, Extra, Extra2
  'wsNet.RemoteHost = Computer
  If Not NR Then
    wsNet.RemoteHost = Computer
    Opponent = PName
  End If
  
  Select Case UCase(Command)
    Case "CONNREQ"
      If Host Then
        c = "CONNREQ_OK"
        If NR Then c = "CONNREQ_NO"
        If Extra <> VERID Then c = c = "CONNREQ_NO"
      Else
        c = "CONNREQ_NO"
      End If
      RawData = Pack(c, VERID, ".")
      wsNet.SendData RawData
    Case "CONNREQ_OK"
      If Not Host Then
        NR = True
        RawData = Pack("NETREADY", ".", ".")
        wsNet.SendData RawData
        'MsgBox "Client Ready"
        Me.Caption = App.Title & " - " & Player & " vs " & Opponent & " [CLIENT]"
        SetStatus "Connected to " & Computer
      End If
    Case "CONNREQ_NO"
      If Not Host Then
        Me.Hide
        If Extra <> VERID Then
          c = "Host is running a different version"
        Else
          c = "Connection with the host was denied"
        End If
        MsgBox c
        End
      End If
    Case "NETREADY"
      If Host Then
        NR = True
        'MsgBox "Host Ready"
        Me.Caption = App.Title & " - " & Player & " vs " & Opponent & " [HOST]"
        SetStatus "Connected to " & Computer
      End If
    Case "PLAYREADY"
      CR = True
      LastFire = Now
      UnpackGrid Extra, OppGrid
      UnpackGrid Extra2, OppGrid2
    Case "FIRE"
      If (Host And Extra2 = "C") Or (Not Host And Extra2 = "H") Then
        c = FireOnGrid(Extra)
        LastFire = Now
        RawData = Pack("FIREINFO", c, Extra)
        wsNet.SendData RawData
        If c <> "=" And c <> "-" Then
          lblHitShotStatus = "You been hit (" & Extra & ")"
          lstHits.AddItem Extra & " +", 0
        Else
          lblHitShotStatus = "You were missed (" & Extra & ")"
          lstHits.AddItem Extra, 0
        End If
        MyTurn = True
        CountShips
        UpdateGrids
        RawData = Pack("NUMSHIPS", Trim(Str(NumPlayShips)), ".")
        wsNet.SendData RawData
        If NumPlayShips = 0 Then
          RawData = Pack("YOUWIN", PackGrid(ShipGrid), ".")
          wsNet.SendData RawData
          'PR = False
          'CR = False
          'MsgBox "You lose"
          PlayerWon = False
          frmOutcome.Show
          ResetGame
        End If
      End If
    Case "FIREINFO"
      ShotsFired = ShotsFired + 1
      If Extra = "=" Then
        c = "-"
        lblHitShotStatus = Extra2 & " - Miss"
        lstShots.AddItem Extra2, 0
      Else
        c = Extra
        lblHitShotStatus = Extra2 & " - Hit"
        lstShots.AddItem Extra2 & " +", 0
        ShotsHit = ShotsHit + 1
      End If
      lblShotHitRatio = ShotsFired & "/" & ShotsHit & " (" & Round((ShotsHit / ShotsFired) * 100) & "%)"
      
      SetGridVal FireGrid, CoordX(Extra2), CoordY(Extra2), c
      UpdateGrids
    Case "NUMSHIPS"
      NumOppShips = Val(Extra)
      UpdateGrids
    Case "YOUWIN"
      'PR = False
      'CR = False
      'RawData = Pack("RESTART", ".", ".")
      'wsNet.SendData RawData
      'MsgBox "You win"
      PlayerWon = True
      UnpackGrid Extra, OppGrid
      frmOutcome.Show
      ResetGame
    'Case "RESTART"
    '  ResetGame
    Case "CHATMSG"
      txtOppChat = Extra & Chr(13) & Chr(10) & txtOppChat
      SetStatus Opponent & " has messaged you"
    Case "MYGRID"
      UnpackGrid Extra, OppGrid
    Case "AFREQ"
      i = MsgBox(Opponent & " has requested Auto-Fire be switched " & Extra & vbCr & _
         vbCr & "Do you comply?", vbYesNo)
      If i = vbYes Then
        If Extra = "ON" Then
          AutoFire = True
        Else
          AutoFire = False
          lblAutoFire = "--"
        End If
        RawData = Pack("AFREQ_REP", Extra, "OK")
        wsNet.SendData RawData
        LastFire = Now
      End If
      If i = vbNo Then
        RawData = Pack("AFREQ_REP", Extra, "NO")
        wsNet.SendData RawData
      End If
    Case "AFREQ_REP"
      If Extra2 = "OK" Then
        LastFire = Now
        If Extra = "ON" Then
          AutoFire = True
          MsgBox "Auto-Fire ON"
        Else
          AutoFire = False
          lblAutoFire = "--"
          MsgBox "Auto-Fire OFF"
        End If
      End If
      If Extra2 = "NO" Then
        MsgBox "Your request to switch Auto-Fire " & Extra & " has been denied!"
      End If
    Case "QUITGAME"
      MsgBox Opponent & " has quit the game!", vbOKOnly
      If Host Then
        ResetGame
      Else
        If frmOutcome.Visible Then Unload frmOutcome
        If frmLayout.Visible Then Unload frmLayout
        BackToConnect
      End If
  End Select
End Sub

Function ResetGame()
  NR = False
  Form_Load
  lblHitShotStatus = "-"
  lstHits.Clear
  lstShots.Clear
  lblAutoFire = "--"
  Me.Caption = App.Title
  SetStatus ""
End Function

Function FireOnGrid(Coord$) As String
  Dim x%, y%
  Dim c$
  
  x = CoordX(Coord)
  y = CoordY(Coord)
  FireOnGrid = GetGridVal(ShipGrid, x, y)
  If FireOnGrid <> "=" Then
    c = LCase(FireOnGrid)
    FireOnGrid = "X"
  Else
    c = "-"
  End If
  SetGridVal ShipGrid, x, y, c
End Function

Function CountShips()
  Dim i%, c%, b%, d%, s%, t%
  
  c = 0
  b = 0
  d = 0
  s = 0
  t = 0
  For i = 1 To 100
    If ShipGrid(i) = "C" Then c = 1
    If ShipGrid(i) = "B" Then b = 1
    If ShipGrid(i) = "D" Then d = 1
    If ShipGrid(i) = "S" Then s = 1
    If ShipGrid(i) = "T" Then t = 1
  Next
  NumPlayShips = c + b + d + s + t
End Function

Function BackToConnect()
  If NR Then
    RawData = Pack("QUITGAME", PackGrid(ShipGrid), ".")
    wsNet.SendData RawData
  End If
  NR = False
  PR = False
  CR = False
  MyTurn = False
  frmMain.Visible = False
  Unload frmMain
  frmConnect.Show
  Unload Me
End Function

Function SetStatus(status$)
  lblStatus2 = status
End Function
