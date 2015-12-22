VERSION 5.00
Begin VB.Form frmOutcome 
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7140
   Icon            =   "frmOutcome.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   7140
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   $"frmOutcome.frx":08CA
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   360
      TabIndex        =   8
      Top             =   3840
      Width           =   4455
   End
   Begin VB.Label lblShipGridCover 
      BackColor       =   &H00000000&
      Height          =   2865
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   2865
   End
   Begin VB.Image Image1 
      Height          =   2835
      Left            =   0
      Picture         =   "frmOutcome.frx":095C
      Top             =   285
      Width           =   270
   End
   Begin VB.Image Image4 
      Height          =   270
      Left            =   285
      Picture         =   "frmOutcome.frx":32F6
      Top             =   0
      Width           =   2835
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   1
      Left            =   285
      Top             =   285
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   100
      Left            =   2850
      Top             =   2850
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   99
      Left            =   2565
      Top             =   2850
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   98
      Left            =   2280
      Top             =   2850
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   97
      Left            =   1995
      Top             =   2850
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   96
      Left            =   1710
      Top             =   2850
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   95
      Left            =   1425
      Top             =   2850
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   94
      Left            =   1140
      Top             =   2850
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   93
      Left            =   855
      Top             =   2850
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   92
      Left            =   570
      Top             =   2850
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   91
      Left            =   285
      Top             =   2850
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   90
      Left            =   2850
      Top             =   2565
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   89
      Left            =   2565
      Top             =   2565
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   88
      Left            =   2280
      Top             =   2565
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   87
      Left            =   1995
      Top             =   2565
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   86
      Left            =   1710
      Top             =   2565
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   85
      Left            =   1425
      Top             =   2565
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   84
      Left            =   1140
      Top             =   2565
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   83
      Left            =   855
      Top             =   2565
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   82
      Left            =   570
      Top             =   2565
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   81
      Left            =   285
      Top             =   2565
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   80
      Left            =   2850
      Top             =   2280
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   79
      Left            =   2565
      Top             =   2280
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   78
      Left            =   2280
      Top             =   2280
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   77
      Left            =   1995
      Top             =   2280
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   76
      Left            =   1710
      Top             =   2280
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   75
      Left            =   1425
      Top             =   2280
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   74
      Left            =   1140
      Top             =   2280
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   73
      Left            =   855
      Top             =   2280
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   72
      Left            =   570
      Top             =   2280
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   71
      Left            =   285
      Top             =   2280
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   70
      Left            =   2850
      Top             =   1995
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   69
      Left            =   2565
      Top             =   1995
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   68
      Left            =   2280
      Top             =   1995
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   67
      Left            =   1995
      Top             =   1995
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   66
      Left            =   1710
      Top             =   1995
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   65
      Left            =   1425
      Top             =   1995
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   64
      Left            =   1140
      Top             =   1995
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   63
      Left            =   855
      Top             =   1995
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   62
      Left            =   570
      Top             =   1995
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   61
      Left            =   285
      Top             =   1995
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   60
      Left            =   2850
      Top             =   1710
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   59
      Left            =   2565
      Top             =   1710
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   58
      Left            =   2280
      Top             =   1710
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   57
      Left            =   1995
      Top             =   1710
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   56
      Left            =   1710
      Top             =   1710
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   55
      Left            =   1425
      Top             =   1710
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   54
      Left            =   1140
      Top             =   1710
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   53
      Left            =   855
      Top             =   1710
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   52
      Left            =   570
      Top             =   1710
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   51
      Left            =   285
      Top             =   1710
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   50
      Left            =   2850
      Top             =   1425
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   49
      Left            =   2565
      Top             =   1425
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   48
      Left            =   2280
      Top             =   1425
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   47
      Left            =   1995
      Top             =   1425
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   46
      Left            =   1710
      Top             =   1425
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   45
      Left            =   1425
      Top             =   1425
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   44
      Left            =   1140
      Top             =   1425
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   43
      Left            =   855
      Top             =   1425
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   42
      Left            =   570
      Top             =   1425
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   41
      Left            =   285
      Top             =   1425
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   40
      Left            =   2850
      Top             =   1140
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   39
      Left            =   2565
      Top             =   1140
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   38
      Left            =   2280
      Top             =   1140
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   37
      Left            =   1995
      Top             =   1140
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   36
      Left            =   1710
      Top             =   1140
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   35
      Left            =   1425
      Top             =   1140
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   34
      Left            =   1140
      Top             =   1140
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   33
      Left            =   855
      Top             =   1140
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   32
      Left            =   570
      Top             =   1140
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   31
      Left            =   285
      Top             =   1140
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   30
      Left            =   2850
      Top             =   855
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   29
      Left            =   2565
      Top             =   855
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   28
      Left            =   2280
      Top             =   855
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   27
      Left            =   1995
      Top             =   855
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   26
      Left            =   1710
      Top             =   855
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   25
      Left            =   1425
      Top             =   855
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   24
      Left            =   1140
      Top             =   855
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   23
      Left            =   855
      Top             =   855
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   22
      Left            =   570
      Top             =   855
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   21
      Left            =   285
      Top             =   855
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   20
      Left            =   2850
      Top             =   570
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   19
      Left            =   2565
      Top             =   570
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   18
      Left            =   2280
      Top             =   570
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   17
      Left            =   1995
      Top             =   570
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   16
      Left            =   1710
      Top             =   570
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   15
      Left            =   1425
      Top             =   570
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   14
      Left            =   1140
      Top             =   570
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   13
      Left            =   855
      Top             =   570
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   12
      Left            =   570
      Top             =   570
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   11
      Left            =   285
      Top             =   570
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   10
      Left            =   2850
      Top             =   285
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   9
      Left            =   2565
      Top             =   285
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   8
      Left            =   2280
      Top             =   285
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   7
      Left            =   1995
      Top             =   285
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   6
      Left            =   1710
      Top             =   285
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   5
      Left            =   1425
      Top             =   285
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   4
      Left            =   1140
      Top             =   285
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   3
      Left            =   855
      Top             =   285
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   2
      Left            =   570
      Top             =   285
      Width           =   270
   End
   Begin VB.Label lblShip 
      BackColor       =   &H00C09440&
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      Caption         =   "-"
      Height          =   255
      Left            =   5640
      TabIndex        =   5
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Your Grid"
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblPlayGrid 
      Caption         =   "=========="
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2415
      Left            =   3960
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Opponent Grid"
      Height          =   255
      Left            =   5520
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblOppGrid 
      Caption         =   "=========="
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2415
      Left            =   5640
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmOutcome"
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


Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Dim x%, y%, ps%, os%, i%
  Dim s$, c$
  
  frmMain.Visible = False
  If PlayerWon Then
    Me.Caption = "You Win - " & Player & " beat " & Opponent
  Else
    Me.Caption = "You Lose - " & Opponent & " beat " & Player
  End If
  
  ps = 0
  os = 0
  For i = 1 To 100
    c = ShipGrid(i)
    If c = UCase(c) And c <> "=" And c <> "-" Then ps = ps + 1
    c = OppGrid(i)
    If c = UCase(c) And c <> "=" And c <> "-" Then os = os + 1
  Next
  
  lblScore = ps & " - " & os
  If PlayerWon Then
    s = "Battleships" & vbCr & vbCr & Player & " on " & GetTheComputerName & "  beat  " & Opponent & " on " & Computer & "   (" & lblScore & ")"
  End If
  
  s = ""
  For y = 1 To 10
    For x = 1 To 10
      s = s & GetGridVal(OppGrid, x, y)
    Next
    s = s & vbCr
  Next
  lblOppGrid = s
  s = ""
  For y = 1 To 10
    For x = 1 To 10
      s = s & GetGridVal(ShipGrid, x, y)
    Next
    s = s & vbCr
  Next
  lblPlayGrid = s
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmMain.Visible = True
End Sub
