VERSION 5.00
Begin VB.Form frmLayout 
   BackColor       =   &H00C09440&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ShipGrid Layout"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   Icon            =   "frmLayout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   341
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   404
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSet 
      Caption         =   "Set"
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdShip 
      Caption         =   "Cruiser (2)"
      Height          =   375
      Index           =   5
      Left            =   4080
      TabIndex        =   5
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton cmdShip 
      Caption         =   "Submarine (3)"
      Height          =   375
      Index           =   4
      Left            =   4080
      TabIndex        =   4
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton cmdShip 
      Caption         =   "Destroyer (3)"
      Height          =   375
      Index           =   3
      Left            =   4080
      TabIndex        =   3
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdShip 
      Caption         =   "Battleship (4)"
      Height          =   375
      Index           =   2
      Left            =   4080
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmdShip 
      Caption         =   "Carrier (5)"
      Height          =   375
      Index           =   1
      Left            =   4080
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lblShip 
      BackColor       =   &H00C09440&
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C09440&
      Caption         =   "Left click places ship horiztonally, Right click places ship vertically"
      Height          =   615
      Left            =   3720
      TabIndex        =   7
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   2
      Left            =   810
      Top             =   525
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   3
      Left            =   1095
      Top             =   525
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   4
      Left            =   1380
      Top             =   525
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   5
      Left            =   1665
      Top             =   525
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   6
      Left            =   1950
      Top             =   525
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   7
      Left            =   2235
      Top             =   525
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   8
      Left            =   2520
      Top             =   525
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   9
      Left            =   2805
      Top             =   525
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   10
      Left            =   3090
      Top             =   525
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   11
      Left            =   525
      Top             =   810
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   12
      Left            =   810
      Top             =   810
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   13
      Left            =   1095
      Top             =   810
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   14
      Left            =   1380
      Top             =   810
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   15
      Left            =   1665
      Top             =   810
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   16
      Left            =   1950
      Top             =   810
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   17
      Left            =   2235
      Top             =   810
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   18
      Left            =   2520
      Top             =   810
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   19
      Left            =   2805
      Top             =   810
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   20
      Left            =   3090
      Top             =   810
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   21
      Left            =   525
      Top             =   1095
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   22
      Left            =   810
      Top             =   1095
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   23
      Left            =   1095
      Top             =   1095
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   24
      Left            =   1380
      Top             =   1095
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   25
      Left            =   1665
      Top             =   1095
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   26
      Left            =   1950
      Top             =   1095
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   27
      Left            =   2235
      Top             =   1095
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   28
      Left            =   2520
      Top             =   1095
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   29
      Left            =   2805
      Top             =   1095
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   30
      Left            =   3090
      Top             =   1095
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   31
      Left            =   525
      Top             =   1380
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   32
      Left            =   810
      Top             =   1380
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   33
      Left            =   1095
      Top             =   1380
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   34
      Left            =   1380
      Top             =   1380
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   35
      Left            =   1665
      Top             =   1380
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   36
      Left            =   1950
      Top             =   1380
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   37
      Left            =   2235
      Top             =   1380
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   38
      Left            =   2520
      Top             =   1380
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   39
      Left            =   2805
      Top             =   1380
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   40
      Left            =   3090
      Top             =   1380
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   41
      Left            =   525
      Top             =   1665
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   42
      Left            =   810
      Top             =   1665
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   43
      Left            =   1095
      Top             =   1665
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   44
      Left            =   1380
      Top             =   1665
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   45
      Left            =   1665
      Top             =   1665
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   46
      Left            =   1950
      Top             =   1665
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   47
      Left            =   2235
      Top             =   1665
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   48
      Left            =   2520
      Top             =   1665
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   49
      Left            =   2805
      Top             =   1665
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   50
      Left            =   3090
      Top             =   1665
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   51
      Left            =   525
      Top             =   1950
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   52
      Left            =   810
      Top             =   1950
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   53
      Left            =   1095
      Top             =   1950
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   54
      Left            =   1380
      Top             =   1950
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   55
      Left            =   1665
      Top             =   1950
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   56
      Left            =   1950
      Top             =   1950
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   57
      Left            =   2235
      Top             =   1950
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   58
      Left            =   2520
      Top             =   1950
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   59
      Left            =   2805
      Top             =   1950
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   60
      Left            =   3090
      Top             =   1950
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   61
      Left            =   525
      Top             =   2235
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   62
      Left            =   810
      Top             =   2235
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   63
      Left            =   1095
      Top             =   2235
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   64
      Left            =   1380
      Top             =   2235
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   65
      Left            =   1665
      Top             =   2235
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   66
      Left            =   1950
      Top             =   2235
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   67
      Left            =   2235
      Top             =   2235
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   68
      Left            =   2520
      Top             =   2235
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   69
      Left            =   2805
      Top             =   2235
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   70
      Left            =   3090
      Top             =   2235
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   71
      Left            =   525
      Top             =   2520
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   72
      Left            =   810
      Top             =   2520
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   73
      Left            =   1095
      Top             =   2520
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   74
      Left            =   1380
      Top             =   2520
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   75
      Left            =   1665
      Top             =   2520
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   76
      Left            =   1950
      Top             =   2520
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   77
      Left            =   2235
      Top             =   2520
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   78
      Left            =   2520
      Top             =   2520
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   79
      Left            =   2805
      Top             =   2520
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   80
      Left            =   3090
      Top             =   2520
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   81
      Left            =   525
      Top             =   2805
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   82
      Left            =   810
      Top             =   2805
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   83
      Left            =   1095
      Top             =   2805
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   84
      Left            =   1380
      Top             =   2805
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   85
      Left            =   1665
      Top             =   2805
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   86
      Left            =   1950
      Top             =   2805
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   87
      Left            =   2235
      Top             =   2805
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   88
      Left            =   2520
      Top             =   2805
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   89
      Left            =   2805
      Top             =   2805
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   90
      Left            =   3090
      Top             =   2805
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   91
      Left            =   525
      Top             =   3090
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   92
      Left            =   810
      Top             =   3090
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   93
      Left            =   1095
      Top             =   3090
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   94
      Left            =   1380
      Top             =   3090
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   95
      Left            =   1665
      Top             =   3090
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   96
      Left            =   1950
      Top             =   3090
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   97
      Left            =   2235
      Top             =   3090
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   98
      Left            =   2520
      Top             =   3090
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   99
      Left            =   2805
      Top             =   3090
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   100
      Left            =   3090
      Top             =   3090
      Width           =   270
   End
   Begin VB.Image imgShipGrid 
      Height          =   270
      Index           =   1
      Left            =   525
      Top             =   525
      Width           =   270
   End
   Begin VB.Image Image4 
      Height          =   270
      Left            =   525
      Picture         =   "frmLayout.frx":08CA
      Top             =   240
      Width           =   2835
   End
   Begin VB.Image Image1 
      Height          =   2835
      Left            =   240
      Picture         =   "frmLayout.frx":30FC
      Top             =   525
      Width           =   270
   End
   Begin VB.Label lblShipGridCover 
      BackColor       =   &H00000000&
      Height          =   2865
      Left            =   510
      TabIndex        =   9
      Top             =   510
      Width           =   2865
   End
End
Attribute VB_Name = "frmLayout"
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


Dim LayGrid(100) As String
Dim LayGrid2(100) As String

Dim CurShip$

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdSet_Click()
  Dim i%, c%
  
  c = 0
  For i = 1 To 100
    If LayGrid(i) <> "=" Then c = c + 1
  Next
  If c <> 17 Then
    MsgBox "All ships have not been placed"
  Else
    For i = 1 To 100
      ShipGrid(i) = LayGrid(i)
      ShipGrid2(i) = LayGrid2(i)
    Next
    frmMain.UpdateGrids
    cmdClose_Click
  End If
End Sub

Private Sub cmdShip_Click(Index As Integer)
  Dim Flag As Boolean
  Dim i%
  
  Select Case Index
    Case 1
      CurShip = "CCCCC"
      lblShip = "Carrier"
    Case 2
      CurShip = "BBBB"
      lblShip = "Battleship"
    Case 3
      CurShip = "DDD"
      lblShip = "Destroyer"
    Case 4
      CurShip = "SSS"
      lblShip = "Submarine"
    Case 5
      CurShip = "TT"
      lblShip = "Cruiser"
  End Select
  
  Flag = False
  For i = 1 To 100
    If UCase(LayGrid(i)) = Mid(CurShip, 1, 1) Then
      LayGrid(i) = "="
      LayGrid2(i) = ""
      Flag = True
    End If
  Next
  
  If Flag Then lblShip = ""
   
  DrawGrid
End Sub

Private Sub Form_Load()
  Dim i%
  
  frmMain.Visible = False
  frmMain.Enabled = False
  
  For i = 1 To 100
    LayGrid(i) = "=" 'ShipGrid(i)
    LayGrid2(i) = "" 'ShipGrid(2)
  Next
  
  DrawGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmMain.Enabled = True
  frmMain.Visible = True
End Sub

Function DrawGrid()
  Dim x%, y%, i%, j%
  Dim c$, hv$, Key$
  
  For i = 1 To 100
    c = LayGrid2(i)
    
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
      
    End If

    
    Key = 2 & "_" & Key
    
    Set imgShipGrid(i).Picture = frmMain.ilsPics.ListImages(Key).Picture
    'Set picShipGrid(i).Picture = ilsPics.ListImages(Key).Picture
    
  Next
  
  
End Function

Private Sub imgShipGrid_MouseDown(Index As Integer, Button As Integer, Shift As Integer, MX As Single, MY As Single)
  Dim hv$, c$
  Dim i%, x%, y%, xa%, ya%
  
  Flag = False
  If CurShip = "" Then Exit Sub
  For i = 1 To 100
    If UCase(LayGrid(i)) = Mid(CurShip, 1, 1) Then
      MsgBox "That ship has been placed already", vbOKOnly
      Exit Sub
    End If
  Next
  
  If Flag Then GoTo Finish
  
  If Button = 1 Then
    hv = "h"
    xa = 1
    ya = 0
  Else
    hv = "v"
    xa = 0
    ya = 1
  End If
  
  x = Index Mod 10
  If x = 0 Then x = 10
  y = ((Index - x) / 10) + 1
  'If y = 10 Then y = 0
  
  
  If (hv = "h" And (x + Len(CurShip) > 11)) Or (hv = "v" And (y + Len(CurShip) > 11)) Then
    MsgBox "Your ship is over the border"
    Exit Sub
  End If
  
  For i = 1 To Len(CurShip)
    If GetGridVal(LayGrid, x + ((i - 1) * xa), y + ((i - 1) * ya)) <> "=" Then
      MsgBox "Ship is over an existing ship"
      Exit Sub
    End If
  Next
  
  For i = 1 To Len(CurShip)
    SetGridVal LayGrid, x + ((i - 1) * xa), y + ((i - 1) * ya), Mid(CurShip, i, 1)
    c = "m"
    If i = 1 Then c = "b"
    If i = Len(CurShip) Then c = "s"
    If hv = "h" Then c = UCase(c)
    SetGridVal LayGrid2, x + ((i - 1) * xa), y + ((i - 1) * ya), c
  Next
  
Finish:
  DrawGrid
  CurShip = ""
End Sub


