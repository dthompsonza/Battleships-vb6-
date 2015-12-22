Attribute VB_Name = "Globals"
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

Global Host As Boolean
Global HostComputer As String

Global Player As String
Global Opponent As String
Global Computer As String

Global FirstRun As Boolean

Global NR As Boolean 'Network conn estab
Global PR As Boolean 'Player-Ready
Global CR As Boolean 'Opponent-Ready
Global MyTurn As Boolean

Global NumPlayShips As Integer
Global NumOppShips As Integer

Global ShipGrid(100) As String
Global FireGrid(100) As String
Global OppGrid(100) As String

Global SeaTypeGrid(100) As Integer 'What color type sea is 1-5
Global FireTypeGrid(100) As Integer 'What color type sea is 1-5
Global ShipGrid2(100) As String 'Ship positions
Global OppGrid2(100) As String  'Ship positions


Global LastFire As Date
Global Const AUTOFIRETIME% = 45
Global AutoFire As Boolean
Global PlayerWon As Boolean

Global Broadcast As Boolean
Global BroadcastTo As String

Global ShotsFired As Integer
Global ShotsHit As Integer

Global Const VERID$ = "BattleShips1.18"

Function GetVersion() As String
  GetVersion = App.Major & "." & App.Minor & "." & App.Revision
End Function

Function Pack(Command$, Extra$, Extra2$) As String
  Pack = GetTheComputerName & vbCr & _
    Format(Now, "dd mmm yyyy hh:nn:ss") & vbCr & _
     Player & vbCr & Command & vbCr & Extra & vbCr & Extra2
End Function

Function Unpack(RawData$, Computer$, Stamp$, Command$, PlayerName$, Extra$, Extra2$)
  Dim Data() As String
  
  Data = Split(RawData, vbCr)
  If UBound(Data) <> 5 Then Exit Function
  Computer = Data(0)
  Stamp = Data(1)
  PlayerName = Data(2)
  Command = Data(3)
  Extra = Data(4)
  Extra2 = Data(5)
End Function

Function PackGrid(Grid() As String) As String
  Dim i%
  Dim s$
  
  For i = 1 To 100
    If i > 1 Then s = s & "/"
    s = s & Grid(i)
  Next
  PackGrid = s
End Function

Function UnpackGrid(Packed As String, Grid() As String)
  Dim Data() As String
  Dim i%
  Data = Split(Packed, "/")
  For i = LBound(Data) To UBound(Data)
    'If i > 0 And i < 101 Then Grid(i) = Data(i)
    Grid(i + 1) = Data(i)
  Next
End Function

Function GetGridVal(PGrid() As String, x%, y%) As String
  GetGridVal = PGrid(((y - 1) * 10) + x)
End Function

Function SetGridVal(PGrid() As String, x%, y%, Valu$)
  PGrid(((y - 1) * 10) + x) = Valu
End Function

Function CoordOk(Coord$) As Boolean
  Dim NumFound As Boolean
  Dim AlphFound As Boolean
  Dim i%
  
  CoordOk = False
  If Len(CoordOk) <> 2 Then Exit Function
  AlphFound = False
  NumFound = False
  For i = Asc("A") To Asc("J")
    If UCase(Mid(Coord, 1, 1)) = Chr(i) Then AlphFound = True
  Next
  For i = Asc("0") To Asc("9")
    If Mid(Coord, 2, 1) = Chr(i) Then NumFound = True
  Next
  If AlphFound And NumFound Then CoordOk = True
End Function

Function CoordX(Coord$) As Integer
  CoordX = Asc(Mid(UCase(Coord), 1, 1)) - 64
End Function

Function CoordY(Coord$) As Integer
  CoordY = Mid(Coord, 2, 1)
  If CoordY = 0 Then CoordY = 10
End Function

