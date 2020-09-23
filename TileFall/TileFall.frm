VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5490
   ClientLeft      =   150
   ClientTop       =   705
   ClientWidth     =   9030
   Icon            =   "TileFall.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   366
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   602
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   5430
      Left            =   0
      ScaleHeight     =   358
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   598
      TabIndex        =   0
      Top             =   30
      Width           =   9030
   End
   Begin VB.Menu NewGame 
      Caption         =   "&New Game"
      Index           =   1
   End
   Begin VB.Menu Size 
      Caption         =   "&Size"
      Index           =   2
      Begin VB.Menu Size_Small 
         Caption         =   "Small"
      End
      Begin VB.Menu Size_Medium 
         Caption         =   "Medium"
      End
      Begin VB.Menu Size_Large 
         Caption         =   "Large"
      End
      Begin VB.Menu Size_Huge 
         Caption         =   "Huge"
      End
   End
   Begin VB.Menu Colours 
      Caption         =   "&Colours"
      Begin VB.Menu Colours_3 
         Caption         =   "3 Colours"
      End
      Begin VB.Menu Colours_4 
         Caption         =   "4 Colours"
      End
      Begin VB.Menu Colours_5 
         Caption         =   "5 Colours"
      End
      Begin VB.Menu Colours_6 
         Caption         =   "6 Colours"
      End
   End
   Begin VB.Menu TopTen 
      Caption         =   "&Hall of Fame"
      Index           =   4
   End
   Begin VB.Menu z 
      Caption         =   "                                 TileFall - ©2001 RVT                            "
      Index           =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Tilefall - ©2001 Ron van Tilburg - My version of a game I knew 10 years or so ago from my Amiga Days
'Freely usable with appropriate credits to author

'The basic data structure representing the tiles
Private Type SQ
  c As Byte   'colour>1 is visible, =0 is gone
  s As Byte   'Selected
End Type

'what is kept in the Hall of Fame

Private Type ScoreEntry
  Name As String * 16
  Score As Long
  Date As Date
End Type

Private Type TopTen
  Scores(0 To 9) As ScoreEntry
End Type

'some useful constants
Const SQSIZE As Long = 60
Const GAME_SMALL As Long = 0
Const GAME_MEDIUM As Long = 1
Const GAME_LARGE  As Long = 2
Const GAME_HUGE  As Long = 3

'Globals
Dim Form2 As Form2
Dim Form3 As Form3

Dim Game() As SQ
Dim CurrGame As Long
Dim NColours As Long
Dim MaxCols  As Long
Dim MaxRows  As Long
Dim GameSize As Long
Dim SqrSize  As Long
Dim NLeft As Long

Dim Username As String * 16
Dim Score As Long
Dim HiScore As Long
Dim HOFHIScore As Long
Dim HallOfFame(0 To 15) As TopTen

Dim Inks(0 To 9) As Long

Private Sub Form_Load()
  NColours = 4
  MaxCols = 9
  MaxRows = 5
  SqrSize = 60
  GameSize = GAME_SMALL
  Inks(0) = &HCEDEDD    'BG
  Inks(1) = &HF8F8F8    'WHITE
  Inks(2) = &HA0A0A0    'Grey
  Inks(3) = &H50E0F0
  Inks(4) = &HF08050
  Inks(5) = &H7060E0
  Inks(6) = &H50D050
  Inks(7) = &H50A0F0
  Inks(8) = &HE050E0
  On Error GoTo HOFErr
  Open App.Path & "\HallOfFame.dat" For Binary Access Read As #1
  Get #1, , HallOfFame()
  Close #1
HOFErr:
  On Error GoTo 0
  Set Form2 = New Form2
  Set Form3 = New Form3
  Username = "?"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Open App.Path & "\HallOfFame.dat" For Binary Access Write As #1
  Put #1, , HallOfFame()
  Close #1
  Unload Form2
End Sub

'A brand new game

Private Sub NewGame_Click(Index As Integer)
  Dim i As Long, j As Long
  Dim x0 As Long, y0 As Long, x1 As Long, y1 As Long
  
  If Left(Username, 1) = "?" Then       'if we havent asked for a username do so now (ONCE ONLY)
    Form3.Show vbModal, Form1
    Username = Form3.Username
    Unload Form3
    Set Form3 = Nothing
  End If
  
  'If we have changed games (Size or Nr of Colours) then adjust parameters and scores
  If (GameSize * 4 + NColours - 3) <> CurrGame Then
    CurrGame = GameSize * 4 + NColours - 3
    HiScore = 0
  End If
  HOFHIScore = HallOfFame(CurrGame).Scores(0).Score
  
  'Set up the Game
  Randomize Timer
  ReDim Game(0 To MaxCols, 0 To MaxRows) As SQ
  For i = 0 To MaxCols
    For j = 0 To MaxRows
      Game(i, j).c = 3 + Int(NColours * Rnd)
      Game(i, j).s = 0
    Next
  Next
  Score = 0: NLeft = (MaxCols + 1) * (MaxRows + 1)
  Call DrawTableau
End Sub

'This routine finds all connected tiles of a given colour
'NOTE it is a fill routine in disguise
'NOTE NOTE NOTE - this is RECURSIVELY CALLED

Private Sub CheckEdges(ByVal i As Long, ByVal j As Long)
  If Game(i, j).c <> 0 Then
    If j > 0 Then
      If Game(i, j - 1).s <> Game(i, j).s Then
        If Game(i, j - 1).c = Game(i, j).c Then
          Game(i, j - 1).s = Game(i, j).s
          Call CheckEdges(i, j - 1)
        End If
      End If
    End If
    
    If j < MaxRows Then
      If Game(i, j + 1).s <> Game(i, j).s Then
        If Game(i, j + 1).c = Game(i, j).c Then
          Game(i, j + 1).s = Game(i, j).s
          Call CheckEdges(i, j + 1)
        End If
      End If
    End If
      
    If i > 0 Then
      If Game(i - 1, j).s <> Game(i, j).s Then
        If Game(i - 1, j).c = Game(i, j).c Then
          Game(i - 1, j).s = Game(i, j).s
          Call CheckEdges(i - 1, j)
        End If
      End If
    End If
      
    If i < MaxCols Then
      If Game(i + 1, j).s <> Game(i, j).s Then
        If Game(i + 1, j).c = Game(i, j).c Then
          Game(i + 1, j).s = Game(i, j).s
          Call CheckEdges(i + 1, j)
        End If
      End If
    End If
  End If
End Sub

'For every selected tile - remove it, return the score for this move

Private Function DelEdges() As Long
  Dim i As Long, j As Long, k As Long, q As Long
  
  'First remove all tiles selected in all columns
  q = 0
  For i = 0 To MaxCols
    j = MaxRows
    If Game(i, j).c <> 0 Then             'Something is left in this column
      Do While j >= 0
        If Game(i, j).s = 1 Then          'It is selected
          q = q + 1
          For k = j To 1 Step -1          'Move column down 1
            Game(i, k) = Game(i, k - 1)
          Next
          Game(i, 0).c = 0
          Game(i, 0).s = 0
        Else
          j = j - 1                       'test next row up in this column
        End If
      Loop
    End If
  Next
  
  'now move all columns left to remove empty columns
  For i = 0 To MaxCols - 1
    If Game(i, MaxRows).c = 0 Then        'Everything has been removed in this column
      For k = i + 1 To MaxCols
        If Game(k, MaxRows).c <> 0 Then Exit For  'find the first nonempty column
      Next
      If k <= MaxCols Then
        For j = 0 To MaxRows              'move all the next ones left
          Game(i, j) = Game(k, j)
          Game(k, j).c = 0: Game(k, j).s = 0
        Next
      End If
    End If
  Next
  
  'work out the score
  NLeft = NLeft - q
  If q = 1 Then
    DelEdges = -5
  Else
    DelEdges = q
    k = 4
    For j = 3 To 31 Step 4
      k = k + k
      If q > j Then DelEdges = DelEdges + k
    Next
  End If
End Function

'the routine for the drawing of the tableua. Not that it is capable of drawing a transition
'state where selected tiles are shown depressed
Private Sub DrawTableau()
  Dim i As Long, j As Long, ink As Long
  Dim x0 As Long, y0 As Long, x1 As Long, y1 As Long
  
  For i = 0 To MaxCols
    For j = 0 To MaxRows
      x0 = i * SqrSize: y0 = j * SqrSize
      x1 = x0 + SqrSize - 1: y1 = y0 + SqrSize - 1
      
      Picture1.Line (x0, y0)-(x1, y1), Inks(Game(i, j).c), BF
      If Game(i, j).c <> 0 Then
        If Game(i, j).s = 0 Then ink = Inks(1) Else ink = Inks(2)
        Picture1.Line (x0, y0)-(x1, y0), ink
        Picture1.Line (x0, y0)-(x0, y1), ink
        If Game(i, j).s = 0 Then ink = Inks(2) Else ink = Inks(1)
        Picture1.Line (x0, y1)-(x1, y1), ink
        Picture1.Line (x1, y0)-(x1, y1), ink
      End If
    Next
  Next
  Me.Caption = "TileFall  - Score= " & Score & "   BestScore= " & HiScore & "  AllTime HiScore= " & HOFHIScore
End Sub

'handles every mousedown click
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim i As Long, j As Long
  
  i = x \ SqrSize: j = y \ SqrSize
  If Game(i, j).c = 0 Then Exit Sub
  If Game(i, j).s = 0 Then    'wasnt selected
    Game(i, j).s = 1
    Call CheckEdges(i, j)
    Score = Score + DelEdges()
    If NLeft = 0 Then     'the end of a game
      If Score > HiScore Then HiScore = Score
      If HiScore > HOFHIScore Then HOFHIScore = HiScore
    End If
    Call DrawTableau
  End If
  If NLeft = 0 And Score > HallOfFame(CurrGame).Scores(9).Score Then Call UpdateHallOfFame
End Sub

'If the score is high enough replace it in the hall of fame
Private Sub UpdateHallOfFame()
  Dim i As Long, k As Long
  
  With HallOfFame(CurrGame)
    For i = 0 To 9
      If Score > .Scores(i).Score Then k = i: Exit For
    Next
    
    For i = 9 To k Step -1
      If i > 0 Then .Scores(i) = .Scores(i - 1)
    Next
    .Scores(k).Name = Username
    .Scores(k).Date = Now()
    .Scores(k).Score = Score
  End With
  Call TopTen_Click(0)
End Sub

  ' the menu handling functions

Private Sub Colours_3_Click()
  NColours = 3
  Call NewGame_Click(0)
End Sub

Private Sub Colours_4_Click()
  NColours = 4
  Call NewGame_Click(0)
End Sub

Private Sub Colours_5_Click()
  NColours = 5
  Call NewGame_Click(0)
End Sub

Private Sub Colours_6_Click()
  NColours = 6
  Call NewGame_Click(0)
End Sub

Private Sub Size_Small_Click()
  MaxCols = 9
  MaxRows = 5
  SqrSize = 60
  GameSize = GAME_SMALL
  Call NewGame_Click(0)
End Sub

Private Sub Size_Medium_Click()
  MaxCols = 14
  MaxRows = 8
  SqrSize = 40
  GameSize = GAME_MEDIUM
  Call NewGame_Click(0)
End Sub

Private Sub Size_Large_Click()
  MaxCols = 19
  MaxRows = 11
  SqrSize = 30
  GameSize = GAME_LARGE
  Call NewGame_Click(0)
End Sub

Private Sub Size_Huge_Click()
  MaxCols = 29
  MaxRows = 17
  SqrSize = 20
  GameSize = GAME_HUGE
  Call NewGame_Click(0)
End Sub

Private Sub TopTen_Click(Index As Integer)
  Dim i As Long
  
  For i = 0 To 9
    With HallOfFame(CurrGame).Scores(i)
      Form2.List1.AddItem Right$(" " + Str$(i + 1), 2) & " " & .Name & " " & .Score & " " & .Date
    End With
  Next
  Form2.Show vbModal, Form1
End Sub

