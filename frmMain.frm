VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rock Away"
   ClientHeight    =   6045
   ClientLeft      =   150
   ClientTop       =   420
   ClientWidth     =   6120
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   6120
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   5160
      Top             =   600
   End
   Begin VB.Label lblTime 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   6
      Left            =   600
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":030A
      Top             =   1320
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   5
      Left            =   600
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":160C
      Top             =   600
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   1
      Left            =   1320
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":290E
      Top             =   1320
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   46
      Left            =   3480
      MouseIcon       =   "frmMain.frx":3C10
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   45
      Left            =   2760
      MouseIcon       =   "frmMain.frx":3F1A
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   44
      Left            =   2040
      MouseIcon       =   "frmMain.frx":4224
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   39
      Left            =   3480
      MouseIcon       =   "frmMain.frx":452E
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   38
      Left            =   2760
      MouseIcon       =   "frmMain.frx":4838
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   37
      Left            =   2040
      MouseIcon       =   "frmMain.frx":4B42
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   34
      Left            =   4920
      MouseIcon       =   "frmMain.frx":4E4C
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   33
      Left            =   4200
      MouseIcon       =   "frmMain.frx":5156
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   32
      Left            =   3480
      MouseIcon       =   "frmMain.frx":5460
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   31
      Left            =   2760
      MouseIcon       =   "frmMain.frx":576A
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   30
      Left            =   2040
      MouseIcon       =   "frmMain.frx":5A74
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   29
      Left            =   1320
      MouseIcon       =   "frmMain.frx":5D7E
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   28
      Left            =   600
      MouseIcon       =   "frmMain.frx":6088
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   27
      Left            =   4920
      MouseIcon       =   "frmMain.frx":6392
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   26
      Left            =   4200
      MouseIcon       =   "frmMain.frx":669C
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   25
      Left            =   3480
      MouseIcon       =   "frmMain.frx":69A6
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   24
      Left            =   2760
      MouseIcon       =   "frmMain.frx":6CB0
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   23
      Left            =   2040
      MouseIcon       =   "frmMain.frx":6FBA
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   22
      Left            =   1320
      MouseIcon       =   "frmMain.frx":72C4
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   21
      Left            =   600
      MouseIcon       =   "frmMain.frx":75CE
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   20
      Left            =   4920
      MouseIcon       =   "frmMain.frx":78D8
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   19
      Left            =   4200
      MouseIcon       =   "frmMain.frx":7BE2
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   18
      Left            =   3480
      MouseIcon       =   "frmMain.frx":7EEC
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   17
      Left            =   2760
      MouseIcon       =   "frmMain.frx":81F6
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   16
      Left            =   2040
      MouseIcon       =   "frmMain.frx":8500
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   15
      Left            =   1320
      MouseIcon       =   "frmMain.frx":880A
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   14
      Left            =   600
      MouseIcon       =   "frmMain.frx":8B14
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   11
      Left            =   3480
      MouseIcon       =   "frmMain.frx":8E1E
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   10
      Left            =   2760
      MouseIcon       =   "frmMain.frx":9128
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   9
      Left            =   2040
      MouseIcon       =   "frmMain.frx":9432
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   0
      Left            =   1320
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":973C
      Top             =   600
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   4
      Left            =   3480
      MouseIcon       =   "frmMain.frx":AA3E
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   600
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   3
      Left            =   2760
      MouseIcon       =   "frmMain.frx":AD48
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   600
      Width           =   600
   End
   Begin VB.Image imgRock 
      Height          =   600
      Index           =   2
      Left            =   2040
      MouseIcon       =   "frmMain.frx":B052
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   600
      Width           =   600
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuHighScores 
         Caption         =   "&High Scores"
      End
      Begin VB.Menu l2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuUndo 
      Caption         =   "<< &Undo"
   End
   Begin VB.Menu mnuRedo 
      Caption         =   "&Redo >>"
   End
   Begin VB.Menu mnoOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuSound 
         Caption         =   "&Sound"
         Checked         =   -1  'True
      End
      Begin VB.Menu l1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBlack 
         Caption         =   "&Black"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuWhite 
         Caption         =   "&White"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Game(46) As Integer
Private Undo(50, 2) As Integer
Private Rocks
Private PrevPos As Integer
Private PrevX As Integer
Private PrevY As Integer
Private GameCount As Integer
Private lastMove As Integer
Private SoundOff As Boolean
Private picHole As Integer
Private picRock As Integer
Private totalTime As Double
Private endGame As Boolean

Private Sub Form_Load()
'load high scores
Call loadHighScore
'set view
Me.BackColor = &H0&
Me.lblTime.ForeColor = &HFFFFFF
picHole = 0
picRock = 1
Call ResetGame
End Sub

Private Sub cmdReset_Click()
Call ResetGame
End Sub

Private Sub imgRock_Click(Index As Integer)
Dim x As Integer
Dim y As Integer
Dim newPos As Integer
Dim midPos As Integer
'get grid coord
y = Int(Index / 7)
x = Index - (y * 7)
newPos = x + (y * 7)
'get a rock
If Game(newPos) = 1 Then
    PrevPos = newPos
    PrevX = x
    PrevY = y
    Exit Sub
    End If
'if first click, set it
If PrevPos = 0 Then Exit Sub
'check move of rock
If x = PrevX And y = PrevY + 2 Then
    'take from up to down
    midPos = x + ((y - 1) * 7)
    If midPos < 0 Or midPos > 47 Then Exit Sub
    If Game(midPos) = 0 Then Exit Sub
    Game(midPos) = 0
ElseIf x = PrevX And y = PrevY - 2 Then
    'take from down to up
    midPos = x + ((y + 1) * 7)
    If midPos < 0 Or midPos > 47 Then Exit Sub
    If Game(midPos) = 0 Then Exit Sub
ElseIf x = PrevX + 2 And y = PrevY Then
    'take form left to right
    midPos = (x - 1) + (y * 7)
    If midPos < 0 Or midPos > 47 Then Exit Sub
    If Game(midPos) = 0 Then Exit Sub
ElseIf x = PrevX - 2 And y = PrevY Then
    'take from right to left
    midPos = (x + 1) + (y * 7)
    If midPos < 0 Or midPos > 47 Then Exit Sub
    If Game(midPos) = 0 Then Exit Sub
Else
    'wrong selection
    PrevPos = 0
    PrevX = x
    PrevY = y
    Exit Sub
End If
'clear previous position
Game(PrevPos) = 0
Me.imgRock(PrevPos).Picture = Me.imgRock(picHole)
'put rock in new position
Game(newPos) = 1
Me.imgRock(newPos).Picture = Me.imgRock(picRock)
'clear middle rock
Game(midPos) = 0
Me.imgRock(midPos).Picture = Me.imgRock(picHole)
PrevPos = 0
'tick sound
If SoundOff = False Then BeginPlaySound 1
'memorize step
GameCount = GameCount + 1
lastMove = GameCount
Undo(GameCount, 0) = PrevX + (PrevY * 7)
Undo(GameCount, 1) = newPos
Undo(GameCount, 2) = midPos
'check for win
If GameCount = 31 And endGame = False Then
    'check if
    If Game(24) = 1 Then
        'last rock is in center gameboard
        endGame = True
        Me.Timer.Enabled = False
        'check for top 10
        If totalTime < Val(Scores(9, 1)) Or Val(Scores(9, 1)) = 0 Then
            Call SaveHighScore
            Else
            MsgBox "Congratulations! You got the last rock in the center of the board!" & vbCrLf & vbCrLf & "Unfortunally you were to slow for a high score.", vbExclamation, " Rock Away"
            End If
        Else
        'last rock NOT in center gameboard
        MsgBox "Very nice, but the last rock should be in the center of the board!" & vbCrLf & vbCrLf & "Try again!", vbExclamation, " Rock Away"
    End If
End If
End Sub

Private Sub ResetGame()
'refill gameboard and reset timer
Dim k As Integer
Rocks = Array(2, 3, 4, 9, 10, 11, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 37, 38, 39, 44, 45, 46)
For k = 0 To 32
    Game(Rocks(k)) = 1
    Me.imgRock(Rocks(k)).Picture = Me.imgRock(picRock)
Next k
Game(24) = 0
Me.imgRock(24).Picture = Me.imgRock(picHole)
GameCount = 0
totalTime = 0
Me.lblTime.Caption = "0"
Me.Timer.Enabled = True
endGame = False
End Sub

Private Sub rocksUndo()
'recall previous positions
If GameCount < 1 Then Exit Sub '<<<<<<<<<<<<<<
'replace jumping rock
Me.imgRock(Undo(GameCount, 0)).Picture = Me.imgRock(picRock)
Game(Undo(GameCount, 0)) = 1
'clear end jump rock
Me.imgRock(Undo(GameCount, 1)).Picture = Me.imgRock(picHole)
Game(Undo(GameCount, 1)) = 0
'replace mid rock
Me.imgRock(Undo(GameCount, 2)).Picture = Me.imgRock(picRock)
Game(Undo(GameCount, 2)) = 1
'go back on counter
GameCount = GameCount - 1
'sound
If SoundOff = False Then BeginPlaySound 1
End Sub

Private Sub rocksRedo()
'redo steps
Dim k As Integer
If GameCount >= lastMove Then Exit Sub
GameCount = GameCount + 1
'clear jumping rock
Me.imgRock(Undo(GameCount, 0)).Picture = Me.imgRock(picHole)
Game(Undo(GameCount, 0)) = 0
'set end jump rock
Me.imgRock(Undo(GameCount, 1)).Picture = Me.imgRock(picRock)
Game(Undo(GameCount, 1)) = 1
'clear mid rock
Me.imgRock(Undo(GameCount, 2)).Picture = Me.imgRock(picHole)
Game(Undo(GameCount, 2)) = 0
'sound
If SoundOff = False Then BeginPlaySound 1
End Sub

Private Sub mnuHighScores_Click()
Dim k As Integer
Me.Timer.Enabled = False
frmHighScores.Show (vbModal)
If endGame = False Then Me.Timer.Enabled = True
End Sub

Private Sub mnuNew_Click()
Call ResetGame
End Sub

Private Sub mnuRedo_Click()
Call rocksRedo
End Sub

Private Sub mnuSound_Click()
If Me.mnuSound.Checked = True Then
    Me.mnuSound.Checked = False
    SoundOff = True
    Else
    Me.mnuSound.Checked = True
    SoundOff = False
    End If
End Sub

Private Sub mnuUndo_Click()
Call rocksUndo
End Sub

Private Sub mnuWhite_Click()
'set gameboard color
Me.mnuBlack.Checked = False
Me.mnuWhite.Checked = True
Me.BackColor = &HFFFFFF
Me.lblTime.ForeColor = &H0&
picHole = 5
picRock = 6
Call Redraw
End Sub

Private Sub mnuBlack_Click()
'set gameboard color
Me.mnuBlack.Checked = True
Me.mnuWhite.Checked = False
Me.lblTime.ForeColor = &HFFFFFF
Me.BackColor = &H0&
picHole = 0
picRock = 1
Call Redraw
End Sub

Private Sub Redraw()
Dim k As Integer
For k = 0 To 32
    If Game(Rocks(k)) = 1 Then
        Me.imgRock(Rocks(k)).Picture = Me.imgRock(picRock)
        Else
        Me.imgRock(Rocks(k)).Picture = Me.imgRock(picHole)
    End If
Next
End Sub

Private Sub loadHighScore()
Dim fileO As Integer
Dim a As Integer
On Error GoTo skip
fileO = FreeFile
Open App.Path & "\highscore.dat" For Input As #fileO
a = -1
Do While Not EOF(fileO)
a = a + 1
If a > 9 Then Exit Do
Input #fileO, Scores(a, 0)
Input #fileO, Scores(a, 1)
Loop
Close #fileO
skip:
End Sub

Private Sub SaveHighScore()
Dim fileO As Integer
Dim strName As String
Dim newPos As Integer
Dim k As Integer
Dim j As Integer
On Error Resume Next
strName = InputBox("Congratulations!" & vbCrLf & vbCrLf & "You got the last rock in the center of the board!" & vbCrLf & vbCrLf & "Please enter your name", " New High Score")
If strName = "" Then strName = "Anonymous"
For k = 0 To 9
If Val(Scores(k, 1)) <> 0 And totalTime < Val(Scores(k, 1)) Then
    newPos = k
    For j = 8 To k Step -1
        Scores(j + 1, 0) = Scores(j, 0)
        Scores(j + 1, 1) = Scores(j, 1)
    Next
    Scores(newPos, 0) = strName
    Scores(newPos, 1) = Str(totalTime)
    Exit For
ElseIf Val(Scores(k, 1)) = 0 Then
    Scores(k, 0) = strName
    Scores(k, 1) = Str(totalTime)
    Exit For
End If
Next
'save highscore
fileO = FreeFile
Open App.Path & "\highscore.dat" For Output As fileO
For k = 0 To 9
Print #fileO, Scores(k, 0)
Print #fileO, Scores(k, 1)
Next
Close #fileO
frmHighScores.Show (vbModal)
End Sub

Private Sub Timer_Timer()
totalTime = totalTime + 1
Me.lblTime.Caption = Trim(Str(totalTime))
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuHelp_Click()
Me.Timer.Enabled = False
frmHelp.Show (vbModal)
If endGame = False Then Me.Timer.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmHighScores
Unload frmHelp
End Sub


