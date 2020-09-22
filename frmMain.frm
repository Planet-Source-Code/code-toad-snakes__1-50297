VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Nib"
   ClientHeight    =   3105
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5835
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNewGame 
      Caption         =   "New Game"
      Height          =   315
      Left            =   4680
      TabIndex        =   11
      Top             =   1080
      Width           =   1035
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "Options"
      Height          =   315
      Left            =   4680
      TabIndex        =   10
      Top             =   780
      Width           =   1035
   End
   Begin VB.PictureBox picSegment2 
      AutoRedraw      =   -1  'True
      Height          =   210
      Left            =   5640
      Picture         =   "frmMain.frx":0442
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picFood 
      AutoRedraw      =   -1  'True
      Height          =   210
      Left            =   5400
      Picture         =   "frmMain.frx":0831
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picSegment 
      AutoRedraw      =   -1  'True
      Height          =   210
      Left            =   5160
      Picture         =   "frmMain.frx":0BE8
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picWall 
      AutoRedraw      =   -1  'True
      Height          =   210
      Left            =   4920
      Picture         =   "frmMain.frx":0F96
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      Height          =   210
      Left            =   4680
      Picture         =   "frmMain.frx":12C5
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picBoard 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   60
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   305
      TabIndex        =   0
      Top             =   60
      Width           =   4575
   End
   Begin VB.Label lblScore 
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   5340
      TabIndex        =   9
      Top             =   240
      Width           =   1635
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Player 2:"
      Height          =   255
      Left            =   4200
      TabIndex        =   8
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Player 1:"
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label lblScore 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   5340
      TabIndex        =   6
      Top             =   0
      Width           =   1635
   End
   Begin VB.Menu mnuFile 
      Caption         =   "Game"
      Visible         =   0   'False
      Begin VB.Menu mnuFileOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuFileNewGame 
         Caption         =   "New Game"
      End
   End
   Begin VB.Menu mnuTest 
      Caption         =   "Test"
      Visible         =   0   'False
      Begin VB.Menu mnuTestDrawGrid 
         Caption         =   "Draw Grid"
      End
      Begin VB.Menu mnuTestDrawBackground 
         Caption         =   "Draw Background"
      End
      Begin VB.Menu mnuTestPlaceFood 
         Caption         =   "Place Food"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mBlnPause As Boolean

Private Sub cmdNewGame_Click()
Dim intCounter As Integer
game.initialize 'reset scores, etc.
For intCounter = 1 To game.snakes.Count
  game.snakes(intCounter).turnsRemaining = 3
Next intCounter
game.loadLevel
gameLoop
End Sub

Private Sub cmdOptions_Click()
Load frmOptions
frmOptions.Show 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lngRet As Long

Select Case KeyCode
  'THESE ARE FOR FIRST PLAYER
  Case 37 'left
    game.snakes(1).lastDirection = iLeft
  Case 38 'up
    game.snakes(1).lastDirection = iUp
  Case 39 'right
    game.snakes(1).lastDirection = iRight
  Case 40 'down
    game.snakes(1).lastDirection = iDown
End Select


If game.playerCount > 1 Then
  Select Case KeyCode
    
    'THESE ARE FOR SECOND PLAYER
    Case 97 'Left 1
      game.snakes(2).lastDirection = iLeft
    Case 101 'Up 5
      game.snakes(2).lastDirection = iUp
    Case 99 'Right 3
      game.snakes(2).lastDirection = iRight
    Case 98 'Down 2
      game.snakes(2).lastDirection = iDown
  End Select
End If
Me.Caption = KeyCode

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 80 Or KeyAscii = 112 Then 'If user pressed 'P', pause the game
  mBlnPause = Not mBlnPause
End If


End Sub

Private Sub Form_Load()
initializeSettings
game.playerCount = 1
game.level = 1
game.gameSpeed = 125
End Sub

Private Sub mnuFileNewGame_Click()
Dim intCounter As Integer

game.initialize 'reset scores, etc.
For intCounter = 1 To game.snakes.Count
  game.snakes(intCounter).turnsRemaining = 3
Next intCounter
game.loadLevel
gameLoop
End Sub

Private Sub mnuFileOptions_Click()
Load frmOptions
frmOptions.Show 1
End Sub

Private Sub mnuTestDrawBackground_Click()
game.drawBackground
End Sub

Private Sub mnuTestDrawGrid_Click()
game.drawGrid
End Sub

Private Sub mnuTestPlaceFood_Click()
Dim blnContinue      As Boolean
Dim lngLastTickCount As Long

On Error GoTo errorHandler:
Do While True
  game.placeFood
  blnContinue = False
  Do While blnContinue = False
    If GetTickCount() - lngLastTickCount > 1 Then
      lngLastTickCount = GetTickCount
      blnContinue = True
    End If
    DoEvents
  Loop
Loop

Exit Sub
errorHandler:
  MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub gameLoop()
Dim intCounter          As Integer
Dim blnContinue         As Boolean
Dim lngLastTickCount    As Long

Dim blnContinueGameLoop As Boolean
picBoard.SetFocus
lngLastTickCount = GetTickCount

blnContinueGameLoop = True
Do While blnContinueGameLoop
  For intCounter = 1 To game.playerCount
    'show the score for each player
    lblScore(intCounter - 1).Caption = game.snakes(intCounter).score
    
    game.snakes(intCounter).collisionOccured = False
    game.snakes(intCounter).moveSnake game.snakes(intCounter).lastDirection
    Select Case game.detectCollision(intCounter)
      Case 0
        game.drawSnake (intCounter)
      Case 2
        game.snakes(intCounter).increaseLength
        game.drawSnake (intCounter)
      Case Else
        If game.snakes(intCounter).turnsRemaining = 0 Then
          blnContinueGameLoop = False
        Else
          game.snakes(intCounter).turnsRemaining = game.snakes(intCounter).turnsRemaining - 1
        End If
        game.snakes(intCounter).collisionOccured = True
    End Select
  Next intCounter
  
  For intCounter = 1 To game.playerCount
    If game.snakes(intCounter).collisionOccured = True Then
      game.loadLevelInformational "P" & intCounter & "D"
      pause 2000
      game.initializeTurn
    End If
  Next intCounter
  
  If game.foodLeft = 0 Then 'if there is no food laying on the board
    If game.foodPlaced < game.foodRequired Then 'if we haven't placed max amt of food left
      game.placeFood
    Else
      game.loadLevelInformational "lc"
      pause 2000 'show level complete screen for ~2 seconds
      game.level = game.level + 1
      game.loadLevel
    End If
  End If
  
  Do While mBlnPause = True 'Pause the game
    DoEvents
  Loop
  
  pause game.gameSpeed 'control speed of game loop
  
  'if user closed the form, end the game
  If Me.Visible = False Then
    End
  End If
Loop
game.loadLevelInformational "GmOvr" 'load the gameover screen
End Sub

Private Sub initializeSettings()
game.colCount = 30
game.rowCount = 20
game.board = picBoard
game.tile = picTile
game.wall = picWall
game.food = picFood
game.segment = picSegment
game.segment2 = picSegment2
End Sub
