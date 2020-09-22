VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private pPicBoard     As PictureBox 'The playfield
Private pPicTile      As PictureBox 'Picture that defines what floor (background) will look like
Private pPicWall      As PictureBox 'Picture that defines what wall will look like
Private pPicFood      As PictureBox 'Picture that defines what food will look like
Private pPicSegment   As PictureBox 'Picture used to display a segment of snake 1
Private pPicSegment2  As PictureBox 'Picture used to display a segment of snake 2
Private pRowCount     As Long       '# of rows in the playfield
Private pColCount     As Long       '# of cols in the playfield
Private pBoard()      As Integer    'An array that will hold information about the current board
                                    'pBoard(X,Y)=0 : this cell should display the background
                                    'pBoard(X,Y)=1 : this cell should display a wall
                                    'pBoard(X,Y)=2 : this cell should display food
Private pFoodX        As Integer    'The column that the food is in
Private pFoodY        As Integer    'The row that the food is in
Private pPlayerCount  As Integer    'The number of players in the game
Private pLevel        As Integer    'The level to load
Private pFoodPlaced   As Integer    'The amount of food placed on the ground so far this level
Private pFoodRequired As Integer    'The amount of food players are required to pick up on this level
Private pLevelLoaded  As Boolean    'Check this variable to determine whether or not a level has been loaded
Private pGameSpeed    As Long       'lower number, faster speed
 
'pBoard(x,y) tile 0 on/off
Public snakes     As New Collection
'Public snake      As New clsSnake

Public Sub addSnake()
'Add a snake to the snakes collection
Dim snake As New clsSnake
snakes.Add snake
End Sub

Public Property Let board(picBoard As PictureBox)
Set pPicBoard = picBoard
pPicBoard.Width = Me.cellWidth * Me.colCount * Screen.TwipsPerPixelX
pPicBoard.Height = Me.cellHeight * Me.rowCount * Screen.TwipsPerPixelY
End Property

Public Property Let segment(picSegment As PictureBox)
Set pPicSegment = picSegment
End Property

Public Property Let segment2(picSegment As PictureBox)
Set pPicSegment2 = picSegment
End Property


Public Property Let food(picFood As PictureBox)
Set pPicFood = picFood
End Property

Public Property Let wall(picWall As PictureBox)
Set pPicWall = picWall
End Property

Public Property Let tile(picTile As PictureBox)
Set pPicTile = picTile
End Property

Public Property Get foodLeft() As Integer
Dim intX As Integer
Dim intY As Integer
Dim intTemp As Integer
For intX = 0 To Me.colCount - 1
  For intY = 0 To Me.rowCount - 1
    If pBoard(intX, intY) = 2 Then
      intTemp = intTemp + 1
    End If
  Next intY
Next intX
foodLeft = intTemp
End Property

Public Property Get cellWidth() As Long
cellWidth = Me.boardWidth \ Me.colCount 'scale width of the board picture divided by # columns
End Property

Public Property Get cellHeight() As Long
cellHeight = Me.boardHeight \ Me.rowCount 'scale height of the board picture divided by # rows
End Property

Public Property Get boardWidth() As Long
boardWidth = pPicBoard.ScaleWidth
End Property

Public Property Get boardHeight() As Long
boardHeight = pPicBoard.ScaleHeight
End Property

Public Property Get rowCount() As Long
rowCount = pRowCount
End Property

Public Property Let rowCount(rowCount As Long)
pRowCount = rowCount
End Property

Public Property Get gameSpeed() As Long
gameSpeed = pGameSpeed
End Property

Public Property Let gameSpeed(gameSpeed As Long)
pGameSpeed = gameSpeed
End Property

Public Property Get colCount() As Long
colCount = pColCount
End Property

Public Property Let colCount(colCount As Long)
pColCount = colCount
End Property

Public Property Get foodRequired() As Long
foodRequired = pFoodRequired
End Property

Public Property Let foodRequired(foodRequired As Long)
pFoodRequired = foodRequired
End Property

Public Property Get foodPlaced() As Long
foodPlaced = pFoodPlaced
End Property

Public Property Let foodPlaced(foodPlaced As Long)
pFoodPlaced = foodPlaced
End Property

Public Property Get playerCount() As Long
playerCount = pPlayerCount
End Property

Public Property Let playerCount(playerCount As Long)
pPlayerCount = playerCount
End Property

Public Property Get level() As Long
level = pLevel
End Property

Public Property Let level(level As Long)
pLevel = level
End Property

'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
'Public subs and functions
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
Public Sub drawGrid()
'This is just a test function, it just draws black lines around each cell
Dim lngCounter As Long

For lngCounter = 0 To Me.boardHeight Step Me.cellHeight
  BitBlt pPicBoard.hdc, 0, lngCounter, Me.boardWidth, 1, pPicBoard.hdc, 0, 0, vbBlackness
Next lngCounter

For lngCounter = 0 To Me.boardWidth Step Me.cellWidth
  BitBlt pPicBoard.hdc, lngCounter, 0, 1, Me.boardHeight, pPicBoard.hdc, 0, 0, vbBlackness
Next lngCounter

End Sub

Public Sub drawBackground()
Dim lngCounterY As Long
Dim lngCounterX As Long
Dim lngCellX    As Long
Dim lngCellY    As Long

'Loop through each row
For lngCounterY = 0 To Me.boardHeight - 1 Step Me.cellHeight
  'loop through each column
  lngCellX = 0
  For lngCounterX = 0 To Me.boardWidth - 1 Step Me.cellWidth
    Select Case pBoard(lngCellX, lngCellY)
      Case 0 'Background
        BitBlt pPicBoard.hdc, lngCounterX, lngCounterY, pPicTile.ScaleWidth, pPicTile.ScaleHeight, pPicTile.hdc, 0, 0, vbSrcCopy
      Case 1 'Wall
        BitBlt pPicBoard.hdc, lngCounterX, lngCounterY, pPicTile.ScaleWidth, pPicTile.ScaleHeight, pPicWall.hdc, 0, 0, vbSrcCopy
      Case 2 'Food
        BitBlt pPicBoard.hdc, lngCounterX, lngCounterY, pPicTile.ScaleWidth, pPicTile.ScaleHeight, pPicFood.hdc, 0, 0, vbSrcCopy
    End Select
    lngCellX = lngCellX + 1
  Next lngCounterX
  lngCellY = lngCellY + 1
Next lngCounterY

End Sub
Public Sub initialize()
'initialize the snake objects (reset scores, etc for snakes)
Dim intCounter As Integer
'First, remove any snakes that may exist from previous games
If snakes.Count > 0 Then
  For intCounter = snakes.Count To 1 Step -1
    snakes.Remove intCounter
  Next intCounter
End If
'Now, add correct # of snakes
For intCounter = 1 To pPlayerCount
  Me.addSnake
Next intCounter
End Sub

Public Sub saveLevel(strLevel As String)
Dim intFileNumber As Integer
Dim strLine       As String
Dim strChar       As String
Dim intX          As Integer
Dim intY          As Integer
Dim intCounter    As Integer
Dim intS1X        As Integer 'Snake 1 x pos
Dim intS1Y        As Integer 'snake 1 y pos
Dim intS1L        As Integer 'snake 1 length
Dim intS1D        As Integer 'snake 1 direction
Dim intS2X        As Integer 'Snake 2 x pos
Dim intS2Y        As Integer 'snake 2 y pos
Dim intS2L        As Integer 'snake 2 length
Dim intS2D        As Integer 'snake 2 direction
Dim str           As String

Dim s As clsSnake

'ReDim pBoard(0 To Me.colCount - 1, 0 To Me.rowCount - 1)

intFileNumber = FreeFile
Open App.Path & "\" & strLevel & ".txt" For Output As intFileNumber

For intY = 0 To Me.rowCount - 1
  str = ""
  For intX = 0 To Me.colCount - 1
    str = str & pBoard(intX, intY)
  Next intX
  Print #intFileNumber, str
Next intY
Close intFileNumber
End Sub

Public Sub loadLevel()
Dim intFileNumber As Integer
Dim strLine       As String
Dim strChar       As String
Dim intX          As Integer
Dim intY          As Integer
Dim intCounter    As Integer
Dim intS1X        As Integer 'Snake 1 x pos
Dim intS1Y        As Integer 'snake 1 y pos
Dim intS1L        As Integer 'snake 1 length
Dim intS1D        As Integer 'snake 1 direction
Dim intS2X        As Integer 'Snake 2 x pos
Dim intS2Y        As Integer 'snake 2 y pos
Dim intS2L        As Integer 'snake 2 length
Dim intS2D        As Integer 'snake 2 direction

Dim s As clsSnake

ReDim pBoard(0 To Me.colCount - 1, 0 To Me.rowCount - 1)

intFileNumber = FreeFile
Open App.Path & "\screens\" & Me.level & ".lvl" For Input As intFileNumber

intY = 0
Do While Not EOF(intFileNumber)
  intX = 0
  Line Input #intFileNumber, strLine
  Select Case UCase(Left$(strLine, 4))
    Case "SN1D" 'snake 1 direction
      intS1D = Right$(strLine, Len(strLine) - 5)
    Case "SN1L" 'snake 1 length
      intS1L = Right$(strLine, Len(strLine) - 5)
    Case "SN1X" 'snake 1 head x pos
      intS1X = Right$(strLine, Len(strLine) - 5)
    Case "SN1Y" 'snake 1 head y pos
      intS1Y = Right$(strLine, Len(strLine) - 5)
    Case "SN2D" 'snake 2 direction
      intS2D = Right$(strLine, Len(strLine) - 5)
    Case "SN2L" 'snake 2 length
      intS2L = Right$(strLine, Len(strLine) - 5)
    Case "SN2X" 'snake 2 head x pos
      intS2X = Right$(strLine, Len(strLine) - 5)
    Case "SN2Y" 'snake 2 head y pos
      intS2Y = Right$(strLine, Len(strLine) - 5)
    
    Case "FOOD" '# of pieces of food required to be eaten on this level
      pFoodRequired = Right$(strLine, Len(strLine) - 5)
    Case Else 'must be a line in the file that defines how to display the board
      Do While strLine <> ""
        strChar = Left$(strLine, 1)
        strLine = Right$(strLine, Len(strLine) - 1)
        If strChar <> "0" And strChar <> "1" Then
          Debug.Print
        End If
        pBoard(intX, intY) = strChar
        intX = intX + 1
      Loop
      intY = intY + 1
  End Select
Loop
Close intFileNumber

Me.drawBackground 'Draw the game board

pFoodPlaced = 0 'reset total food placed to zero

'Now, initialize snake positions
intCounter = 1
For Each s In snakes
  If intCounter = 1 Then
    s.initialize intS1X, intS1Y
    s.increaseLength intS1L - 1
    s.lastDirection = intS1D
  Else
    s.initialize intS2X, intS2Y
    s.increaseLength intS2L - 1
    s.lastDirection = intS2D
  End If
  intCounter = intCounter + 1
Next s

pLevelLoaded = True
End Sub

Public Sub loadLevelInformational(strLevel As String)
'Loads a screen that displays info, such as player 1 died
Dim intFileNumber As Integer
Dim strLine       As String
Dim strChar       As String
Dim intX          As Integer
Dim intY          As Integer
Dim intCounter    As Integer
Dim intS1X        As Integer 'Snake 1 x pos
Dim intS1Y        As Integer 'snake 1 y pos
Dim intS1L        As Integer 'snake 1 length
Dim intS1D        As Integer 'snake 1 direction
Dim intS2X        As Integer 'Snake 2 x pos
Dim intS2Y        As Integer 'snake 2 y pos
Dim intS2L        As Integer 'snake 2 length
Dim intS2D        As Integer 'snake 2 direction
Dim pBoard()      As Integer
Dim blnContinue   As Boolean
Dim lngLastTickCount As Long
Dim lngCounterX   As Long
Dim lngCounterY   As Long
Dim lngCellX      As Long
Dim lngCellY      As Long
Dim s As clsSnake

ReDim pBoard(0 To Me.colCount - 1, 0 To Me.rowCount - 1)

intFileNumber = FreeFile
Open App.Path & "\screens\" & strLevel & ".lvli" For Input As intFileNumber

intY = 0
Do While Not EOF(intFileNumber)
  intX = 0
  Line Input #intFileNumber, strLine
  Select Case UCase(Left$(strLine, 4))
    Case "SN1D"
      intS1D = Right$(strLine, Len(strLine) - 5)
    Case "SN1L"
      intS1L = Right$(strLine, Len(strLine) - 5)
    Case "SN1X"
      intS1X = Right$(strLine, Len(strLine) - 5)
    Case "SN1Y"
      intS1Y = Right$(strLine, Len(strLine) - 5)
    Case "SN2D"
      intS2D = Right$(strLine, Len(strLine) - 5)
    Case "SN2L"
      intS2L = Right$(strLine, Len(strLine) - 5)
    Case "SN2X"
      intS2X = Right$(strLine, Len(strLine) - 5)
    Case "SN2Y"
      intS2Y = Right$(strLine, Len(strLine) - 5)
    
    Case "FOOD"
      pFoodRequired = Right$(strLine, Len(strLine) - 5)
    Case Else
      Do While strLine <> ""
        strChar = Left$(strLine, 1)
        strLine = Right$(strLine, Len(strLine) - 1)
        If strChar <> "0" And strChar <> "1" Then
          Debug.Print
        End If
        pBoard(intX, intY) = strChar
        intX = intX + 1
      Loop
      intY = intY + 1
  End Select
Loop

Close intFileNumber

'draw the information
For lngCounterY = 0 To Me.boardHeight - 1 Step Me.cellHeight
  lngCellX = 0
  For lngCounterX = 0 To Me.boardWidth - 1 Step Me.cellWidth
    Select Case pBoard(lngCellX, lngCellY)
      Case 0 'Background
        BitBlt pPicBoard.hdc, lngCounterX, lngCounterY, pPicTile.ScaleWidth, pPicTile.ScaleHeight, pPicTile.hdc, 0, 0, vbSrcCopy
      Case 1 'Wall
        BitBlt pPicBoard.hdc, lngCounterX, lngCounterY, pPicTile.ScaleWidth, pPicTile.ScaleHeight, pPicWall.hdc, 0, 0, vbSrcCopy
      Case 2 'Food
        BitBlt pPicBoard.hdc, lngCounterX, lngCounterY, pPicTile.ScaleWidth, pPicTile.ScaleHeight, pPicFood.hdc, 0, 0, vbSrcCopy
    End Select
    lngCellX = lngCellX + 1
    lngLastTickCount = GetTickCount()
  Next lngCounterX
   Sleep 1
  lngCellY = lngCellY + 1
Next lngCounterY
End Sub

Public Sub setBoardData(intX As Integer, intY As Integer, intData As Integer)
pBoard(intX, intY) = intData
Me.drawBackground
End Sub

Public Function getBoardData(intX As Integer, intY As Integer) As Integer
getBoardData = pBoard(intX, intY)
End Function

Public Sub initializeTurn()
Dim intFileNumber As Integer
Dim strLine       As String
Dim intCounter    As Integer
Dim intS1X        As Integer 'Snake 1 x pos
Dim intS1Y        As Integer 'snake 1 y pos
Dim intS1L        As Integer 'snake 1 length
Dim intS1D        As Integer 'snake 1 direction
Dim intS2X        As Integer 'Snake 2 x pos
Dim intS2Y        As Integer 'snake 2 y pos
Dim intS2L        As Integer 'snake 2 length
Dim intS2D        As Integer 'snake 2 direction

Dim s As clsSnake

intFileNumber = FreeFile
Open App.Path & "\screens\" & Me.level & ".lvl" For Input As intFileNumber

Do While Not EOF(intFileNumber)
  Line Input #intFileNumber, strLine
  Select Case UCase(Left$(strLine, 4))
    Case "SN1D" 'Snake 1 Direction
      intS1D = Right$(strLine, Len(strLine) - 5)
    Case "SN1X" 'Snake 1 Head X position
      intS1X = Right$(strLine, Len(strLine) - 5)
    Case "SN1Y" 'Snake 1 Head Y Position
      intS1Y = Right$(strLine, Len(strLine) - 5)
    Case "SN1L" 'Snake 1 Length
      intS1L = Right$(strLine, Len(strLine) - 5)
    Case "SN2D" 'Snake 2 Direction
      intS2D = Right$(strLine, Len(strLine) - 5)
    Case "SN2X" 'Snake 2 Head X position
      intS2X = Right$(strLine, Len(strLine) - 5)
    Case "SN2Y" 'Snake 2 Head Y position
      intS2Y = Right$(strLine, Len(strLine) - 5)
    Case "SN2L" 'Snake 2 Length
      intS2L = Right$(strLine, Len(strLine) - 5)
    
  End Select
Loop
Close intFileNumber

Me.drawBackground 'redraw the board


'Now, initialize snake positions
intCounter = 1
For Each s In snakes
  If intCounter = 1 Then
    s.initialize intS1X, intS1Y 'Initialize the head position
    s.lastDirection = intS1D    'Set the direction for the snake to start off in
    s.increaseLength intS1L - 1
  Else
    s.initialize intS2X, intS2Y 'Initialize the head position
    s.lastDirection = intS2D    'Set the direction for the snake to start off in
    s.increaseLength intS2L - 1
  End If
  intCounter = intCounter + 1
Next s

End Sub
Public Sub drawSnake(intSnake As Integer)
'•••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
'Purpose: Use this routine after the snake has moved to a new position (the    '
'         array will reflect the new position, but the game display will not)  '
'         to draw the snake's new position.                                    '
'Notes:   Only the head and tail position need to be updated on the screen.    '
'         The head needs to be drawn in the new position, the tail will be     '
'         erased, by drawing the background over the correct tile              '
'•••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
Dim intCounter As Integer
Dim lngX       As Long
Dim lngY       As Long
Dim snake      As clsSnake

Set snake = snakes.Item(intSnake) 'Set this snake object to the snake we are to draw

'••••••••••••••••••••••••••••••••••••••
'Draw the head in it's new position
'••••••••••••••••••••••••••••••••••••••
lngX = snake.segX(0)
lngY = snake.segY(0)
If intSnake = 1 Then 'draw player 1 segment
  BitBlt pPicBoard.hdc, lngX * Me.cellWidth, lngY * Me.cellHeight, Me.cellWidth, Me.cellHeight, pPicSegment.hdc, 0, 0, vbSrcCopy
Else 'player 2
  BitBlt pPicBoard.hdc, lngX * Me.cellWidth, lngY * Me.cellHeight, Me.cellWidth, Me.cellHeight, pPicSegment2.hdc, 0, 0, vbSrcCopy
End If

'••••••••••••••••••••••••••••••••••••••
'Erase the tail from it's old position
'••••••••••••••••••••••••••••••••••••••
lngX = snake.segXPrevious(snake.length - 1) 'Get the x position where the tail was before it moved
lngY = snake.segYPrevious(snake.length - 1) 'Get the y position where the tail was before it moved
BitBlt pPicBoard.hdc, lngX * Me.cellWidth, lngY * Me.cellHeight, Me.cellWidth, Me.cellHeight, pPicTile.hdc, 0, 0, vbSrcCopy

End Sub

Public Function detectCollision(intSnake As Integer) As Long
'detect head collision with food, wall, or another snake
'call this before drawSnake
'return: 0 - no collision
'        1 - collide with wall
'        2 - collide with food
'        3 - collide with another snake

Dim intCounter As Integer
Dim intCounter2 As Integer
Dim lngX       As Long
Dim lngY       As Long
Dim snake      As clsSnake
Dim blnCollideSnake As Boolean
Dim blnCollideWall  As Boolean
Dim blnCollideFood  As Boolean

Set snake = snakes.Item(intSnake)

'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
'first, see if the snake has collided with food or the wall
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
lngX = snake.segX(0) 'get head x
lngY = snake.segY(0) 'get head y
Select Case pBoard(lngX, lngY)
  Case 1 'wall
    blnCollideWall = True
  Case 2 'food
    blnCollideFood = True
End Select
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
'now, see if the snake has collided with another snake
'••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
For intCounter = 1 To snakes.Count
  'if this is a snake besides the one we are checking
  If intCounter <> intSnake Then
    'loop through each segment of the snake we are trying to determine if we collided with
    For intCounter2 = 0 To snakes.Item(intCounter).length - 1
      If snakes.Item(intCounter).segX(intCounter2) = lngX And snakes.Item(intCounter).segY(intCounter2) = lngY Then
        blnCollideSnake = True
      End If
    Next intCounter2
  Else 'check all positions against the head of the same snake to see if it collided with itself
    'loop through each segment of the snake (EXCEPT HEAD, so start counter at 1) we are trying to determine if we collided with
    For intCounter2 = 1 To snakes.Item(intCounter).length - 1
      If snakes.Item(intCounter).segX(intCounter2) = lngX And snakes.Item(intCounter).segY(intCounter2) = lngY Then
        blnCollideSnake = True
      End If
    Next intCounter2
  
  End If
Next intCounter

If blnCollideWall = False And blnCollideSnake = False And blnCollideFood = False Then
  detectCollision = 0
  Exit Function
End If

If blnCollideWall = True Then
  detectCollision = 1
  Exit Function
End If

If blnCollideFood = True Then
  snakes.Item(intSnake).score = snakes.Item(intSnake).score + (1000 \ Me.gameSpeed)
  'replace food with background
  pBoard(lngX, lngY) = 0
  detectCollision = 2
  Exit Function
End If
  
If blnCollideSnake = True Then
  detectCollision = 3
  Exit Function
End If
End Function

Public Sub drawBoard()
StretchBlt pPicBoard.hdc, 0, 0, Me.boardWidth, Me.boardHeight, pPicBoard.hdc, 0, 0, 1, 1, vbSrcCopy
End Sub

Public Sub placeFood()
Dim intCounter As Integer
Dim lngX       As Long
Dim lngY       As Long
Dim blnContinue As Boolean
Dim snake       As clsSnake


If pLevelLoaded = False Then
  'MsgBox "No Level has been loaded."
  Exit Sub
End If
  
Randomize
Do Until blnContinue
  lngX = Int((Me.colCount - 1 - 0 + 1) * Rnd + 0)
  lngY = Int((Me.rowCount - 1 - 0 + 1) * Rnd + 0)
  'If this position is not a wall or already occupied by food
  If pBoard(lngX, lngY) = 0 Then
    blnContinue = True
    'if this position is occupied by a segment of a snake
    'try again
    For Each snake In snakes
      For intCounter = 0 To snake.length - 1
        If snake.segX(intCounter) = lngX And snake.segY(intCounter) = lngY Then
          blnContinue = False
          intCounter = snake.length
        End If
      Next intCounter
    Next snake
  End If
Loop
pFoodPlaced = pFoodPlaced + 1 'Total food placed

BitBlt pPicBoard.hdc, lngX * Me.cellWidth, lngY * Me.cellHeight, Me.cellWidth, Me.cellHeight, pPicFood.hdc, 0, 0, vbSrcCopy
pBoard(lngX, lngY) = 2 '2=food, 1 = wall, 0=background
End Sub

Private Sub Class_Initialize()
pPlayerCount = 1
pLevel = 1
pLevelLoaded = False
End Sub
