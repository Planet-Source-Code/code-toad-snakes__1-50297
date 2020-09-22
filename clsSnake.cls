VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSnake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private pPos() As Long
Private pLastPos() As Long

  'pPos(0,0)=head x
  'pPos(1,0)=head y
  
  'for n=1 to ...
  'pPos(0,n)=segment x
  'ppos(1,n)=segment y
Private pScore        As Long
Private pLength       As Long
Private pLastDirection As iDirection
Private pTurnsRemaining As Long
Private pCollisionOccured As Boolean

Public Enum iDirection
  iLeft = 0
  iRight = 1
  iup = 2
  iDown = 3
End Enum

Public Property Get score() As Long
score = pScore
End Property

Public Property Let score(score As Long)
pScore = score
End Property


Public Property Get turnsRemaining() As Long
turnsRemaining = pTurnsRemaining
End Property

Public Property Let turnsRemaining(turnsRemaining As Long)
pTurnsRemaining = turnsRemaining
End Property

Public Property Get collisionOccured() As Boolean
collisionOccured = pCollisionOccured
End Property

Public Property Let collisionOccured(collisionOccured As Boolean)
pCollisionOccured = collisionOccured
End Property

Public Property Get length() As Long
'length of snake = max position of 2nd dimension of array (+1, because array is 0 based)
length = UBound(pPos, 2) + 1
End Property

Public Property Get segX(intSegment As Integer) As Long
segX = pPos(0, intSegment)
End Property
Public Property Get segY(intSegment As Integer) As Long
segY = pPos(1, intSegment)
End Property

Public Property Get segXPrevious(intSegment As Integer) As Long
segXPrevious = pLastPos(0, intSegment)
End Property
Public Property Get segYPrevious(intSegment As Integer) As Long
segYPrevious = pLastPos(1, intSegment)
End Property

Public Sub increaseLength(Optional intLength As Integer = 1)
'Add an array position.  Make the last segment of the snake
'appear in the same position as the "old" last position
Dim intCounter As Integer

For intCounter = 1 To intLength
  ReDim Preserve pPos(0 To 1, 0 To UBound(pPos, 2) + 1)
  pPos(0, UBound(pPos, 2)) = pPos(0, UBound(pPos, 2) - 1)
  pPos(1, UBound(pPos, 2)) = pPos(1, UBound(pPos, 2) - 1)

  ReDim Preserve pLastPos(0 To 1, 0 To UBound(pLastPos, 2) + 1)
  pLastPos(0, UBound(pLastPos, 2)) = pLastPos(0, UBound(pLastPos, 2) - 1)
  pLastPos(1, UBound(pLastPos, 2)) = pLastPos(1, UBound(pLastPos, 2) - 1)

Next intCounter

End Sub

Public Sub decreaseLength()
ReDim Preserve pPos(0 To 1, 0 To UBound(pPos, 2) - 1)
End Sub

Public Function moveSnake(direction As iDirection) As Long
'returns:
'1=snake move succesfully
'0=snake cant move this direction
'2=snake collided with itself
Dim intLastX As Integer
Dim intLastY As Integer
Dim intLastXTemp As Integer
Dim intLastYTemp As Integer

Dim intCounter As Integer

pLastPos = pPos
'first see if we are trying to move snake in direction
'not allowed (for example, if the snake is headed right,
'you can't move left)
If Me.length > 1 Then
  Select Case direction
    Case iLeft
      'if first two segments have same y coord
      If pPos(1, 0) = pPos(1, 1) Then
        'if the new x position of the head overlaps next segment
        If pPos(0, 0) - 1 = pPos(0, 1) Then
          moveSnake = 0
          Exit Function
        End If
      End If
    Case iRight
      'if first two segments have same y coord
      If pPos(1, 0) = pPos(1, 1) Then
        'if the new x position of the head overlaps next segment
        If pPos(0, 0) + 1 = pPos(0, 1) Then
          moveSnake = 0
          Exit Function
        End If
      End If
    Case iup
      'if first two segments have same x coord
      If pPos(0, 0) = pPos(0, 1) Then
        'if the new y position of the head overlaps next segment
        If pPos(1, 0) - 1 = pPos(1, 1) Then
          moveSnake = 0
          Exit Function
        End If
      End If
    Case iDown
      'if first two segments have same x coord
      If pPos(0, 0) = pPos(0, 1) Then
        'if the new y position of the head overlaps next segment
        If pPos(1, 0) + 1 = pPos(1, 1) Then
          moveSnake = 0
          Exit Function
        End If
      End If
  End Select
End If

intLastX = pPos(0, 0)
intLastY = pPos(1, 0)

'Move the first segment (head)
Select Case direction
  Case iLeft
    pPos(0, 0) = pPos(0, 0) - 1
  Case iRight
    pPos(0, 0) = pPos(0, 0) + 1
  Case iup
    pPos(1, 0) = pPos(1, 0) - 1
  Case iDown
    pPos(1, 0) = pPos(1, 0) + 1
End Select
'If there are other segments
If Me.length > 1 Then
  'loop through other segments
  For intCounter = 1 To UBound(pPos, 2)
    intLastXTemp = pPos(0, intCounter)
    intLastYTemp = pPos(1, intCounter)
    pPos(0, intCounter) = intLastX
    pPos(1, intCounter) = intLastY
    intLastX = intLastXTemp
    intLastY = intLastYTemp
  Next intCounter
End If
pLastDirection = direction
End Function

Public Property Get lastDirection() As iDirection
lastDirection = pLastDirection
End Property

Public Property Let lastDirection(lastDirection As iDirection)
pLastDirection = lastDirection
End Property

Private Sub Class_Initialize()
initialize
End Sub

Public Sub initialize(Optional intX As Integer = 0, Optional intY As Integer = 0)
ReDim pPos(0 To 1, 0 To 0) As Long
ReDim pLastPos(0 To 1, 0 To 0) As Long
'pPos(0,0)=head x
'pPos(1,0)=head y

'for n=1 to ...
'pPos(0,n)=segment x
'ppos(1,n)=segment y

pPos(0, 0) = intX
pPos(1, 0) = intY
pLastPos(0, 0) = intX
pLastPos(1, 0) = intY

'pScore = 0

End Sub

