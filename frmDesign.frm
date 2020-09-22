VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3960
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1140
      TabIndex        =   6
      Top             =   3300
      Width           =   1935
   End
   Begin VB.PictureBox picBoard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   60
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   305
      TabIndex        =   5
      Top             =   60
      Width           =   4575
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      Height          =   210
      Left            =   3480
      Picture         =   "frmDesign.frx":0000
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   4
      Top             =   3120
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picWall 
      AutoRedraw      =   -1  'True
      Height          =   210
      Left            =   3720
      Picture         =   "frmDesign.frx":02DF
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picSegment 
      AutoRedraw      =   -1  'True
      Height          =   210
      Left            =   3960
      Picture         =   "frmDesign.frx":060E
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   2
      Top             =   3120
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picFood 
      AutoRedraw      =   -1  'True
      Height          =   210
      Left            =   4200
      Picture         =   "frmDesign.frx":09BC
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   1
      Top             =   3120
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picSegment2 
      AutoRedraw      =   -1  'True
      Height          =   210
      Left            =   4440
      Picture         =   "frmDesign.frx":0D73
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   0
      Top             =   3120
      Visible         =   0   'False
      Width           =   210
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
game.saveLevel "p1d.txt"
End Sub

Private Sub Form_Load()
initializeSettings
game.level = 0
game.initialize
game.loadLevel
game.drawBackground
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

Private Sub picBoard_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'button: left=1, right=2
Dim intTemp As Integer
intTemp = game.getBoardData(X \ game.cellWidth, Y \ game.cellHeight)
If intTemp = 1 Then
  intTemp = 0
Else
  intTemp = 1
End If
Select Case Button
  Case 1 'left click
    game.setBoardData X \ game.cellWidth, Y \ game.cellHeight, intTemp
  Case 2 'right click
    game.setBoardData X \ game.cellWidth, Y \ game.cellHeight, intTemp
End Select

End Sub

Private Sub picBoard_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'button: left=1, right=2
'Dim intTemp As Integer
'intTemp = game.getBoardData(X \ game.cellWidth, Y \ game.cellHeight)
'If intTemp = 1 Then
'  intTemp = 0
'Else
'  intTemp = 1
'End If
Select Case Button
  Case 1 'left click
    game.setBoardData X \ game.cellWidth, Y \ game.cellHeight, 1
  Case 2 'right click
    game.setBoardData X \ game.cellWidth, Y \ game.cellHeight, 0
End Select
End Sub

