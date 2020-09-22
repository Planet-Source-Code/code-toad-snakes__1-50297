VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2970
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   2970
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Slider sld 
      Height          =   255
      Left            =   1260
      TabIndex        =   7
      Top             =   1200
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   450
      _Version        =   393216
      Min             =   1
      Max             =   1000
      SelStart        =   250
      TickStyle       =   3
      Value           =   125
   End
   Begin VB.ComboBox cboLevel 
      Height          =   315
      ItemData        =   "frmOptions.frx":0000
      Left            =   1260
      List            =   "frmOptions.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   420
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   1140
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   1560
      Width           =   975
   End
   Begin VB.ComboBox cboPlayers 
      Height          =   315
      ItemData        =   "frmOptions.frx":0004
      Left            =   1260
      List            =   "frmOptions.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   60
      Width           =   855
   End
   Begin VB.Label lblInfo 
      Caption         =   "Fast"
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   9
      Top             =   960
      Width           =   375
   End
   Begin VB.Label lblInfo 
      Caption         =   "Slow"
      Height          =   255
      Index           =   3
      Left            =   2220
      TabIndex        =   8
      Top             =   960
      Width           =   375
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Snake Speed:"
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   6
      Top             =   1200
      Width           =   1155
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Start on Level:"
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   4
      Top             =   480
      Width           =   1155
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "# Players:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
game.playerCount = cboPlayers.Text
game.level = cboLevel.Text
game.gameSpeed = sld.Value
Me.Hide
DoEvents
Unload Me

End Sub

Private Sub Form_Load()
Dim str As String

cboPlayers.Text = "1"

str = Dir(App.Path & "\screens\*.lvl")
Do While str <> ""
  'if this is not an informational screen
  If UCase(Right$(str, 4)) <> "LVLI" Then
    cboLevel.AddItem Left$(str, Len(str) - 4)
  End If
  str = Dir
Loop
cboLevel.Text = cboLevel.List(0)
End Sub

