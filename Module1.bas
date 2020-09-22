Attribute VB_Name = "modGlobal"
Option Explicit
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public game As New clsBoard

Public Sub pause(lngMilliseconds As Long)
'Pauses for a specified # of milliseconds
Dim blnContinue As Boolean
Dim lngLastTickCount As Long
blnContinue = False
lngLastTickCount = GetTickCount
Do While blnContinue = False
  If GetTickCount - lngLastTickCount >= lngMilliseconds Then
    lngLastTickCount = GetTickCount
    blnContinue = True
  End If
  DoEvents
Loop
End Sub
