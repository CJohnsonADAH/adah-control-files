VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cStopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private mlngStart As Long
Private mlngEnd As Long

Public Property Get ms() As Long
    Dim lEnd As Long
    If mlngEnd = 0 Then
        lEnd = GetTickCount
    Else
        lEnd = mlngEnd
    End If
    Let ms = (lEnd - mlngStart)
End Property

Public Property Get Running() As Boolean
    Let Running = (mlngStart <> 0)
End Property

Public Property Let Running(TurnItOn As Boolean)
    If TurnItOn Then
        StartTimer
    Else
        StopTimer
    End If
End Property

Public Sub ResetTimer()
    Let mlngStart = 0
    Let mlngEnd = 0
End Sub

Public Sub StartTimer()
    Let mlngStart = GetTickCount
    Let mlngEnd = 0
End Sub

Public Sub StopTimer()
    Let mlngEnd = GetTickCount
End Sub

