VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd 
      Caption         =   "Exit"
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Go!"
      Height          =   615
      Index           =   0
      Left            =   2520
      TabIndex        =   0
      Top             =   2400
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Private moTimer As cPerformanceTimer

Private Sub cmd_Click(Index As Integer)
    If Index Then
        Unload Me
    Else
        DoCompare
    End If
End Sub

Private Sub DoCompare()
    Const Iterations = 10000
    
    Dim loTimer As cPerformanceTimer: Set loTimer = New cPerformanceTimer
    Dim i As Long
    
    Dim lsString As String
    Dim lbTemp   As Boolean
    
    Dim ldblCompare As Double
    Dim ldblLen     As Double
    
    lsString = ""
    
    Me.Cls
    DoEvents
    '-------------------------------------------------------
    'Testing String = ""
    loTimer.TimerStart
    For i = 1 To Iterations
        lbTemp = (lsString = "")
    Next
    loTimer.TimerStop
    ldblCompare = loTimer.TimerElapsed
    Me.Print "String = """" took " & ldblCompare & " seconds."
    '-------------------------------------------------------
    DoEvents
    '-------------------------------------------------------
    'Testing LenB(String) = 0&
    loTimer.TimerStart
    For i = 1 To Iterations
        lbTemp = (LenB(lsString) = 0&)
    Next
    loTimer.TimerStop
    ldblLen = loTimer.TimerElapsed
    Me.Print "LenB(String) = 0& took " & ldblLen & " seconds."
    '-------------------------------------------------------
    DoEvents
    Me.Print "On this trial, String = """" was " & CLng(((ldblCompare - ldblLen) / ldblLen) * 100) & "% Slower."
    
End Sub
