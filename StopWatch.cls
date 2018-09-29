VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" ( _
    lpPerformanceCount As UInt64) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" ( _
    lpFrequency As UInt64) As Long

Private pFrequency As Double
Private pStartTS As UInt64
Private pEndTS As UInt64
Private pElapsed As Double
Private pRunning As Boolean

Private Type UInt64
    LowPart As Long
    HighPart As Long
End Type

Private Const BShift_32 = 4294967296# ' 2 ^ 32

Private Function U64Dbl(U64 As UInt64) As Double
    Dim lDbl As Double, hDbl As Double
    lDbl = U64.LowPart
    hDbl = U64.HighPart
    If lDbl < 0 Then lDbl = lDbl + BShift_32
    If hDbl < 0 Then hDbl = hDbl + BShift_32
    U64Dbl = lDbl + BShift_32 * hDbl
End Function

Private Sub Class_Initialize()
    Dim PerfFrequency As UInt64
    QueryPerformanceFrequency PerfFrequency
    pFrequency = U64Dbl(PerfFrequency)
End Sub

Public Property Get Elapsed() As Double
    If pRunning Then
        Dim pNow As UInt64
        QueryPerformanceCounter pNow
        Elapsed = pElapsed + (U64Dbl(pNow) - U64Dbl(pStartTS)) / pFrequency
    Else
        Elapsed = pElapsed
    End If
End Property

Public Sub Start()
    If Not pRunning Then
        QueryPerformanceCounter pStartTS
        pRunning = True
    End If
End Sub

Public Sub Pause()
    If pRunning Then
        QueryPerformanceCounter pEndTS
        pRunning = False
        pElapsed = pElapsed + (U64Dbl(pEndTS) - U64Dbl(pStartTS)) / pFrequency
    End If
End Sub

Public Sub Reset()
    pElapsed = 0
    pRunning = False
End Sub

Public Sub Restart()
    pElapsed = 0
    QueryPerformanceCounter pStartTS
    pRunning = True
End Sub

Public Property Get Running() As Boolean
   Running = pRunning
End Property

