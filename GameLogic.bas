Attribute VB_Name = "GameLogic"
Option Explicit
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)

Sub RunGame()

Dim board                               As GameBoard
Dim sleepWatch                          As stopWatch
Dim generateIncSpaceObjects             As stopWatch
Const interval = 1.25

Set board = New GameBoard
board.Show vbModeless

CreateGameItem.CreateGameItem board, Ship

Set generateIncSpaceObjects = New stopWatch
generateIncSpaceObjects.Start
Set sleepWatch = New stopWatch
sleepWatch.Start

Do
    MoveSpaceObjects.MoveMissiles
    MoveSpaceObjects.MoveIncomingSpaceObjects
    If CheckCollisions.HandleShipIncSpaceObjectCollisions Then Exit Do
    HandleMissileIncSpaceObjectCollisions board
    
    If Format(generateIncSpaceObjects.Elapsed, "0.000000") > 0.75 Then
        CreateGameItem.CreateGameItem board, Application.WorksheetFunction.RandBetween(1, 3)
        generateIncSpaceObjects.Restart
    End If

    If Format(sleepWatch.Elapsed, "0.000000") < interval Then
        Sleep interval - Format(sleepWatch.Elapsed, "0.000000")
        sleepWatch.Restart
    End If
DoEvents
Loop

End Sub

Public Sub HandleSendKeys(ByRef board As GameBoard, ByRef caseNum As Long)
    Select Case caseNum
        Case 37
            MoveSpaceObjects.MoveShip left, board
        Case 39
            MoveSpaceObjects.MoveShip Right, board
        Case 32
            If MissileCount.Count < 25 Then
                CreateGameItem.CreateGameItem board, missile
            End If
    End Select
End Sub

