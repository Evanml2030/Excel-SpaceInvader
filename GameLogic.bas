Attribute VB_Name = "GameLogic"
Option Explicit
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)

Sub RunGame()
Dim board                               As GameBoard
Dim sleepWatch                          As StopWatch
Dim generateIncSpaceObjectsRound1       As StopWatch
Dim generateIncSpaceObjectsRound2       As StopWatch
Const interval = 3

Set board = New GameBoard
board.Show vbModeless

CreateGameItem.CreateGameItem board, objectType.Ship

Set generateIncSpaceObjectsRound1 = New StopWatch
generateIncSpaceObjectsRound1.Start

Set generateIncSpaceObjectsRound2 = New StopWatch
generateIncSpaceObjectsRound2.Start

Set sleepWatch = New StopWatch
sleepWatch.Start

Do
    If CheckCollisions.HandleShipIncSpaceObjectCollisions Then Exit Do
    CheckCollisions.HandleMissileIncSpaceObjectCollisions board
    
    MoveSpaceObjects.MoveIncomingSpaceObjects board
    MoveSpaceObjects.MoveMissiles board

    If Format(generateIncSpaceObjectsRound1.Elapsed, "0.000000") > 1.25 Then
        CreateGameItem.CreateGameItem board, Application.WorksheetFunction.RandBetween(1, 3)
        CreateGameItem.CreateGameItem board, Application.WorksheetFunction.RandBetween(1, 3)
        CreateGameItem.CreateGameItem board, Application.WorksheetFunction.RandBetween(1, 3)
        generateIncSpaceObjectsRound1.Restart
        Score.IncrementScore
        Score.UpdateGameBoard board
    End If
    
    If Format(generateIncSpaceObjectsRound2.Elapsed, "0.000000") > 4.25 Then
        CreateGameItem.CreateGameItem board, Application.WorksheetFunction.RandBetween(1, 3)
        CreateGameItem.CreateGameItem board, Application.WorksheetFunction.RandBetween(1, 3)
        CreateGameItem.CreateGameItem board, Application.WorksheetFunction.RandBetween(1, 3)
        CreateGameItem.CreateGameItem board, Application.WorksheetFunction.RandBetween(1, 3)
        generateIncSpaceObjectsRound2.Restart
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
                CreateGameItem.CreateGameItem board, objectType.missile
            End If
    End Select
End Sub


