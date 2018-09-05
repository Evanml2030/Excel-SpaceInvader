Attribute VB_Name = "GameLogic"
Option Explicit
Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)

Sub RunGame()
Dim newBoard                        As GameBoard
Dim shipObj                         As Ship
Dim ShipCntrl                       As Control
Dim startTime                       As Long
Dim endTime                         As Long
Dim x                               As Long

    Set newBoard = New GameBoard
    newBoard.Show vbModeless
    
    ScaleItems.MaxSize = 60
    
    Set ShipCntrl = SHLoadShipOntoGameBoard.LoadShipOntoBoard(newBoard)
    
    startTime = timeGetTime
    
    Do While x < 100
        endTime = timeGetTime
        If (endTime - startTime) > 2000 Then
            startTime = endTime
            SOLoadSpaceObjectOntoGameBoard.LoadSpaceObjectOntoBoard newBoard
        End If
        CollisionsMissleSpaceObject.HandleMissleSpaceObjectCollisions newBoard
        If CollisionsShipSpaceObject.HandleShipSpaceObjectCollisions(newBoard) Then Exit Do
        SOMoveSpaceObjects.MoveSpaceObjects newBoard
        MMoveMissles.MoveMissleObjects newBoard
        DoEvents
        Sleep 25
    Loop
End Sub

Public Sub HandleSendKeys(ByRef board As GameBoard, ByRef caseNum As Long)
    Select Case caseNum
        Case "37"
            SHMoveShip.moveShipLeft board
        Case "39"
            SHMoveShip.moveShipRight board
        Case "32"
            MLoadMissleObjectOntoBoard.LoadMissleObjectOntoBoard board
            MissleCount.IncrementMissleCount
            ChangeBoardLabelMissleCount board
    End Select
End Sub

Private Sub ChangeBoardLabelMissleCount(ByRef board As GameBoard)
    board.MissleCount.Caption = CStr(25 - MissleCount.Count)
End Sub



