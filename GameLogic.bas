Attribute VB_Name = "GameLogic"
Option Explicit

Public Enum SpaceObjectType
    Alien = 1
    Comet = 2
    Star = 3
    Missile = 4
    Ship = 5
End Enum

Public Enum PlayerShipHit
    Hit = 1
    NotHit = 0
End Enum

Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
Public GamePiecesCollection As Scripting.Dictionary
Const Interval = 3

Sub RunGame(ByVal BoardWith As Long, ByVal BoardHeight As Long)
    Dim SleepWatch As StopWatch
    Dim GenerateIncSpaceObjectsRound1 As StopWatch
    
    Set SleepWatch = New StopWatch
    SleepWatch.Start
    
    BoardDimensions.Width = BoardWith
    BoardDimensions.Height = BoardHeight

    Set GamePiecesCollection = New Scripting.Dictionary
    
    Set GenerateIncSpaceObjectsRound1 = New StopWatch
    GenerateIncSpaceObjectsRound1.Start
    
    Set SleepWatch = New StopWatch
    SleepWatch.Start
    
    InitializePlayerShip

    Do
        GameBoard.RefreshGameBoard GamePiecesCollection
        MoveSpaceObjects.MoveIncomingSpaceObjectsAndMissiles
        
        Collisions.HandleMissileCollisions
        Collisions.HandleShipCollisions
        
        If Format(GenerateIncSpaceObjectsRound1.Elapsed, "0.000000") > 3.25 Then
            ReleaseIncomingSpaceObject
            ReleaseIncomingSpaceObject
            ReleaseIncomingSpaceObject
            GenerateIncSpaceObjectsRound1.Restart
        End If
        
        If Format(SleepWatch.Elapsed, "0.000000") < Interval Then
            Sleep Interval - Format(SleepWatch.Elapsed, "0.000000")
            SleepWatch.Restart
        End If
        
        DoEvents
    Loop
End Sub

Public Sub HandleSendKeys(ByVal KeyCode As Long)
    Select Case KeyCode
        Case 37
            MoveSpaceObjects.MoveShip Left
        Case 39
            MoveSpaceObjects.MoveShip Right
        Case 32
            LaunchMissile
    End Select
End Sub

Private Function InitializePlayerShip()
Dim PlayerShip As ISpaceObject
    Set PlayerShip = SpaceObjectFactory.NewSpaceObject(Ship)
    GamePiecesCollection.Add PlayerShip.Name, PlayerShip
End Function

Private Function LaunchMissile()
Dim LaunchedMissile As ISpaceObject
    CountMissiles.IncrementMissileCount
    Set LaunchedMissile = SpaceObjectFactory.NewSpaceObject(Missile)
    GamePiecesCollection.Add LaunchedMissile.Name, LaunchedMissile
End Function

Private Function ReleaseIncomingSpaceObject()
Dim IncomingSpaceObject As ISpaceObject
    CountIncomingSpaceObjects.IncrementCountIncomingSpaceObjects
    Set IncomingSpaceObject = SpaceObjectFactory.NewSpaceObject(Application.WorksheetFunction.RandBetween(1, 3))
    GamePiecesCollection.Add IncomingSpaceObject.Name, IncomingSpaceObject
End Function
