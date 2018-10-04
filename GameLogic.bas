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
Const Interval = 3

Sub RunGame(ByVal BoardWith As Long, ByVal BoardHeight As Long)
    Dim SleepWatch                                                          As StopWatch
    Dim GenerateIncSpaceObjectsRound1                                       As StopWatch
    
    BoardDimensions.Width = BoardWith
    BoardDimensions.Height = BoardHeight
    
    Set SleepWatch = New StopWatch
    SleepWatch.Start
    
    Set GenerateIncSpaceObjectsRound1 = New StopWatch
    GenerateIncSpaceObjectsRound1.Start

    InitializePlayerShip

    Do
        GameBoard.RefreshGameBoard CombineCollections
        
        MoveSpaceObjects.MoveIncomingMissiles
        MoveSpaceObjects.MoveIncomingSpaceObjects
        
        Collisions.HandleMissileCollisions
        If Collisions.HandleShipCollisions Then Exit Do
        
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
    
    GameBoard.CloseGame
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
    Dim PlayerShip                                                          As ISpaceObject
    Set PlayerShip = SpaceObjectFactory.NewSpaceObject(Ship)
    CollectionShips.Add PlayerShip
End Function

Private Function LaunchMissile()
    Dim LaunchedMissile                                                     As ISpaceObject
    Set LaunchedMissile = SpaceObjectFactory.NewSpaceObject(Missile)
    CollectionMissiles.Add LaunchedMissile
End Function

Private Function ReleaseIncomingSpaceObject()
    Dim IncomingSpaceObject                                                 As ISpaceObject
    Set IncomingSpaceObject = SpaceObjectFactory.NewSpaceObject(Application.WorksheetFunction.RandBetween(1, 3))
    CollectionIncomingSpaceObjects.Add IncomingSpaceObject
End Function

Private Function CombineCollections() As Collection
    Dim ISpaceObjectIndex                                                   As Long

    Set CombineCollections = New Collection

    For ISpaceObjectIndex = 1 To CollectionIncomingSpaceObjects.Count
        CombineCollections.Add CollectionIncomingSpaceObjects.Item(ISpaceObjectIndex)
    Next ISpaceObjectIndex
    
    For ISpaceObjectIndex = 1 To CollectionMissiles.Count
        CombineCollections.Add CollectionMissiles.Item(ISpaceObjectIndex)
    Next ISpaceObjectIndex
    
    For ISpaceObjectIndex = 1 To CollectionShips.Count
        CombineCollections.Add CollectionShips.Item(ISpaceObjectIndex)
    Next ISpaceObjectIndex
End Function

