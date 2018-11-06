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

Public CollectionInComingSpaceObjects As New Collection
Public CollectionMissiles As New Collection
Public CollectionShips As New Collection
Public SpaceObjectFactory As ISpaceObjectFactory

Const Interval = 5

    Sub RunGame(ByVal BoardWith As Long, ByVal BoardHeight As Long)
        SetBoardDimensions BoardWith, BoardHeight
     
        Dim SleepWatch As StopWatch
        Set SleepWatch = New StopWatch
        SleepWatch.Start

        Dim GenerateIncSpaceObjectsRound1 As StopWatch
        Set GenerateIncSpaceObjectsRound1 = New StopWatch
        GenerateIncSpaceObjectsRound1.Start

        InitializePlayerShip

        Do
            GameBoard.RefreshGameBoard CombineCollections
            
            MoveSpaceObjects.MoveOutgoingMissiles
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

        Private Sub SetBoardDimensions(ByVal BoardWith As Long, ByVal BoardHeight As Long)
            BoardDimensions.Width = BoardWith
            BoardDimensions.Height = BoardHeight
        End Sub

        Private Sub InitializePlayerShip()
            Set SpaceObjectFactory = SpaceObject.Self
            
            Dim PlayerShip As ISpaceObject
            Set PlayerShip = SpaceObjectFactory.Create(Ship)
            CollectionShips.Add PlayerShip
        End Sub
        
        Private Sub ReleaseIncomingSpaceObject()
            Set SpaceObjectFactory = SpaceObject.Self
            
            Dim IncomingSpaceObject As ISpaceObject
            Set IncomingSpaceObject = SpaceObjectFactory.Create(DetermineTypeOfIncomingSpaceObject)
            CollectionInComingSpaceObjects.Add IncomingSpaceObject
        End Sub

        Private Function DetermineTypeOfIncomingSpaceObject() As SpaceObjectType
            DetermineTypeOfIncomingSpaceObject = Application.WorksheetFunction.RandBetween(1, 3)
        End Function

        Private Function CombineCollections() As Collection
            Set CombineCollections = New Collection
            
            Dim ISpaceObjectIndex As Long
            For ISpaceObjectIndex = 1 To CollectionInComingSpaceObjects.Count
                CombineCollections.Add CollectionInComingSpaceObjects.Item(ISpaceObjectIndex)
            Next ISpaceObjectIndex
            
            For ISpaceObjectIndex = 1 To CollectionMissiles.Count
                CombineCollections.Add CollectionMissiles.Item(ISpaceObjectIndex)
            Next ISpaceObjectIndex
            
            For ISpaceObjectIndex = 1 To CollectionShips.Count
                CombineCollections.Add CollectionShips.Item(ISpaceObjectIndex)
            Next ISpaceObjectIndex
        End Function

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

        Private Sub LaunchMissile()
            Set SpaceObjectFactory = SpaceObject.Self
            
            Dim LaunchedMissile As ISpaceObject
            Set LaunchedMissile = SpaceObjectFactory.Create(Missile)
            CollectionMissiles.Add LaunchedMissile
        End Sub
