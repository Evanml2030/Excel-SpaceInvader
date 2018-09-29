Attribute VB_Name = "SpaceObjectFactory"
Option Explicit

Public Function NewSpaceObject(ByVal SpaceObjectType As SpaceObjectType) As SpaceObject
    Select Case SpaceObjectType
        Case Alien
            Set NewSpaceObject = NewSpaceObjectAlien
        Case Comet
            Set NewSpaceObject = NewSpaceObjectComet
        Case Missile
            Set NewSpaceObject = NewSpaceObjectMissile
        Case Ship
            Set NewSpaceObject = NewSpaceObjectShip
        Case Star
            Set NewSpaceObject = NewSpaceObjectStar
    End Select
End Function

Private Function NewSpaceObjectAlien() As SpaceObject
    CountIncomingSpaceObjects.IncrementCountIncomingSpaceObjects
    With New SpaceObject
        .SetInitialLeft Application.WorksheetFunction.RandBetween(1, BoardDimensions.Width)
        .SetInitialTop 1
        .Height = BoardDimensions.Width / 10
        .Width = BoardDimensions.Width / 10
        .SpaceObjectType = Alien
        .Name = "INCSPACEOBJECT" & CountIncomingSpaceObjects.Count
        Set NewSpaceObjectAlien = .Self
    End With
End Function

Private Function NewSpaceObjectComet() As SpaceObject
    CountIncomingSpaceObjects.IncrementCountIncomingSpaceObjects
    With New SpaceObject
        .SetInitialLeft Application.WorksheetFunction.RandBetween(1, BoardDimensions.Width)
        .SetInitialTop 1
        .Height = BoardDimensions.Height / 7
        .Width = BoardDimensions.Height / 7
        .SpaceObjectType = Comet
        .Name = "INCSPACEOBJECT" & CountIncomingSpaceObjects.Count
        Set NewSpaceObjectComet = .Self
    End With
End Function

Private Function NewSpaceObjectMissile() As SpaceObject
    CountMissiles.IncrementMissileCount
    With New SpaceObject
        .SetInitialLeft ((GamePiecesCollection.Item("SHIP").Width - (BoardDimensions.Width / 20)) / 2) + GamePiecesCollection.Item("SHIP").Left
        .SetInitialTop GamePiecesCollection.Item("SHIP").Top - BoardDimensions.Height / 15
        .Height = BoardDimensions.Height / 15
        .Width = BoardDimensions.Width / 20
        .SpaceObjectType = Missile
        .Name = "MISSILE" & CountMissiles.Count
        Set NewSpaceObjectMissile = .Self
    End With
End Function

Private Function NewSpaceObjectShip() As SpaceObject
    With New SpaceObject
        .SetInitialLeft BoardDimensions.Width / 2 - ((BoardDimensions.Height / 7) / 2)
        .SetInitialTop BoardDimensions.Height - ((BoardDimensions.Height / 7) * 1.25)
        .Height = BoardDimensions.Height / 7
        .Width = BoardDimensions.Width / 7
        .SpaceObjectType = Ship
        .Name = "SHIP"
        Set NewSpaceObjectShip = .Self
    End With
End Function

Private Function NewSpaceObjectStar() As SpaceObject
    CountIncomingSpaceObjects.IncrementCountIncomingSpaceObjects
    With New SpaceObject
        .SetInitialLeft Application.WorksheetFunction.RandBetween(1, BoardDimensions.Width)
        .SetInitialTop 1
        .Height = BoardDimensions.Height / 5
        .Width = BoardDimensions.Height / 5
        .SpaceObjectType = Star
        .Name = "INCSPACEOBJECT" & CountIncomingSpaceObjects.Count
        Set NewSpaceObjectStar = .Self
    End With
End Function
