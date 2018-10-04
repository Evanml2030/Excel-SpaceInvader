Attribute VB_Name = "SpaceObjectFactory"
Option Explicit

Public Function NewSpaceObject(ByVal SpaceObjectType As SpaceObjectType) As Spaceobject
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

Private Function NewSpaceObjectAlien() As Spaceobject
    With New Spaceobject
        .SetInitialLeft Application.WorksheetFunction.RandBetween(1, BoardDimensions.Width)
        .SetInitialTop 1
        .Height = BoardDimensions.Width / 10
        .Width = BoardDimensions.Width / 10
        .SpaceObjectType = Alien
        .Name = "INCSPACEOBJECT" & CollectionIncomingSpaceObjects.Count
        Set NewSpaceObjectAlien = .Self
    End With
End Function

Private Function NewSpaceObjectComet() As Spaceobject
    With New Spaceobject
        .SetInitialLeft Application.WorksheetFunction.RandBetween(1, BoardDimensions.Width)
        .SetInitialTop 1
        .Height = BoardDimensions.Height / 7
        .Width = BoardDimensions.Height / 7
        .SpaceObjectType = Comet
        .Name = "INCSPACEOBJECT" & CollectionIncomingSpaceObjects.Count
        Set NewSpaceObjectComet = .Self
    End With
End Function

Private Function NewSpaceObjectMissile() As Spaceobject
    With New Spaceobject
        .SetInitialLeft ((CollectionShips.Item(1).Width - (BoardDimensions.Width / 20)) / 2) + CollectionShips.Item(1).Left
        .SetInitialTop CollectionShips.Item(1).Top - BoardDimensions.Height / 15
        .Height = BoardDimensions.Height / 15
        .Width = BoardDimensions.Height / 20
        .SpaceObjectType = Missile
        .Name = "MISSILE" & CollectionMissiles.Count
        Set NewSpaceObjectMissile = .Self
    End With
End Function

Private Function NewSpaceObjectShip() As Spaceobject
    With New Spaceobject
        .SetInitialLeft BoardDimensions.Width / 2 - ((BoardDimensions.Height / 7) / 2)
        .SetInitialTop Round(BoardDimensions.Height - ((BoardDimensions.Height / 7) * 1.25), 0)
        .Height = BoardDimensions.Height / 7
        .Width = BoardDimensions.Width / 7
        .SpaceObjectType = Ship
        .Name = "SHIP"
        Set NewSpaceObjectShip = .Self
    End With
End Function

Private Function NewSpaceObjectStar() As Spaceobject
    With New Spaceobject
        .SetInitialLeft Application.WorksheetFunction.RandBetween(1, BoardDimensions.Width)
        .SetInitialTop 1
        .Height = BoardDimensions.Height / 5
        .Width = BoardDimensions.Height / 5
        .SpaceObjectType = Star
        .Name = "INCSPACEOBJECT" & CollectionIncomingSpaceObjects.Count
        Set NewSpaceObjectStar = .Self
    End With
End Function

