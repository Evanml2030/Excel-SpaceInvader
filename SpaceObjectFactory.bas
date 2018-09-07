Attribute VB_Name = "SpaceObjectFactory"
Option Explicit

Public Function NewShip(ByVal board As GameBoard) As SpaceObjectShip
Dim width                           As Long
Dim height                          As Long

    width = 15
    height = 30
    With New SpaceObjectShip
        .ImgPathWay = "C:\Users\evanm\OneDrive\Desktop\Excel\SpaceInvader\SpaceShip.jpg"
        .SetInitialLeft board.width / 2
        .SetInitialTop board.height - (board.height / 8.5)
        .height = height
        .width = width
        .ImageName = "Ship"
        Set NewShip = .Self
    End With
End Function

Public Function NewMissile(ByVal board As GameBoard) As SpaceObjectMissile
Dim width                           As Long
Dim height                          As Long

    width = 15
    height = 30
    IncrementMissileCount
    With New SpaceObjectMissile
        .ImgPathWay = "C:\Users\evanm\OneDrive\Desktop\Excel\SpaceInvader\Missle.jpg"
        .SetInitialLeft ((CollectionShips.Item(1).spaceObject.width - width) / 2) + CollectionShips.Item(1).spaceObject.left
        .SetInitialTop CollectionShips.Item(1).spaceObject.top - height
        .height = height
        .width = width
        .ImageName = "Missile" & CStr(MissileCount.Count)
        Set NewMissile = .Self
    End With
End Function

Private Sub IncrementMissileCount()
    MissileCount.IncrementMissileCount
End Sub

Public Function NewAlien(ByRef board As GameBoard) As SpaceObjectAlien
Dim width                           As Long
Dim height                          As Long

    width = 20
    height = 20
    IncrementIncSpaceObjectCount
    With New SpaceObjectAlien
        .ImgPathWay = "C:\Users\evanm\OneDrive\Desktop\Excel\SpaceInvader\AlienShip.jpg"
        .SetInitialLeft Application.WorksheetFunction.RandBetween(0, board.InsideWidth - width)
        .SetInitialTop 0
        .height = height
        .width = width
        .ImageName = "SpaceObject" & CStr(SpaceObjectCount.Count)
        Set NewAlien = .Self
    End With
End Function

Public Function NewComet(ByRef board As GameBoard) As SpaceObjectComet
Dim width                           As Long
Dim height                          As Long

    width = 30
    height = 30
    IncrementIncSpaceObjectCount
    With New SpaceObjectComet
        .ImgPathWay = "C:\Users\evanm\OneDrive\Desktop\Excel\SpaceInvader\Comet.jpg"
        .SetInitialLeft Application.WorksheetFunction.RandBetween(0, board.InsideWidth - width)
        .SetInitialTop 0
        .width = width
        .height = height
        .ImageName = "SpaceObject" & CStr(SpaceObjectCount.Count)
        Set NewComet = .Self
    End With
End Function

Public Function NewStar(ByRef board As GameBoard) As SpaceObjectStar
Dim width                           As Long
Dim height                          As Long

    width = 40
    height = 40
    IncrementIncSpaceObjectCount
    With New SpaceObjectStar
        .ImgPathWay = "C:\Users\evanm\OneDrive\Desktop\Excel\SpaceInvader\Star.jpg"
        .SetInitialLeft Application.WorksheetFunction.RandBetween(0, board.InsideWidth - width)
        .SetInitialTop 0
        .width = width
        .height = height
        .ImageName = "SpaceObject" & CStr(SpaceObjectCount.Count)
        Set NewStar = .Self
    End With
End Function

Private Sub IncrementIncSpaceObjectCount()
    SpaceObjectCount.IncrementCount
End Sub
