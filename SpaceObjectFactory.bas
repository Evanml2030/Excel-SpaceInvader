Option Explicit

Public Function NewSpaceObjectShip(ByVal board As GameBoard) As SpaceObjectShip
Dim width                           As Long
Dim height                          As Long

    width = 15
    height = 30
    With New SpaceObjectShip
        .ImgPathWay = ActiveWorkbook.Path & "SpaceShip.jpg"
        .SetInitialLeft board.width / 2
        .SetInitialTop board.height - (board.height / 8.5)
        .height = height
        .width = width
        .ImageName = "Ship"
        Set NewSpaceObjectShip = .Self
    End With
End Function

Public Function NewSpaceObjectMissile(ByVal board As GameBoard) As SpaceObjectMissile
Dim width                           As Long
Dim height                          As Long

    width = 15
    height = 30
    IncrementMissileCount
    With New SpaceObjectMissile
        .ImgPathWay = ActiveWorkbook.Path & "\Missile.jpg"
        .SetInitialLeft ((CollectionShips.Item(1).spaceObject.width - width) / 2) + CollectionShips.Item(1).spaceObject.left
        .SetInitialTop CollectionShips.Item(1).spaceObject.top - height
        .height = height
        .width = width
        .ImageName = "Missile" & CStr(MissileCount.Count)
        Set NewSpaceObjectMissile = .Self
    End With
End Function

Private Sub IncrementMissileCount()
    MissileCount.IncrementMissileCount
End Sub

Public Function NewSpaceObjectAlien(ByRef board As GameBoard) As SpaceObjectAlien
Dim width                           As Long
Dim height                          As Long

    width = 20
    height = 20
    IncrementIncSpaceObjectCount
    With New SpaceObjectAlien
        .ImgPathWay = ActiveWorkbook.Path & "\AlienShip.jpg"
        .SetInitialLeft Application.WorksheetFunction.RandBetween(0, board.InsideWidth - width)
        .SetInitialTop 0
        .height = height
        .width = width
        .ImageName = "SpaceObject" & CStr(SpaceObjectCount.Count)
        Set NewSpaceObjectAlien = .Self
    End With
End Function

Public Function NewSpaceObjectComet(ByRef board As GameBoard) As SpaceObjectComet
Dim width                           As Long
Dim height                          As Long

    width = 20
    height = 20
    IncrementIncSpaceObjectCount
    With New SpaceObjectComet
        .ImgPathWay = ActiveWorkbook.Path & "\Comet.jpg"
        .SetInitialLeft Application.WorksheetFunction.RandBetween(0, board.InsideWidth - width)
        .SetInitialTop 0
        .width = width
        .height = height
        .ImageName = "SpaceObject" & CStr(SpaceObjectCount.Count)
        Set NewSpaceObjectComet = .Self
    End With
End Function

Public Function NewSpaceObjectStar(ByRef board As GameBoard) As SpaceObjectStar
Dim width                           As Long
Dim height                          As Long

    width = 40
    height = 40
    IncrementIncSpaceObjectCount
    With New SpaceObjectStar
        .ImgPathWay = ActiveWorkbook.Path & "\Star.jpg"
        .SetInitialLeft Application.WorksheetFunction.RandBetween(0, board.InsideWidth - width)
        .SetInitialTop 0
        .width = width
        .height = height
        .ImageName = "SpaceObject" & CStr(SpaceObjectCount.Count)
        Set NewSpaceObjectStar = .Self
    End With
End Function

Private Sub IncrementIncSpaceObjectCount()
    SpaceObjectCount.IncrementCount
End Sub
