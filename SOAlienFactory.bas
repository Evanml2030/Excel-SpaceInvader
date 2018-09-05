Attribute VB_Name = "SOAlienFactory"
Option Explicit

Public Function NewAlien(ByRef board As GameBoard) As SpaceObjectAlien
Dim width                           As Long
Dim height                          As Long

    width = ScaleItems.MaxSize / 1.5
    height = ScaleItems.MaxSize / 1.5
    IncrementSpaceObjectCount
    With New SpaceObjectAlien
        .ImgPathWay = "Z:\Desktop Storage\EXCEL & C# PRACTICE\SpaceInvaders\alienShip.jpg"
        .SetInitialLeft Application.WorksheetFunction.RandBetween(0, board.width - width)
        .SetInitialTop 0
        .height = height
        .width = width
        .ImageName = "SpaceObject" & CStr(SpaceObjectCount.Count)
        Set NewAlien = .Self
    End With
End Function

Private Sub IncrementSpaceObjectCount()
    SpaceObjectCount.IncrementCount
End Sub
