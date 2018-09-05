Attribute VB_Name = "SOStarFactory"
Option Explicit

Public Function NewStar(ByRef board As GameBoard) As SpaceObjectStar
Dim width                           As Long
Dim height                          As Long

    width = ScaleItems.MaxSize
    height = ScaleItems.MaxSize
    IncrementSpaceObjectCount
    With New SpaceObjectStar
        .ImgPathWay = "Z:\Desktop Storage\EXCEL & C# PRACTICE\SpaceInvaders\yellowStar.jpg"
        .SetInitialLeft Application.WorksheetFunction.RandBetween(0, board.width - width)
        .SetInitialTop 0
        .width = width
        .height = height
        .ImageName = "SpaceObject" & CStr(SpaceObjectCount.Count)
        Set NewStar = .Self
    End With
End Function

Private Sub IncrementSpaceObjectCount()
    SpaceObjectCount.IncrementCount
End Sub

