Attribute VB_Name = "SOCometFactory"
Option Explicit

Public Function NewComet(ByRef board As GameBoard) As SpaceObjectComet
Dim width                           As Long
Dim height                          As Long

    width = ScaleItems.MaxSize / 1.75
    height = ScaleItems.MaxSize / 1.75
    IncrementSpaceObjectCount
    With New SpaceObjectComet
        .ImgPathWay = "Z:\Desktop Storage\EXCEL & C# PRACTICE\SpaceInvaders\regComet.jpg"
        .SetInitialLeft Application.WorksheetFunction.RandBetween(0, board.width - width)
        .SetInitialTop 0
        .width = width
        .height = height
        .ImageName = "SpaceObject" & CStr(SpaceObjectCount.Count)
        Set NewComet = .Self
    End With
End Function

Private Sub IncrementSpaceObjectCount()
    SpaceObjectCount.IncrementCount
End Sub
