Attribute VB_Name = "MMissleFactory"
Option Explicit

Public Function NewMissle() As missle
Dim width                           As Long
Dim height                          As Long

    width = ScaleItems.MaxSize / 2
    height = ScaleItems.MaxSize / 2.15
    IncrementMissleCount
    With New missle
        .ImgPathWay = "Z:\Desktop Storage\EXCEL & C# PRACTICE\SpaceInvaders\laserBeam.jpg"
        .SetInitialLeft ((Ship.width - width) / 2) + Ship.left
        .SetInitialTop Ship.top - height
        .height = height
        .width = width
        .ImageName = "Missle" & CStr(MissleCount.Count)
        Set NewMissle = .Self
    End With
End Function

Private Sub IncrementMissleCount()
    MissleCount.IncrementMissleCount
End Sub
