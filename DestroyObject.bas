Attribute VB_Name = "DestroyObject"
Public Sub DestroySpaceObject(ByVal board As GameBoard, ByRef objectToDestroy As IBoundControl)
    board.Controls.remove objectToDestroy.spaceObject.ImageName
End Sub

