Attribute VB_Name = "SODestroySpaceObject"
Option Explicit

Sub DestroySpaceObject(ByRef board As GameBoard, ByRef spaceObject As ISpaceObject, ByRef index As Long)
    board.Controls.Remove spaceObject.ImageName
    SpaceObjectDataCol.Remove index
    SpaceObjectCntrlsCol.Remove index
End Sub
