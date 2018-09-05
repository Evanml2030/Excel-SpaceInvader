Attribute VB_Name = "MDestroyMissleObject"
Option Explicit

Sub DestroyMissleObject(ByRef board As GameBoard, ByRef missleObject As missle, ByRef index As Long)
    board.Controls.Remove missleObject.ImageName
    MissleObjectsDataCol.Remove index
    MissleCntrlsCol.Remove index
End Sub

