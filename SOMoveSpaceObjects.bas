Attribute VB_Name = "SOMoveSpaceObjects"
Option Explicit

Sub MoveSpaceObjects(ByRef board As GameBoard)
Dim spaceObject                     As ISpaceObject
Dim spaceObjectCntrl                As Control
Dim index                           As Long

For index = SpaceObjectDataCol.Count To 1 Step -1
    Set spaceObject = SpaceObjectDataCol.Item(index)
    Set spaceObjectCntrl = SpaceObjectCntrlsCol.Item(index)
    If SpaceObjectOutOfBounds(board, spaceObject) Then
        SODestroySpaceObject.DestroySpaceObject board, spaceObject, index
        Set spaceObject = Nothing
        Set spaceObjectCntrl = Nothing
    Else
        MoveSpaceObject spaceObject, spaceObjectCntrl
    End If
Next index
End Sub

Private Function SpaceObjectOutOfBounds(ByRef board As GameBoard, ByRef spaceObject As ISpaceObject) As Boolean
    If spaceObject.top + spaceObject.height > board.height Then
        SpaceObjectOutOfBounds = True
    Else
        SpaceObjectOutOfBounds = False
    End If
End Function

Private Sub MoveSpaceObject(ByRef spaceObject As ISpaceObject, ByRef spaceObjectCntrl As Control)
    spaceObject.top = spaceObject.top + 1
    spaceObjectCntrl.top = spaceObject.top
End Sub

