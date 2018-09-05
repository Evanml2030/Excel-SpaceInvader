Attribute VB_Name = "MMoveMissles"
Option Explicit

Sub MoveMissleObjects(ByRef board As GameBoard)
Dim missleObject                    As missle
Dim missleObjectCntrl               As Control
Dim index                           As Long

    For index = MissleObjectsDataCol.Count To 1 Step -1
        Set missleObject = MissleObjectsDataCol.Item(index)
        Set missleObjectCntrl = MissleCntrlsCol.Item(index)
        If MissleObjectOutOfBounds(board, missleObject) Then
            MDestroyMissleObject.DestroyMissleObject board, missleObject, index
            Set missleObject = Nothing
            Set missleObjectCntrl = Nothing
        Else
            MoveMissleObject missleObject, missleObjectCntrl
        End If
    Next index
End Sub

Private Function MissleObjectOutOfBounds(ByRef board As GameBoard, ByRef missleObject As missle) As Boolean
    If missleObject.top = 0 Then
        MissleObjectOutOfBounds = True
    Else
        MissleObjectOutOfBounds = False
    End If
End Function

Private Sub MoveMissleObject(ByRef missleObject As missle, ByRef missleObjectCntrl As Control)
    missleObject.top = missleObject.top - 1
    missleObjectCntrl.top = missleObject.top
End Sub