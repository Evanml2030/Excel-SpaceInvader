Attribute VB_Name = "CollisionsMissleSpaceObject"
Option Explicit

Sub HandleMissleSpaceObjectCollisions(ByRef board As GameBoard)
Dim spaceObject                     As ISpaceObject
Dim spaceObjectCntrl                As Control
Dim missle                          As missle
Dim missleCntrl                     As Control
Dim indexMissle                     As Long
Dim indexSpaceObject                As Long

    For indexMissle = MissleObjectsDataCol.Count To 1 Step -1
        Set missle = MissleObjectsDataCol.Item(indexMissle)
        Set missleCntrl = MissleCntrlsCol.Item(indexMissle)
        For indexSpaceObject = SpaceObjectDataCol.Count To 1 Step -1
            Set spaceObject = SpaceObjectDataCol.Item(indexSpaceObject)
            Set spaceObjectCntrl = SpaceObjectCntrlsCol.Item(indexSpaceObject)
            If CheckIfCollided(missle, spaceObject) Then
                MDestroyMissleObject.DestroyMissleObject board, missle, indexMissle
                SODestroySpaceObject.DestroySpaceObject board, spaceObject, indexSpaceObject
            End If
        Next indexSpaceObject
    Next indexMissle
End Sub

Private Function CheckIfCollided(ByRef missle As missle, ByRef spaceObject As ISpaceObject) As Boolean
Dim hOverlap                        As Boolean
Dim vOverlap                        As Boolean

    hOverlap = (missle.left - spaceObject.width < spaceObject.left) And (spaceObject.left < missle.left + missle.width)
    vOverlap = (missle.top - spaceObject.height < spaceObject.top) And (spaceObject.top < missle.top + missle.height)
    CheckIfCollided = hOverlap And vOverlap
End Function
