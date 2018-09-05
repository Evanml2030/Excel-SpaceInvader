Attribute VB_Name = "CollisionsShipSpaceObject"
Option Explicit

Function HandleShipSpaceObjectCollisions(ByRef board As GameBoard) As Boolean
Dim spaceObject                     As ISpaceObject
Dim spaceObjectCntrl                As Control
Dim indexSpaceObject                As Long

    For indexSpaceObject = SpaceObjectDataCol.Count To 1 Step -1
        Set spaceObject = SpaceObjectDataCol.Item(indexSpaceObject)
        Set spaceObjectCntrl = SpaceObjectCntrlsCol.Item(indexSpaceObject)
        If CheckIfCollided(spaceObject) Then
            HandleShipSpaceObjectCollisions = True
        End If
    Next indexSpaceObject
End Function

Private Function CheckIfCollided(ByRef spaceObject As ISpaceObject) As Boolean
Dim hOverlap                        As Boolean
Dim vOverlap                        As Boolean

    hOverlap = (Ship.left - spaceObject.width < spaceObject.left) And (spaceObject.left < Ship.left + Ship.width)
    vOverlap = (Ship.top - spaceObject.height < spaceObject.top) And (spaceObject.top < Ship.top + Ship.height)
    CheckIfCollided = hOverlap And vOverlap
End Function
