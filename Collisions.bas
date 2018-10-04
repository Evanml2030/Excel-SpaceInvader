Attribute VB_Name = "Collisions"
Option Explicit

Sub HandleMissileCollisions()
    Dim MissileObject                                                       As ISpaceObject
    Dim IncomingSpaceObject                                                 As ISpaceObject
    Dim MissileObjectsIndex                                                 As Long
    Dim IncomingSpaceObjectIndex                                            As Long

    For MissileObjectsIndex = CollectionMissiles.Count To 1 Step -1
        Set MissileObject = CollectionMissiles.Item(MissileObjectsIndex)
        For IncomingSpaceObjectIndex = CollectionIncomingSpaceObjects.Count To 1 Step -1
            Set IncomingSpaceObject = CollectionIncomingSpaceObjects.Item(IncomingSpaceObjectIndex)
            If CheckIfCollided(MissileObject, IncomingSpaceObject) Then
                CollectionMissiles.Remove MissileObjectsIndex
                CollectionIncomingSpaceObjects.Remove IncomingSpaceObjectIndex
                Exit For
            End If
        Next IncomingSpaceObjectIndex
    Next MissileObjectsIndex
End Sub

Function HandleShipCollisions() As Boolean
    Dim ShipObject                                                          As ISpaceObject
    Dim IncomingSpaceObject                                                 As ISpaceObject
    Dim ShipObjectIndex                                                     As Long
    Dim IncomingSpaceObjectIndex                                            As Long

    For ShipObjectIndex = CollectionShips.Count To 1 Step -1
        Set ShipObject = CollectionShips.Item(ShipObjectIndex)
        For IncomingSpaceObjectIndex = CollectionIncomingSpaceObjects.Count To 1 Step -1
            Set IncomingSpaceObject = CollectionIncomingSpaceObjects.Item(IncomingSpaceObjectIndex)
            If CheckIfCollided(ShipObject, IncomingSpaceObject) Then
                HandleShipCollisions = True
                Exit For
            End If
        Next IncomingSpaceObjectIndex
    Next ShipObjectIndex
End Function

Private Function CheckIfCollided(ByVal First As ISpaceObject, ByVal Second As ISpaceObject) As Boolean
    Dim HorizontalOverlap                                                   As Boolean
    Dim VerticalOverlap                                                     As Boolean

    HorizontalOverlap = (First.Left - Second.Width < Second.Left) And (Second.Left < First.Left + First.Width)
    VerticalOverlap = (First.Top - Second.Height < Second.Top) And (Second.Top < First.Top + First.Height)
    CheckIfCollided = HorizontalOverlap And VerticalOverlap
End Function

