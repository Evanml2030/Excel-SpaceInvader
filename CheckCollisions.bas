Attribute VB_Name = "CheckCollisions"
Sub HandleMissileIncSpaceObjectCollisions(ByVal board As GameBoard)
Dim MissileIterator                                 As IBoundControl
Dim IncSpaceObjectIterator                          As IBoundControl
Dim MissileController                               As Control
Dim SpaceObjectController                           As Control
Dim x                                               As Long
Dim y                                               As Long
    
    For x = CollectionMissiles.Count To 1 Step -1
    Set MissileIterator = CollectionMissiles.Item(x)
        For y = CollectionIncomingSpaceObjects.Count To 1 Step -1
        Set IncSpaceObjectIterator = CollectionIncomingSpaceObjects.Item(y)
             If CheckIfCollided(MissileIterator, IncSpaceObjectIterator) Then
                DestroyObject.DestroySpaceObject board, MissileIterator
                CollectionMissiles.remove x
                DestroyObject.DestroySpaceObject board, IncSpaceObjectIterator
                CollectionIncomingSpaceObjects.remove y
                Exit For
            End If
        Next y
    Next x
End Sub

Function HandleShipIncSpaceObjectCollisions() As Boolean
Dim Ship                                            As IBoundControl
Dim IncSpaceObjectIterator                          As IBoundControl

Set Ship = CollectionShips.Item(1)

    For Each IncSpaceObjectIterator In CollectionIncomingSpaceObjects
        If CheckIfCollided(Ship, IncSpaceObjectIterator) Then
            HandleShipIncSpaceObjectCollisions = True
            Exit For
        End If
    Next IncSpaceObjectIterator
End Function

Private Function CheckIfCollided(ByVal first As IBoundControl, ByVal second As IBoundControl) As Boolean
Dim hOverlap                                        As Boolean
Dim vOverlap                                        As Boolean

    hOverlap = (first.spaceObject.left - second.spaceObject.width < second.spaceObject.left) And (second.spaceObject.left < first.spaceObject.left + first.spaceObject.width)
    vOverlap = (first.spaceObject.top - second.spaceObject.height < second.spaceObject.top) And (second.spaceObject.top < first.spaceObject.top + first.spaceObject.height)
    CheckIfCollided = hOverlap And vOverlap
End Function
