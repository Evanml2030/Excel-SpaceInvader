Attribute VB_Name = "CheckCollisions"
Sub HandleMissileIncSpaceObjectCollisions(ByVal board As GameBoard)
Dim MissileIterator                                 As iboundcontrol
Dim IncSpaceObjectIterator                          As iboundcontrol
Dim toRemoveMissile                                 As Collection
Dim toRemoveSpaceObject                             As Collection

    Set toRemoveMissile = New Collection
    Set toRemoveSpaceObject = New Collection
    
    For Each MissileIterator In CollectionMissiles
        For Each IncSpaceObjectIterator In CollectionIncomingSpaceObjects
             If CheckIfCollided(MissileIterator, IncSpaceObjectIterator) Then
                toRemoveMissile.Add MissileIterator.spaceObject.ImageName
                toRemoveSpaceObject.Add IncSpaceObjectIterator.spaceObject.ImageName
            End If
        Next IncSpaceObjectIterator
    Next MissileIterator
    
    Dim x As Long
    Dim y As Long
    For x = 1 To toRemoveMissile.Count
        For y = 1 To CollectionMissiles.Count
            If toRemoveMissile.Item(x) = CollectionMissiles.Item(y).spaceObject.ImageName Then
                DestroySpaceObject CollectionMissiles.Item(y), board
                CollectionMissiles.remove y
                Exit For
            End If
        Next y
    Next x
     
    For x = 1 To toRemoveSpaceObject.Count
        For y = 1 To CollectionIncomingSpaceObjects.Count
            If toRemoveSpaceObject.Item(x) = CollectionIncomingSpaceObjects.Item(y).spaceObject.ImageName Then
                DestroySpaceObject CollectionIncomingSpaceObjects.Item(y), board
                CollectionIncomingSpaceObjects.remove y
                Exit For
            End If
        Next y
    Next x
End Sub

Function HandleShipIncSpaceObjectCollisions() As Boolean
Dim Ship                                            As iboundcontrol
Dim IncSpaceObjectIterator                          As iboundcontrol

Set Ship = CollectionShips.Item(1)

    For Each IncSpaceObjectIterator In CollectionIncomingSpaceObjects
        If CheckIfCollided(Ship, IncSpaceObjectIterator) Then
            HandleShipIncSpaceObjectCollisions = True
            Exit For
        End If
    Next IncSpaceObjectIterator
End Function

Private Function CheckIfCollided(ByVal first As iboundcontrol, ByVal second As iboundcontrol) As Boolean
Dim hOverlap                                        As Boolean
Dim vOverlap                                        As Boolean

    hOverlap = (first.spaceObject.left - second.spaceObject.width < second.spaceObject.left) And (second.spaceObject.left < first.spaceObject.left + first.spaceObject.width)
    vOverlap = (first.spaceObject.top - second.spaceObject.height < second.spaceObject.top) And (second.spaceObject.top < first.spaceObject.top + first.spaceObject.height)
    CheckIfCollided = hOverlap And vOverlap
End Function

Private Sub DestroySpaceObject(ByRef objectToDestroy As iboundcontrol, ByVal board As GameBoard)
    board.Controls.remove objectToDestroy.spaceObject.ImageName
End Sub

