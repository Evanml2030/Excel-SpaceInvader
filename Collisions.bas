Option Explicit

Function HandleMissileCollisions() As Dictionary
    Dim TempDict                                        As Dictionary
    Dim MissileKey                                      As Variant
    Dim IncomingSpaceObjectKey                          As Variant

    Set TempDict = GamePiecesCollection

    For Each MissileKey In GamePiecesCollection.Keys()
        If CheckObjectType.IsMissile(GamePiecesCollection.Item(MissileKey)) = True Then
            For Each IncomingSpaceObjectKey In GamePiecesCollection.Keys()
                If CheckObjectType.IsIncomingSpaceObject(GamePiecesCollection.Item(IncomingSpaceObjectKey)) And (IncomingSpaceObjectKey <> MissileKey) = True Then
                    If CheckIfCollided(GamePiecesCollection.Item(MissileKey), GamePiecesCollection.Item(IncomingSpaceObjectKey)) Then
                        TempDict.remove MissileKey
                        TempDict.remove IncomingSpaceObjectKey
                    End If
                End If
            Next IncomingSpaceObjectKey
        End If
    Next MissileKey
    Set GamePiecesCollection = TempDict
End Function

Function HandleShipCollisions() As PlayerShipHit
    Dim Ship                                            As ISpaceObject
    Dim IncomingSpaceObjectKey                          As Variant

    Set Ship = GamePiecesCollection.Items(0)

    For Each IncomingSpaceObjectKey In GamePiecesCollection.Keys()
        If CheckObjectType.IsIncomingSpaceObject(GamePiecesCollection.Item(IncomingSpaceObjectKey)) = True Then
            If CheckIfCollided(Ship, GamePiecesCollection(IncomingSpaceObjectKey)) Then
                HandleShipCollisions = Hit
                Exit For
            End If
        End If
    Next IncomingSpaceObjectKey
End Function

Private Function CheckIfCollided(ByVal First As ISpaceObject, ByVal Second As ISpaceObject) As Boolean
    Dim HorizontalOverlap                               As Boolean
    Dim VerticalOverlap                                 As Boolean

    HorizontalOverlap = (First.Left - Second.Width < Second.Left) And (Second.Left < First.Left + First.Width)
    VerticalOverlap = (First.Top - Second.Height < Second.Top) And (Second.Top < First.Top + First.Height)
    CheckIfCollided = HorizontalOverlap And VerticalOverlap
End Function


