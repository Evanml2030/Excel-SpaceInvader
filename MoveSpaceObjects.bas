Attribute VB_Name = "MoveSpaceObjects"
Option Explicit

Public Enum Direction
    left = 0
    Right = 1
End Enum

Sub MoveIncomingSpaceObjects(ByVal board As GameBoard)
Dim iterator                                As IBoundControl
Dim index                                   As Long

    For index = CollectionIncomingSpaceObjects.Count To 1 Step -1
        Set iterator = CollectionIncomingSpaceObjects.Item(index)
        If iterator.spaceObject.top + 1 >= board.height Then
            DestroyObject.DestroySpaceObject board, iterator
            CollectionIncomingSpaceObjects.remove index
        Else
            iterator.spaceObject.top = iterator.spaceObject.top + 1
            iterator.Control.top = iterator.spaceObject.top
        End If
    Next index
End Sub

Sub MoveMissiles(ByVal board As GameBoard)
Dim iterator                                As IBoundControl
Dim index                                   As Long

    For index = CollectionMissiles.Count To 1 Step -1
        Set iterator = CollectionMissiles.Item(index)
        If iterator.spaceObject.top - 1 <= 0 Then
            DestroyObject.DestroySpaceObject board, iterator
            CollectionMissiles.remove index
        Else
            iterator.spaceObject.top = iterator.spaceObject.top - 1
            iterator.Control.top = iterator.spaceObject.top
        End If
    Next index
End Sub

Sub MoveShip(ByVal val As Direction, ByVal board As GameBoard)
    Select Case val
        Case Direction.left
            If CollectionShips.Item(1).spaceObject.left - 5 >= 0 Then
                CollectionShips.Item(1).spaceObject.left = CollectionShips.Item(1).spaceObject.left - 4
                CollectionShips.Item(1).Control.left = CollectionShips.Item(1).spaceObject.left
            Else
                CollectionShips.Item(1).spaceObject.left = 0
                CollectionShips.Item(1).Control.left = CollectionShips.Item(1).spaceObject.left
            End If
        Case Direction.Right
            If (CollectionShips.Item(1).spaceObject.left + CollectionShips.Item(1).spaceObject.width) < board.InsideWidth Then
                CollectionShips.Item(1).spaceObject.left = CollectionShips.Item(1).spaceObject.left + 4
                CollectionShips.Item(1).Control.left = CollectionShips.Item(1).spaceObject.left
            Else
                CollectionShips.Item(1).spaceObject.left = board.InsideWidth - CollectionShips.Item(1).spaceObject.width
                CollectionShips.Item(1).Control.left = CollectionShips.Item(1).spaceObject.left
            End If
    End Select
End Sub
