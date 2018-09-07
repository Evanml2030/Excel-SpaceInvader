Attribute VB_Name = "MoveSpaceObjects"
Option Explicit

Public Enum Direction
    left = 0
    Right = 1
End Enum

Sub MoveIncomingSpaceObjects()
Dim iterator                                As iboundcontrol
    For Each iterator In CollectionIncomingSpaceObjects
        iterator.spaceObject.top = iterator.spaceObject.top + 1
        iterator.Control.top = iterator.spaceObject.top
    Next iterator
End Sub

Sub MoveMissiles()
Dim iterator As iboundcontrol
    For Each iterator In CollectionMissiles
        iterator.spaceObject.top = iterator.spaceObject.top - 1
        iterator.Control.top = iterator.spaceObject.top
    Next iterator
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
