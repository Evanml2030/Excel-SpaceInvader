Attribute VB_Name = "MoveSpaceObjects"
Option Explicit

Public Enum Direction
    Left = 0
    Right = 1
End Enum

Sub MoveIncomingMissiles()
    Dim SpaceObjectIndex                                                    As Variant

    For SpaceObjectIndex = CollectionMissiles.Count To 1 Step -1
        If CollectionMissiles.Item(SpaceObjectIndex).Top - 3 <= 0 Then
            CollectionMissiles.Remove SpaceObjectIndex
        Else
            CollectionMissiles.Item(SpaceObjectIndex).Top = CollectionMissiles.Item(SpaceObjectIndex).Top - 3
        End If
    Next SpaceObjectIndex
End Sub

Sub MoveIncomingSpaceObjects()
    Dim SpaceObjectIndex                                                    As Variant

    For SpaceObjectIndex = CollectionIncomingSpaceObjects.Count To 1 Step -1
        If CollectionIncomingSpaceObjects.Item(SpaceObjectIndex).Top + 3 >= BoardDimensions.Height Then
            CollectionIncomingSpaceObjects.Remove SpaceObjectIndex
        Else
            CollectionIncomingSpaceObjects.Item(SpaceObjectIndex).Top = CollectionIncomingSpaceObjects.Item(SpaceObjectIndex).Top + 3
        End If
    Next SpaceObjectIndex
End Sub

Sub MoveShip(ByVal MoveShipDirection As Direction)
    Select Case MoveShipDirection
    Case Direction.Left
        If CollectionShips.Item(1).Left - 4 >= 0 Then
            CollectionShips.Item(1).Left = CollectionShips.Item(1).Left - 5
        Else
            CollectionShips.Item(1).Left = 0
        End If
    Case Direction.Right
        If (CollectionShips.Item(1).Left + CollectionShips.Item(1).Width) < BoardDimensions.Width Then
            CollectionShips.Item(1).Left = CollectionShips.Item(1).Left + 4
        Else
            CollectionShips.Item(1).Left = BoardDimensions.Width - CollectionShips.Item(1).Width
        End If
    End Select
End Sub

