Attribute VB_Name = "MoveSpaceObjects"
Option Explicit

Public Enum Direction
    Left = 0
    Right = 1
End Enum

Sub MoveIncomingSpaceObjectsAndMissiles()
Dim SpaceObjectIndex                                As Variant

    For Each SpaceObjectIndex In GamePiecesCollection.Keys()
        If CheckObjectType.IsMissile(GamePiecesCollection.Item(SpaceObjectIndex)) = True Then
            If GamePiecesCollection.Item(SpaceObjectIndex).Top - 1 <= 0 Then
                GamePiecesCollection.remove SpaceObjectIndex
            Else
                GamePiecesCollection.Item(SpaceObjectIndex).Top = GamePiecesCollection.Item(SpaceObjectIndex).Top - 1
            End If
        ElseIf CheckObjectType.IsIncomingSpaceObject(GamePiecesCollection.Item(SpaceObjectIndex)) = True Then
            If GamePiecesCollection.Item(SpaceObjectIndex).Top + 1 >= BoardDimensions.Height Then
                GamePiecesCollection.remove SpaceObjectIndex
            Else
                GamePiecesCollection.Item(SpaceObjectIndex).Top = GamePiecesCollection.Item(SpaceObjectIndex).Top + 1
            End If
        End If
    Next SpaceObjectIndex
End Sub

Sub MoveShip(ByVal MoveShipDirection As Direction)
    Select Case MoveShipDirection
        Case Direction.Left
            If GamePiecesCollection.Item("SHIP").Left - 4 >= 0 Then
                GamePiecesCollection.Item("SHIP").Left = GamePiecesCollection.Item("SHIP").Left - 5
            Else
                GamePiecesCollection.Item("SHIP").Left = 0
            End If
        Case Direction.Right
            If (GamePiecesCollection.Item("SHIP").Left + GamePiecesCollection.Item("SHIP").Width) < BoardDimensions.Width Then
                GamePiecesCollection.Item("SHIP").Left = GamePiecesCollection.Item("SHIP").Left + 4
            Else
                GamePiecesCollection.Item("SHIP").Left = BoardDimensions.Width - GamePiecesCollection.Item("SHIP").Width
            End If
    End Select
End Sub
