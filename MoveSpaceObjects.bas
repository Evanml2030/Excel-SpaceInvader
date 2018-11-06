Attribute VB_Name = "MoveSpaceObjects"
Option Explicit

Public Enum Direction
    Left = 0
    Right = 1
End Enum

	Sub MoveOutgoingMissiles()
	    Dim SpaceObjectIndex As Variant

	    For SpaceObjectIndex = CollectionMissiles.Count To 1 Step -1
	        If CollectionMissiles.Item(SpaceObjectIndex).Top - 3 <= 0 Then
	            CollectionMissiles.Remove SpaceObjectIndex
	        Else
	            CollectionMissiles.Item(SpaceObjectIndex).MoveNorth
	        End If
	    Next SpaceObjectIndex
	End Sub

	Sub MoveIncomingSpaceObjects()
	    Dim SpaceObjectIndex As Variant

	    For SpaceObjectIndex = CollectionInComingSpaceObjects.Count To 1 Step -1
	        If CollectionInComingSpaceObjects.Item(SpaceObjectIndex).Top + 3 >= BoardDimensions.Height Then
	            CollectionInComingSpaceObjects.Remove SpaceObjectIndex
	        Else
	            CollectionInComingSpaceObjects.Item(SpaceObjectIndex).MoveSouth
	        End If
	    Next SpaceObjectIndex
	End Sub

	Sub MoveShip(ByVal MoveShipDirection As Direction)
	    Select Case MoveShipDirection
	    Case Direction.Left
	        If CollectionShips.Item(1).Left - 4 >= 0 Then
	            CollectionShips.Item(1).MoveLeft
	        End If
	    Case Direction.Right
	        If (CollectionShips.Item(1).Left + CollectionShips.Item(1).Width) < BoardDimensions.Width Then
	            CollectionShips.Item(1).MoveRight
	        End If
	    End Select
	End Sub

