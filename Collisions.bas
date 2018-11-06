Attribute VB_Name = "Collisions"
Option Explicit

	Sub HandleMissileCollisions()
	    Dim MissileObjectsIndex As Long
	    Dim MissileObject As ISpaceObject
	    Dim IncomingSpaceObjectIndex As Long
	    Dim IncomingSpaceObject As ISpaceObject

	    For MissileObjectsIndex = CollectionMissiles.Count To 1 Step -1
	        Set MissileObject = CollectionMissiles.Item(MissileObjectsIndex)
	        For IncomingSpaceObjectIndex = CollectionInComingSpaceObjects.Count To 1 Step -1
	            Set IncomingSpaceObject = CollectionInComingSpaceObjects.Item(IncomingSpaceObjectIndex)
	            If CheckIfCollided(MissileObject, IncomingSpaceObject) Then
	                CollectionMissiles.Remove MissileObjectsIndex
	                CollectionInComingSpaceObjects.Remove IncomingSpaceObjectIndex
	                Exit For
	            End If
	        Next IncomingSpaceObjectIndex
	    Next MissileObjectsIndex
	End Sub

	Public Function HandleShipCollisions() As Boolean
	    Dim ShipObjectIndex As Long
	    Dim ShipObject As ISpaceObject
	    Dim IncomingSpaceObjectIndex As Long
	    Dim IncomingSpaceObject As ISpaceObject

	    For ShipObjectIndex = CollectionShips.Count To 1 Step -1
	        Set ShipObject = CollectionShips.Item(ShipObjectIndex)
	        For IncomingSpaceObjectIndex = CollectionInComingSpaceObjects.Count To 1 Step -1
	            Set IncomingSpaceObject = CollectionInComingSpaceObjects.Item(IncomingSpaceObjectIndex)
	            If CheckIfCollided(ShipObject, IncomingSpaceObject) Then
	                HandleShipCollisions = True
	                Exit For
	            End If
	        Next IncomingSpaceObjectIndex
	    Next ShipObjectIndex
	End Function

		Private Function CheckIfCollided(ByVal First As ISpaceObject, ByVal Second As ISpaceObject) As Boolean
		    Dim HorizontalOverlap As Boolean
		    Dim VerticalOverlap As Boolean

		    HorizontalOverlap = (First.Left - Second.Width < Second.Left) And (Second.Left < First.Left + First.Width)
		    VerticalOverlap = (First.Top - Second.Height < Second.Top) And (Second.Top < First.Top + First.Height)
		    CheckIfCollided = HorizontalOverlap And VerticalOverlap
		End Function
