VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISpaceObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

	Public Property Get Left() As Long
	End Property

	Public Property Get Top() As Long
	End Property

	Public Property Get Width() As Long
	End Property

	Public Property Get Height() As Long
	End Property

	Public Property Get Name() As String
	End Property

	Public Property Get SpaceObjectType() As SpaceObjectType
	End Property

	Public Property Get DistanceTravledPerMove() As Long
	End Property

	Public Sub MoveLeft()
	End Sub

	Public Sub MoveRight()
	End Sub

	Public Sub MoveNorth()
	End Sub

	Public Sub MoveSouth()
	End Sub
