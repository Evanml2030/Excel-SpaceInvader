Attribute VB_Name = "CreateGameItem"
Option Explicit

Public Enum objectType
    alien = 1
    comet = 2
    star = 3
    missile = 4
    Ship = 5
End Enum

Public Sub CreateGameItem(ByVal board As GameBoard, ByVal val As objectType)
Dim CreateGameItem                      As iboundcontrol
    Select Case val
        Case objectType.alien
            Set CreateGameItem = New SpaceObjectAlien
            Set CreateGameItem.spaceObject = SpaceObjectFactory.NewAlien(board)
            CollectionIncomingSpaceObjects.Add CreateGameItem
        Case objectType.comet
            Set CreateGameItem = New SpaceObjectComet
            Set CreateGameItem.spaceObject = SpaceObjectFactory.NewComet(board)
            CollectionIncomingSpaceObjects.Add CreateGameItem
        Case objectType.star
            Set CreateGameItem = New SpaceObjectStar
            Set CreateGameItem.spaceObject = SpaceObjectFactory.NewStar(board)
            CollectionIncomingSpaceObjects.Add CreateGameItem
        Case objectType.Ship
            Set CreateGameItem = New SpaceObjectShip
            Set CreateGameItem.spaceObject = SpaceObjectFactory.NewShip(board)
            CollectionShips.Add CreateGameItem
        Case objectType.missile
            Set CreateGameItem = New SpaceObjectMissile
            Set CreateGameItem.spaceObject = SpaceObjectFactory.NewMissile(board)
            CollectionMissiles.Add CreateGameItem
    End Select
    
    Set CreateGameItem.Control = LoadControl(board, CreateGameItem)
    InitializeControl CreateGameItem
End Sub

Private Function LoadControl(ByVal board As GameBoard, ByVal gameItem As iboundcontrol) As Control
     Set LoadControl = board.Controls.Add("Forms.Image.1", gameItem.spaceObject.ImageName)
End Function

Private Sub InitializeControl(ByVal gameItem As iboundcontrol)
   With gameItem
        .Control.left = gameItem.spaceObject.left
        .Control.top = gameItem.spaceObject.top
        .Control.height = gameItem.spaceObject.height
        .Control.width = gameItem.spaceObject.width
        .Control.Picture = LoadPicture(gameItem.spaceObject.ImagePathway)
        .Control.PictureSizeMode = 1
    End With
End Sub
