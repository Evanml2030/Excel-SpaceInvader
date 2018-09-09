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
Dim CreateGameItem                      As IBoundControl
    Select Case val
        Case objectType.alien
            Set CreateGameItem = New SpaceObjectAlien
            Set CreateGameItem.spaceObject = SpaceObjectFactory.NewSpaceObjectAlien(board)
            CollectionIncomingSpaceObjects.Add CreateGameItem
        Case objectType.comet
            Set CreateGameItem = New SpaceObjectComet
            Set CreateGameItem.spaceObject = SpaceObjectFactory.NewSpaceObjectComet(board)
            CollectionIncomingSpaceObjects.Add CreateGameItem
        Case objectType.star
            Set CreateGameItem = New SpaceObjectStar
            Set CreateGameItem.spaceObject = SpaceObjectFactory.NewSpaceObjectStar(board)
            CollectionIncomingSpaceObjects.Add CreateGameItem
        Case objectType.Ship
            Set CreateGameItem = New SpaceObjectShip
            Set CreateGameItem.spaceObject = SpaceObjectFactory.NewSpaceObjectShip(board)
            CollectionShips.Add CreateGameItem
        Case objectType.missile
            Set CreateGameItem = New SpaceObjectMissile
            Set CreateGameItem.spaceObject = SpaceObjectFactory.NewSpaceObjectMissile(board)
            CollectionMissiles.Add CreateGameItem
            MissileCount.UpdateGameBoard board
    End Select
    
    Set CreateGameItem.Control = LoadControl(board, CreateGameItem)
    InitializeControl CreateGameItem
End Sub

Private Function LoadControl(ByVal board As GameBoard, ByVal gameItem As IBoundControl) As Control
     Set LoadControl = board.Controls.Add("Forms.Image.1", gameItem.spaceObject.ImageName)
End Function

Private Sub InitializeControl(ByVal gameItem As IBoundControl)
   With gameItem
        .Control.left = gameItem.spaceObject.left
        .Control.top = gameItem.spaceObject.top
        .Control.height = gameItem.spaceObject.height
        .Control.width = gameItem.spaceObject.width
        .Control.Picture = LoadPicture(gameItem.spaceObject.ImagePathway)
        .Control.PictureSizeMode = 1
    End With
End Sub
