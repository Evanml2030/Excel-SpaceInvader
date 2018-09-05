Attribute VB_Name = "SHLoadShipOntoGameBoard"
Option Explicit

Function LoadShipOntoBoard(ByRef board As GameBoard) As Control
    Set LoadShipOntoBoard = board.Controls.Add("Forms.Image.1", Ship.ImageName)
    InitalizeShipImgControl LoadShipOntoBoard
End Function

Private Sub InitalizeShipImgControl(ByRef cntrl As Control)
    With cntrl
        .left = Ship.left
        .top = Ship.top
        .height = Ship.height
        .width = Ship.width
        .Picture = LoadPicture(Ship.ImgPathWay)
        .PictureSizeMode = 1
    End With
End Sub