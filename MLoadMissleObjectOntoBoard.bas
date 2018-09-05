Attribute VB_Name = "MLoadMissleObjectOntoBoard"
Option Explicit

Sub LoadMissleObjectOntoBoard(ByRef board As GameBoard)
Dim missleObject                    As missle
Dim cntrl                           As Control

    Set missleObject = MMissleFactory.NewMissle
    Set cntrl = AddMissleObjectImgControlToBoard(board, missleObject)
    InitalizeMissleObjectImgControl cntrl, missleObject
    AddMissleObjectToDataCol missleObject
    AddMissleObjectCntrlToCntrlsCol cntrl
End Sub

Private Function AddMissleObjectImgControlToBoard(ByRef board As GameBoard, ByRef missleObject As Object) As Control
    Set AddMissleObjectImgControlToBoard = board.Controls.Add("Forms.Image.1", missleObject.ImageName)
End Function

Private Sub InitalizeMissleObjectImgControl(ByRef cntrl As Control, ByRef missleObject As missle)
    With cntrl
        .left = missleObject.left
        .top = missleObject.top
        .height = missleObject.height
        .width = missleObject.width
        .Picture = LoadPicture(missleObject.ImgPathWay)
        .PictureSizeMode = 1
    End With
End Sub

Private Sub AddMissleObjectToDataCol(ByRef missleObject As missle)
    MissleObjectsDataCol.Add missleObject
End Sub

Private Sub AddMissleObjectCntrlToCntrlsCol(ByRef cntrl As Control)
    MissleCntrlsCol.Add cntrl
End Sub
