Attribute VB_Name = "SOLoadSpaceObjectOntoGameBoard"
Option Explicit

Sub LoadSpaceObjectOntoBoard(ByRef board As GameBoard)
Dim spaceObject                     As ISpaceObject
Dim cntrl                           As Control

    Set spaceObject = DeterimineSpaceOjectType(board, Application.WorksheetFunction.RandBetween(1, 3))
    Set cntrl = AddSpaceObjectImgControlToBoard(board, spaceObject)
    InitalizeSpaceObjectImgControl cntrl, spaceObject
    AddSpaceObjectToDataCol spaceObject
    AddSpaceObjectCntrlToCntrlsCol cntrl
End Sub

Private Function DeterimineSpaceOjectType(ByRef board As GameBoard, ByRef num As Long) As ISpaceObject
    Select Case num
        Case 1
            Set DeterimineSpaceOjectType = SOAlienFactory.NewAlien(board)
        Case 2
            Set DeterimineSpaceOjectType = SOCometFactory.NewComet(board)
        Case 3
            Set DeterimineSpaceOjectType = SOStarFactory.NewStar(board)
    End Select
End Function

Private Function AddSpaceObjectImgControlToBoard(ByRef board As GameBoard, ByRef spaceObject As Object) As Control
    Set AddSpaceObjectImgControlToBoard = board.Controls.Add("Forms.Image.1", spaceObject.ImageName)
End Function

Private Sub InitalizeSpaceObjectImgControl(ByRef cntrl As Control, ByRef spaceObject As ISpaceObject)
    With cntrl
        .left = spaceObject.left
        .top = spaceObject.top
        .height = spaceObject.height
        .width = spaceObject.width
        .Picture = LoadPicture(spaceObject.ImagePathway)
        .PictureSizeMode = 1
    End With
End Sub

Private Sub AddSpaceObjectToDataCol(ByRef spaceObject As ISpaceObject)
    SpaceObjectDataCol.Add spaceObject
End Sub

Private Sub AddSpaceObjectCntrlToCntrlsCol(ByRef cntrl As Control)
    SpaceObjectCntrlsCol.Add cntrl
End Sub
