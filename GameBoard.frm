VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GameBoard 
   Caption         =   "UserForm1"
   ClientHeight    =   9495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7725
   OleObjectBlob   =   "GameBoard.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GameBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Activate()
    GameLogic.RunGame Me.InsideHeight, Me.InsideWidth
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Dim passVal                                     As Long
    Select Case KeyCode.value
        Case 37, 39, 32
            passVal = CInt(KeyCode)
            GameLogic.HandleSendKeys passVal
    End Select
End Sub

Public Sub RefreshGameBoard(ByVal ControlsToAdd As Scripting.Dictionary)
Dim Ctrl                                            As Image
Dim SpaceObjectIndex                                As Variant

    For Each Ctrl In Me.Controls
        Me.Controls.remove Ctrl.Name
    Next Ctrl
    
    For Each SpaceObjectIndex In ControlsToAdd.Keys()
        Set Ctrl = Me.Controls.Add("Forms.Image.1", ControlsToAdd.Item(SpaceObjectIndex).Name, True)
        Ctrl.Left = ControlsToAdd.Item(SpaceObjectIndex).Left
        Ctrl.Top = ControlsToAdd.Item(SpaceObjectIndex).Top
        Ctrl.Height = ControlsToAdd.Item(SpaceObjectIndex).Height
        Ctrl.Width = ControlsToAdd.Item(SpaceObjectIndex).Width
        Ctrl.Picture = LoadPicture(LinkToImage(ControlsToAdd.Item(SpaceObjectIndex).SpaceObjectType))
        Ctrl.PictureSizeMode = fmPictureSizeModeZoom
    Next SpaceObjectIndex
End Sub

Private Function LinkToImage(ByVal SpaceObjectType As SpaceObjectType) As String
    Select Case SpaceObjectType
        Case Alien
            LinkToImage = "C:\Users\evanm\OneDrive\Desktop\Excel\SpaceInvader\AlienShip.jpg"
        Case Comet
            LinkToImage = "C:\Users\evanm\OneDrive\Desktop\Excel\SpaceInvader\Comet.jpg"
        Case Star
            LinkToImage = "C:\Users\evanm\OneDrive\Desktop\Excel\SpaceInvader\Star.jpg"
        Case Missile
            LinkToImage = "C:\Users\evanm\OneDrive\Desktop\Excel\SpaceInvader\Missile.jpg"
        Case Ship
            LinkToImage = "C:\Users\evanm\OneDrive\Desktop\Excel\SpaceInvader\SpaceShip.jpg"
    End Select
End Function


