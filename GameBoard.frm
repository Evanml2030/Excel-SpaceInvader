VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GameBoard 
   Caption         =   "SpaceInvaders"
   ClientHeight    =   10410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10755
   OleObjectBlob   =   "GameBoard.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GameBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Private Sub UserForm_Activate()
        GameLogic.RunGame Me.InsideHeight, Me.InsideWidth
    End Sub

    Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
        Dim passVal As Long
        Select Case KeyCode.Value
        Case 37, 39, 32
            passVal = CInt(KeyCode)
            GameLogic.HandleSendKeys passVal
        End Select
    End Sub

    Public Sub CloseGame()
        MsgBox "GAMEOVER"
        Unload Me
    End Sub

    Public Sub RefreshGameBoard(ByVal ControlsToAdd As Collection)
        Dim Ctrl As Image
        Dim SpaceObjectIndex As Variant

        For Each Ctrl In Me.Controls
            Me.Controls.Remove Ctrl.Name
        Next Ctrl
        
        For SpaceObjectIndex = 1 To ControlsToAdd.Count
            Set Ctrl = Me.Controls.Add("Forms.Image.1", ControlsToAdd.Item(SpaceObjectIndex).Name, True)
            Ctrl.Left = ControlsToAdd.Item(SpaceObjectIndex).Left
            Ctrl.Top = ControlsToAdd.Item(SpaceObjectIndex).Top
            Ctrl.Height = ControlsToAdd.Item(SpaceObjectIndex).Height
            Ctrl.Width = ControlsToAdd.Item(SpaceObjectIndex).Width
            Ctrl.Picture = GetPicture(ControlsToAdd.Item(SpaceObjectIndex).SpaceObjectType)
            Ctrl.PictureSizeMode = fmPictureSizeModeStretch
        Next SpaceObjectIndex
    End Sub

        Private Function GetPicture(ByVal SpaceObjectType As SpaceObjectType) As Object
            Select Case SpaceObjectType
                Case Alien
                    Set GetPicture = StoreUserImages.Alien.Picture
                Case Comet
                    Set GetPicture = StoreUserImages.Comet.Picture
                Case Star
                    Set GetPicture = StoreUserImages.Star.Picture
                Case Missile
                    Set GetPicture = StoreUserImages.Missile.Picture
                Case Ship
                    Set GetPicture = StoreUserImages.Ship.Picture
            End Select
        End Function

