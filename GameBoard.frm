VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GameBoard 
   Caption         =   "Space"
   ClientHeight    =   9870
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7470
   OleObjectBlob   =   "GameBoard.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GameBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Dim passVal                             As Long
    Select Case KeyCode.value
        Case 37, 39, 32
            passVal = CInt(KeyCode)
            GameLogic.HandleSendKeys Me, passVal
    End Select
End Sub
