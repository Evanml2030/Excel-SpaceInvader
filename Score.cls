VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Score"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pscore As Long

Public Property Get Score() As Long
    Score = pscore
End Property

Public Property Let Score(ByRef value As Long)
    pscore = value
End Property

Public Sub IncrementScore()
    pscore = pscore + 1
End Sub

Public Sub UpdateGameBoard(ByVal board As GameBoard)
    board.Controls.Item("Score").Caption = pscore
End Sub

