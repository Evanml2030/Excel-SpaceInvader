VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MissileCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pcount As Long

Public Property Get Count() As Long
    Count = pcount
End Property

Public Property Let Count(ByRef value As Long)
    pcount = value
End Property

Public Sub IncrementMissileCount()
    pcount = pcount + 1
End Sub

Public Sub UpdateGameBoard(ByVal board As GameBoard)
    board.Controls.Item("MissileCount").Caption = 25 - pcount
End Sub

