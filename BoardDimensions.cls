VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BoardDimensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

	Private Type BoardDimensionsData
	    Width As Long
	    Height As Long
	End Type

	Private this As BoardDimensionsData

	Public Property Let Width(ByVal Width As Long)
	    this.Width = Width
	End Property

	Public Property Get Width() As Long
	    Width = this.Width
	End Property

	Public Property Let Height(ByVal Height As Long)
	    this.Height = Height
	End Property

	Public Property Get Height() As Long
	    Height = this.Height
	End Property
