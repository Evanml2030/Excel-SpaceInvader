VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IBoundControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get Control() As Control
End Property

Public Property Set Control(bound As Control)
End Property

Public Property Get spaceObject() As ISpaceObject
End Property

Public Property Set spaceObject(bound As ISpaceObject)
End Property

