VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpaceObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ISpaceObject

Private Type SpaceObjectData
    Left As Long
    Top As Long
    Height As Long
    Width As Long
    SpaceObjectType As String
    Name As String
End Type

Private this As SpaceObjectData

Public Sub SetInitialLeft(ByVal InitialLeft As Long)
    this.Left = InitialLeft
End Sub

Public Sub SetInitialTop(ByVal InitialTop As Long)
    this.Top = InitialTop
End Sub

Public Property Let Height(ByVal Height As Long)
    this.Height = Height
End Property

Public Property Let Width(ByVal Width As Long)
    this.Width = Width
End Property

Public Property Let SpaceObjectType(ByVal SpaceObjectType As SpaceObjectType)
    this.SpaceObjectType = SpaceObjectType
End Property

Public Property Let Name(ByVal Name As String)
    this.Name = Name
End Property

Public Property Get Self() As SpaceObject
    Set Self = Me
End Property

Private Property Let ISpaceObject_Top(ByVal changeTop As Long)
    this.Top = changeTop
End Property

Private Property Get ISpaceObject_Top() As Long
    ISpaceObject_Top = this.Top
End Property

Private Property Let ISpaceObject_Left(ByVal changeLeft As Long)
    this.Left = changeLeft
End Property

Private Property Get ISpaceObject_Left() As Long
    ISpaceObject_Left = this.Left
End Property

Private Property Get ISpaceObject_Height() As Long
    ISpaceObject_Height = this.Height
End Property

Private Property Get ISpaceObject_Width() As Long
    ISpaceObject_Width = this.Width
End Property

Private Property Get ISpaceObject_Name() As String
    ISpaceObject_Name = this.Name
End Property

Private Property Get ISpaceObject_SpaceObjectType() As Long
    ISpaceObject_SpaceObjectType = this.SpaceObjectType
End Property
