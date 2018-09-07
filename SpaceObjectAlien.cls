VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpaceObjectAlien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements iboundcontrol
Implements ISpaceObject

Private Type AlienData
    left As Long
    top As Long
    width As Long
    height As Long
    ImgPathWay As String
    ImageName As String
    MyControl As Control
    MySpaceObj As SpaceObjectAlien
End Type

Private this As AlienData

Public Sub SetInitialLeft(ByRef initialLeft As Long)
    this.left = initialLeft
End Sub

Public Sub SetInitialTop(ByRef initialTop As Long)
    this.top = initialTop
End Sub

Public Property Let width(ByRef width As Long)
    this.width = width
End Property

Public Property Get width() As Long
    width = this.width
End Property

Public Property Let height(ByRef height As Long)
    this.height = height
End Property

Public Property Get height() As Long
    height = this.height
End Property

Public Property Let ImgPathWay(ByRef pathWayToImg As String)
    this.ImgPathWay = pathWayToImg
End Property

Public Property Get ImgPathWay() As String
    ImgPathWay = this.ImgPathWay
End Property

Public Property Let ImageName(ByRef Name As String)
    this.ImageName = Name
End Property

Public Property Get ImageName() As String
    ImageName = this.ImageName
End Property

Public Property Get Self() As SpaceObjectAlien
    Set Self = Me
End Property

Private Property Get IspaceObject_ImagePathway() As String
    IspaceObject_ImagePathway = this.ImgPathWay
End Property

Private Property Get ISpaceObject_ImageName() As String
    this.ImageName = ISpaceObject_ImageName
End Property

Private Property Let ISpaceObject_Left(ByRef changeLeft As Long)
    this.left = changeLeft
End Property

Private Property Get ISpaceObject_Left() As Long
    ISpaceObject_Left = this.left
End Property

Private Property Let ISpaceObject_Top(ByRef changeTop As Long)
    this.top = changeTop
End Property

Private Property Get ISpaceObject_Top() As Long
    ISpaceObject_Top = this.top
End Property

Private Property Get ISpaceObject_Height() As Long
    ISpaceObject_Height = this.height
End Property

Private Property Get ISpaceObject_Width() As Long
    ISpaceObject_Width = this.width
End Property

Public Property Set IBoundControl_Control(ByRef form As Control)
    Set this.MyControl = form
End Property

Public Property Get IBoundControl_Control() As Control
    Set IBoundControl_Control = this.MyControl
End Property

Public Property Set IBoundControl_SpaceObject(ByRef spcObj As ISpaceObject)
    Set this.MySpaceObj = spcObj
End Property

Public Property Get IBoundControl_SpaceObject() As ISpaceObject
    Set IBoundControl_SpaceObject = this.MySpaceObj
End Property
