VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Missle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type MissleData
    left As Long
    top As Long
    ImgPathWay As String
    ImageName As String
    width As Long
    height As Long
End Type

Private this As MissleData

Public Property Let ImgPathWay(ByRef pathWayToImg As String)
    this.ImgPathWay = pathWayToImg
End Property
Public Property Get ImgPathWay() As String
    ImgPathWay = this.ImgPathWay
End Property
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
Public Property Get Self() As missle
    Set Self = Me
End Property
Public Property Let ImageName(ByRef Name As String)
    this.ImageName = Name
End Property
 Public Property Get ImageName() As String
    ImageName = this.ImageName
 End Property
Public Property Let left(ByRef changeLeft As Long)
    this.left = changeLeft
End Property
Public Property Get left() As Long
    left = this.left
End Property
Public Property Let top(ByRef changeTop As Long)
    this.top = changeTop
End Property
Public Property Get top() As Long
    top = this.top
End Property


