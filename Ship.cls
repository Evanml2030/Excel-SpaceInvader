VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ship"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type ShipData
    left As Long
    top As Long
    ImgPathWay As String
    ImageName As String
    width As Long
    height As Long
    Name As String
End Type

Private this As ShipData
Private Sub Class_Initialize()
    this.ImgPathWay = "Z:\Desktop Storage\EXCEL & C# PRACTICE\SpaceInvaders\spaceShip.jpg"
    SetInitialLeft GameBoard.width / 2
    SetInitialTop GameBoard.height - (GameBoard.height / 8.5)
    this.width = ScaleItems.MaxSize
    this.height = ScaleItems.MaxSize
    this.ImageName = "Ship"
End Sub

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
Public Property Get Self() As SpaceObjectComet
    Set Self = Me
End Property
Public Property Let left(ByRef left As Long)
    this.left = left
End Property
Public Property Get left() As Long
    left = this.left
End Property
Public Property Let top(ByRef top As Long)
    this.height = height
End Property
Public Property Get top() As Long
    top = this.top
End Property
Public Property Let ImageName(ByRef ImageName As String)
    this.ImageName = height
End Property
Public Property Get ImageName() As String
    ImageName = this.ImageName
End Property
