Attribute VB_Name = "CheckObjectType"
Option Explicit
 
Public Function IsIncomingSpaceObject(ByVal SpaceObjectOne As ISpaceObject) As Boolean
    If SpaceObjectOne.SpaceObjectType < Missile Then
        IsIncomingSpaceObject = True
    Else
        IsIncomingSpaceObject = False
    End If
End Function

Public Function IsMissile(ByVal SpaceObjectTwo As ISpaceObject) As Boolean
    If SpaceObjectTwo.SpaceObjectType = Missile Then
        IsMissile = True
    Else
        IsMissile = False
    End If
End Function
