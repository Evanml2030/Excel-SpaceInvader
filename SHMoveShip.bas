Attribute VB_Name = "SHMoveShip"
Public Function moveShipLeft(ByRef board As GameBoard)
Dim ShipCntrl                       As Control
Set ShipCntrl = board.Controls(Ship.ImageName)

    If Ship.left > 0 Then
        Ship.left = Ship.left - 5
        ShipCntrl.left = Ship.left
    End If
End Function

Function moveShipRight(ByRef board As GameBoard)
Dim ShipCntrl                        As Control
Set ShipCntrl = board.Controls(Ship.ImageName)

    If Ship.left + Ship.width < board.width Then
        Ship.left = Ship.left + 5
        ShipCntrl.left = Ship.left
    Else
    End If
End Function