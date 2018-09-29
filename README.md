# SpaceInvader
VBA version of a classic game

This code runs from Excel

.frx file must in same directory as .frm file to be able to import the frm file

You must also place all of the "skins" for the space objects, the jpgs, inside the same directory as the excel file that you are launching this from. We are uploading via code that reads  ActiveWorkbook.Path & "\SpaceShip.jpg", ActiveWorkbook.Path & "Missile.jpg" etc

Positive: I have decoupled the view from the control / presenter. I have implemented what I think is an MVP style design. I have refactored much of the code, making it leaner and meaner. I explored was to reduce the number of factories, but found that I to either A) create classes with "constructor" that set initial values B) store initial values in separate functions that I call from the factory method. I felt that my solution was most elegant of these.

Negative: I have two BIG issues. First, my method of scaling is not working. I am taking the Game Board dimensions and using them to set the width / length of my game pieces. Somehow my method of figuring these values is not working.

Second I have moved from a custom collection to a dictionary as my method of storing game pieces. However as I loop through the pieces I every so often get a 424 Object Required error. This usually comes in the following line:

If CheckIfCollided(GamePiecesCollection.Item(MissileKey), GamePiecesCollection.Item(IncomingSpaceObjectKey)) Then

My handle ship incoming space objects collision function is not working at all LOL. Almost makes me want to switch back to a custom collection. But for some reason I thought that a dictionary would make condensing all of my collections, that is storing my ship, missiles and incoming space objects in the same collection, would be easier than fitting them into a custom collection.

Here is the code. Note the gameboard form wont load without frx file, which I cannot post here:

Note that I am using the Microsoft Scripting Runtime Library
