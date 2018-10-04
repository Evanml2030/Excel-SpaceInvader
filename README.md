# SpaceInvader
VBA version of a classic game

This code runs from Excel

.frx file must in same directory as .frm file to be able to import the frm file

Note that I am using the Microsoft Scripting Runtime Library

You must also place all of the "skins" for the space objects, the jpgs, inside the same directory as the excel file that you are launching this from. We are uploading via code that reads  ActiveWorkbook.Path & "\SpaceShip.jpg", ActiveWorkbook.Path & "Missile.jpg" etc

OK. I have now decoupled the Model / Presenter from the View. The game is fully functional again. When your ship collides with a space object you game a game over message box and the userform unloads. Rather than throwing all of the spaceobjects into a single collection I have resperated them into seperate collections. This makes handling collisions checking and object removal much more manageable.

In the immediate term I want to re introduce scoring and a missile count limit. Then perhaps missile pack and heat seeking missiles (give me a chance to work on algorithim stuff a bit, should be fun)
