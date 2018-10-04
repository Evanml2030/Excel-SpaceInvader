# SpaceInvader
VBA version of a classic game

This code runs from Excel

.frx file must in same directory as .frm file to be able to import the frm file

Note that I am using the Microsoft Scripting Runtime Library

You must also place all of the "skins" for the space objects, the jpgs, inside the same directory as the excel file that you are launching this from. We are uploading via code that reads  ActiveWorkbook.Path & "\SpaceShip.jpg", ActiveWorkbook.Path & "Missile.jpg" etc

From:
https://codereview.stackexchange.com/questions/204913/space-invader-style-game-model-presenter-decoupled-from-view-functional-ag

I have tried to incorporate all the advice I have received so far. Thank you for taking the time to review and comment.

I have now decoupled the Model / Presenter from the View as much as VBA allows (please correct me if wrong). There are two functions from inside the view which the presenter calls. One of these functions ends the game, the other takes in a collection which the view promptly displays to the player.

When your ship collides with a space object you get a game over message box and then the userform unloads, an upgrade over the previous version which lead to crashes when you tried to run game twice.

Rather than throwing all of the spaceobjects into a single collection I have kept the separate collections for each "type" of spaceobject; incoming, missile and ship. Separate Collections for these objects makes handling collisions and object removal more manageable.

In the immediate term I want to re introduce scoring and a missile count limit. Also I need to fix the scaling of the objects. I have them in a ratio with gameboard height / width. Somehow it was not working for me? Then I am thinking about heat seeking missiles (give me a chance to work on algorithim stuff a bit, should be fun)

Big ups to:

StopWatch was put together by the fellow who runs bytecomb, a great site for vba tips. Link: https://bytecomb.com/accurate-performance-timers-in-vba/

https://openclipart.org/ For the Jpgs that I use
