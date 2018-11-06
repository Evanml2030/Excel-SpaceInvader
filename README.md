# SpaceInvader
VBA version of a classic game

This code runs from Excel

.frx file must in same directory as .frm file to be able to import the frm file

Note that I am using the Microsoft Scripting Runtime Library

You must also place all of the "skins" for the space objects, the jpgs, inside the same directory as the excel file that you are launching this from. We are uploading via code that reads  ActiveWorkbook.Path & "\SpaceShip.jpg", ActiveWorkbook.Path & "Missile.jpg" etc

This is a work in progress. I am incorporating all of the advice I am getting from:
https://codereview.stackexchange.com/questions/204913/space-invader-style-game-model-presenter-decoupled-from-view-functional-ag/205223?noredirect=1#comment396696_205223

Thank you all for your time!

Big ups to:

StopWatch was put together by the fellow who runs bytecomb, a great site for vba tips. Link: https://bytecomb.com/accurate-performance-timers-in-vba/

https://openclipart.org/ For the Jpgs that I use
