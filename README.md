## About
Phys2D is a simple 2D physics engine made in VBA for PowerPoint games.

#### Credits
The code was highly inspired and adapted from [OneLoneCoder's C++ PGE Rectangles project](https://github.com/OneLoneCoder/Javidx9/blob/master/PixelGameEngine/SmallerProjects/OneLoneCoder_PGE_Rectangles.cpp).

## Documentation
Visit the official documentation at [this link](https://pptgames.gitbook.io/pptg-coding/v/phys2d-api).

## Dependencies
Phys2D requires the following dependencies.
- PPTGames Better Arrays (Download the latest version from [official page](https://pptgamespt.wixsite.com/pptg-coding/better-arrays))
- Microsoft Scripting Runtime ([Click here](https://pptgamespt.wixsite.com/pptg-coding/tutorial-enable-dictionary) to learn how to add to your project)

## Example
```vba
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Dim DeltaTime As Double
Dim World As World
Dim WithEvents Player As Body

Function GetSysTime() As Double
    GetSysTime = GetTickCount / 1000
End Function

Sub InitSim()
    Dim Time0 As Double
    Init
    
    Do While ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 1
        Time0 = GetSysTime
        Update ' update simulation
        Shapes("timer").TextFrame.TextRange.Text = Timer
        DoEvents
        DeltaTime = GetSysTime - Time0 ' calculate elapsed time
    Loop
End Sub

Sub Init()
    Set World = New World
    World.Gravity.Y = 600
    
    Set Player = World.AddBody(BodyKinematic, NewVec2(Shapes("player").Left, Shapes("player").Top), NewVec2(Shapes("player").Width, Shapes("player").Height))
    Player.Mass = 50
    
    Dim Shp As Shape
    For Each Shp In Shapes.Range
        If Left(Shp.Name, 4) = "wall" Then
            Shp.Name = "wall" & World.AddBody(BodyStatic, NewVec2(Shp.Left, Shp.Top), NewVec2(Shp.Width, Shp.Height)).Index
        End If
    Next
End Sub

Sub Update()
    UpdateControls
    
    World.Update DeltaTime
    
    ' Update the player's shape position
    Shapes("player").Left = Player.Pos.X
    Shapes("player").Top = Player.Pos.Y
End Sub

Sub UpdateControls()
    If GetAsyncKeyState(vbKeyA) Then
        Player.Vel.X = -200
    ElseIf GetAsyncKeyState(vbKeyD) Then
        Player.Vel.X = 200
    Else
        Player.Vel.X = 0
    End If
    
    If GetAsyncKeyState(vbKeyW) Then
        Player.Vel.Y = -200
    ElseIf GetAsyncKeyState(vbKeyS) Then
        Player.Vel.Y = 200
    Else
        Player.Vel.Y = 0
    End If
End Sub

' Make the collided body red when the player collides with it
Private Sub Player_OnCollide(Body As Body, Side As CollisionSide)
    Shapes("obj" & Body.Index).Fill.ForeColor.RGB = RGB(255, 0, 0)
End Sub

' Revert the body's color when the player is no longer colliding with it
Private Sub Player_OnAvert(Body As Body, Side As CollisionSide)
    Shapes("obj" & Body.Index).Fill.ForeColor.RGB = RGB(100, 100, 100)
End Sub
```
