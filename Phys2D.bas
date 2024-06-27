Attribute VB_Name = "Phys2D"
' ==============================================================
'
' ##############################################################
' #                                                            #
' #                      PPTGames Phys2D                       #
' #                                                            #
' ##############################################################
'
' » Version beta 0.1.1.0
'
' » https://pptgames.gitbook.io/phys2d-api
'
' ===============================================================


Option Explicit

Public Enum BodyType
    BodyStatic = 0
    BodyKinematic = 1
    BodyDynamic = 2
End Enum

Public Enum CollisionSide
    SideTop = 0
    SideLeft = 1
    SideBottom = 2
    SideRight = 3
End Enum

Public Function NewVec2(X As Double, Y As Double) As Vec2
    Set NewVec2 = New Vec2
    NewVec2.X = X
    NewVec2.Y = Y
End Function

Public Function NewBody(BodyType As BodyType, Pos As Vec2, Size As Vec2, Optional Vel As Vec2) As Body
    Set NewBody = New Body
    NewBody.BodyType = BodyType
    Set NewBody.Pos = Pos
    Set NewBody.Size = Size
    If Vel Is Nothing Then
        Set NewBody.Vel = New Vec2
    Else
        Set NewBody.Vel = Vel
    End If
End Function

