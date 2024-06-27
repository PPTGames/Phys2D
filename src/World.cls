VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "World"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Gravity As New Vec2

Private pBodies As New Dictionary
Private pStaticBodies As New List
Private pDynBodies As New List

Private CurBodyIndex As Long

Public Sub Update(ElapsedTime As Double)
    Dim i As Long
    For i = 0 To pDynBodies.Length - 1
        UpdateBody pDynBodies(i), ElapsedTime
    Next
End Sub

Private Sub UpdateBody(Body As Body, ElapsedTime As Double)
    Dim i As Long
    Dim cp As New Vec2
    Dim cn As New Vec2
    Dim ct As Double
    Dim z As New List
    
    Body.Update ElapsedTime
    
    If Body.CollisionEnabled Then
    
        For i = 0 To pStaticBodies.Length - 1
            If DynamicRectVsRect(Body, pStaticBodies(i), cp, cn, ct, ElapsedTime) Then
                z.Add Array(i, ct)
            Else
                Body.SetContact 0, Nothing
                Body.SetContact 1, Nothing
                Body.SetContact 2, Nothing
                Body.SetContact 3, Nothing
            End If
        Next
        
        If z.Length > 0 Then z.Items = SortZ(z.Items)
        
        For i = 0 To z.Length - 1
            ResolveDynamicRectVsRect Body, ElapsedTime, pStaticBodies(z(i)(0))
        Next
    
    End If
    
    Body.Pos.X = Body.Pos.X + Body.Vel.X * ElapsedTime
    Body.Pos.Y = Body.Pos.Y + Body.Vel.Y * ElapsedTime
End Sub

Public Function AddBody(BodyType As BodyType, Pos As Vec2, Size As Vec2, Optional Vel As Vec2) As Body
    Dim Body As New Body
    Body.BodyType = BodyType
    Set Body.Pos = Pos
    Set Body.Size = Size
    If Vel Is Nothing Then
        Set Body.Vel = New Vec2
    Else
        Set Body.Vel = Vel
    End If
    AddBody2 Body
    Set AddBody = Body
End Function

Public Sub AddBody2(Body As Body)
    Set Body.World = Me
    Body.Index = CurBodyIndex
    CurBodyIndex = CurBodyIndex + 1
    pBodies.Add CurBodyIndex, Body
    If Body.BodyType = BodyStatic Then
        pStaticBodies.Add Body
    Else
        pDynBodies.Add Body
    End If
End Sub

Public Function GetBody(Index As Long) As Body
    Set GetBody = pBodies(Index)
End Function

Public Sub RemoveBody(Index As Long)
    pBodies.Remove Index
End Sub

Public Sub RemoveBody2(Body As Body)
    RemoveBody Body.Index
End Sub

Private Function PointVsRect(Point As Vec2, Rect As Body) As Boolean
    PointVsRect = Point.X >= Rect.Pos.X And Point.X < Rect.Pos.X + Rect.Size.X _
            And Point.Y >= Rect.Pos.Y And Point.Y < Rect.Pos.Y + Rect.Size.Y
End Function

Private Function RectVsRect(Rect1 As Body, Rect2 As Body) As Boolean
    RectVsRect = Rect1.Pos.X < Rect2.Pos.X + Rect2.Size.X And Rect1.Pos.X + Rect1.Size.X > Rect2.Pos.X _
            And Rect1.Pos.Y < Rect2.Pos.Y + Rect2.Size.Y And Rect1.Pos.Y + Rect1.Size.Y > Rect2.Pos.Y
End Function

Private Function RayVsRect(RayOrigin As Vec2, RayDir As Vec2, Target As Body, _
        ByRef ContactPoint As Vec2, ByRef ContactNormal As Vec2, ByRef tHitNear2 As Double) As Boolean
    Dim tNear As New Vec2
    Dim tFar As New Vec2
    Dim InvDir As New Vec2
    
    Dim tNearXInf As Integer
    Dim tNearYInf As Integer
    Dim tFarXInf As Integer
    Dim tFarYInf As Integer
    
    ContactPoint.X = 0
    ContactPoint.Y = 0
    ContactNormal.X = 0
    ContactNormal.Y = 0
    
    If RayDir.X <> 0 Then InvDir.X = 1 / RayDir.X
    If RayDir.Y <> 0 Then InvDir.Y = 1 / RayDir.Y
    
    If Target.Pos.X - RayOrigin.X = 0 And RayDir.X = 0 Or _
        Target.Pos.X + Target.Size.X - RayOrigin.X = 0 And RayDir.X = 0 Or _
        Target.Pos.Y - RayOrigin.Y = 0 And RayDir.Y = 0 Or _
        Target.Pos.Y + Target.Size.Y - RayOrigin.Y = 0 And RayDir.Y = 0 Then Exit Function
    
    If RayDir.X = 0 Then
        tNearXInf = Sgn(Target.Pos.X - RayOrigin.X)
        tFarXInf = Sgn(Target.Pos.X + Target.Size.X - RayOrigin.X)
    Else
        tNear.X = (Target.Pos.X - RayOrigin.X) * InvDir.X
        tFar.X = (Target.Pos.X + Target.Size.X - RayOrigin.X) * InvDir.X
    End If
    
    If RayDir.Y = 0 Then
        tNearYInf = Sgn(Target.Pos.Y - RayOrigin.Y)
        tFarYInf = Sgn(Target.Pos.Y + Target.Size.Y - RayOrigin.Y)
    Else
        tNear.Y = (Target.Pos.Y - RayOrigin.Y) * InvDir.Y
        tFar.Y = (Target.Pos.Y + Target.Size.Y - RayOrigin.Y) * InvDir.Y
    End If
    
    If CmpInf(tNear.X, tNearXInf, tFar.X, tFarXInf) > 0 Then
        SwapX tNear, tFar
        SwapInt tNearXInf, tFarXInf
    End If
    If CmpInf(tNear.Y, tNearYInf, tFar.Y, tFarYInf) > 0 Then
        SwapY tNear, tFar
        SwapInt tNearYInf, tFarYInf
    End If
    
    If CmpInf(tNear.X, tNearXInf, tFar.Y, tFarYInf) > 0 Or CmpInf(tNear.Y, tNearYInf, tFar.X, tFarXInf) > 0 Then Exit Function
    
    Dim tHitNear As Double
    Dim tHitNearInf As Double
    Dim tHitFar As Double
    Dim tHitFarInf As Double
    
    If CmpInf(tNear.X, tNearXInf, tNear.Y, tNearYInf) > 0 Then
        If tNearXInf <> 0 Then
            tHitNearInf = tNearXInf
        Else
            tHitNear = tNear.X
        End If
    Else
        If tNearYInf <> 0 Then
            tHitNearInf = tNearYInf
        Else
            tHitNear = tNear.Y
        End If
    End If
    
    If CmpInf(tFar.X, tFarXInf, tFar.Y, tFarYInf) < 0 Then
        If tFarXInf <> 0 Then
            tHitFarInf = tFarXInf
        Else
            tHitFar = tFar.X
        End If
    Else
        If tFarYInf <> 0 Then
            tHitFarInf = tFarYInf
        Else
            tHitFar = tFar.Y
        End If
    End If
    
    If tHitNearInf <> 0 Then Exit Function
    
    If tHitFar < 0 Or tHitFarInf < 0 Then Exit Function
    
    tHitNear2 = tHitNear
    
    ContactPoint.X = RayOrigin.X + tHitNear * RayDir.X
    ContactPoint.Y = RayOrigin.Y + tHitNear * RayDir.Y
    
    If CmpInf(tNear.X, tNearXInf, tNear.Y, tNearYInf) > 0 Then
        If RayDir.X < 0 Then
            ContactNormal.X = 1
        Else
            ContactNormal.X = -1
        End If
    ElseIf CmpInf(tNear.X, tNearXInf, tNear.Y, tNearYInf) < 0 Then
        If RayDir.Y < 0 Then
            ContactNormal.Y = 1
        Else
            ContactNormal.Y = -1
        End If
    End If
    
    RayVsRect = True
End Function

Private Function CmpInf(X As Double, XInf As Integer, Y As Double, YInf As Integer) As Double
    If XInf = 0 And YInf = 0 Then
        CmpInf = X - Y
    ElseIf XInf <> 0 And YInf = 0 Then
        If XInf < 0 Then
            CmpInf = -1
        ElseIf XInf > 0 Then
            CmpInf = 1
        End If
    ElseIf XInf = 0 And YInf <> 0 Then
        If YInf < 0 Then
            CmpInf = 1
        ElseIf YInf > 0 Then
            CmpInf = -1
        End If
    ElseIf XInf <> 0 And YInf <> 0 Then
        CmpInf = XInf - YInf
    End If
End Function

Private Function DynamicRectVsRect(BodyIn As Body, Target As Body, _
        ByRef ContactPoint As Vec2, ByRef ContactNormal As Vec2, ByRef ContactTime As Double, ElapsedTime As Double) As Boolean
    If BodyIn.Vel.X = 0 And BodyIn.Vel.Y = 0 Then Exit Function
    
    Dim ExpandedTarget As New Body
    ExpandedTarget.Pos.X = Target.Pos.X - BodyIn.Size.X / 2
    ExpandedTarget.Pos.Y = Target.Pos.Y - BodyIn.Size.Y / 2
    ExpandedTarget.Size.X = Target.Size.X + BodyIn.Size.X
    ExpandedTarget.Size.Y = Target.Size.Y + BodyIn.Size.Y
    
    If RayVsRect(NewVec2(BodyIn.Pos.X + BodyIn.Size.X / 2, BodyIn.Pos.Y + BodyIn.Size.Y / 2), _
            NewVec2(BodyIn.Vel.X * ElapsedTime, BodyIn.Vel.Y * ElapsedTime), ExpandedTarget, ContactPoint, ContactNormal, ContactTime) Then
        If ContactTime >= 0 And ContactTime < 1 Then
            DynamicRectVsRect = True
            Exit Function
        End If
    End If
End Function

Private Function ResolveDynamicRectVsRect(RDyn As Body, TimeStep As Double, RStatic As Body) As Boolean
    Dim ContactPoint As New Vec2
    Dim ContactNormal As New Vec2
    Dim ContactTime As Double
    If DynamicRectVsRect(RDyn, RStatic, ContactPoint, ContactNormal, ContactTime, TimeStep) Then
        If ContactNormal.Y > 0 Then RDyn.SetContact 0, RStatic Else RDyn.SetContact 0, Nothing
        If ContactNormal.X < 0 Then RDyn.SetContact 1, RStatic Else RDyn.SetContact 1, Nothing
        If ContactNormal.Y < 0 Then RDyn.SetContact 2, RStatic Else RDyn.SetContact 2, Nothing
        If ContactNormal.X > 0 Then RDyn.SetContact 3, RStatic Else RDyn.SetContact 3, Nothing
        
        RDyn.Vel.X = RDyn.Vel.X + ContactNormal.X * Abs(RDyn.Vel.X) * (1 - ContactTime)
        RDyn.Vel.Y = RDyn.Vel.Y + ContactNormal.Y * Abs(RDyn.Vel.Y) * (1 - ContactTime)
        
        ResolveDynamicRectVsRect = True
    End If
End Function

Private Sub SwapX(ByRef a As Vec2, ByRef b As Vec2)
    Dim c As Double
    c = a.X
    a.X = b.X
    b.X = c
End Sub

Private Sub SwapY(ByRef a As Vec2, ByRef b As Vec2)
    Dim c As Double
    c = a.Y
    a.Y = b.Y
    b.Y = c
End Sub

Private Sub SwapInt(ByRef a As Integer, ByRef b As Integer)
    Dim c As Integer
    c = a
    a = b
    b = c
End Sub

Private Function SortZ(arr As Variant) As Variant
    Dim Tmp As Variant, i As Long, j As Long
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i)(1) > arr(j)(1) Then
                If IsObject(arr(i)) And IsObject(arr(j)) Then
                    Set Tmp = arr(i)
                    Set arr(i) = arr(j)
                    Set arr(j) = Tmp
                Else
                    Tmp = arr(i)
                    arr(i) = arr(j)
                    arr(j) = Tmp
                End If
            End If
        Next j
    Next i
    SortZ = arr
End Function
