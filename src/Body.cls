VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Body"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public World As World

Public Index As Long
Public Pos As New Vec2
Public Vel As New Vec2
Public Size As New Vec2
Public BodyType As BodyType
Public GravityScale As New Vec2
Public FreezePos As Boolean
Public CollisionEnabled As Boolean

Private pContacts(3) As Body

Private pMass As Double
Private pInvMass As Double

Private pElapsedTime As Double
Private pForce As New Vec2

Public Event OnCollide(Body As Body, Side As CollisionSide)
Public Event OnAvert(Body As Body, Side As CollisionSide)

Private Sub Class_Initialize()
    CollisionEnabled = True
    GravityScale.X = 1
    GravityScale.Y = 1
End Sub

Public Property Get Mass() As Double
    Mass = pMass
End Property

Public Property Let Mass(NewValue As Double)
    pMass = NewValue
    If Mass = 0 Then
        pInvMass = 0
    Else
        pInvMass = 1 / Mass
    End If
End Property

Public Property Get InvMass() As Double
    InvMass = pInvMass
End Property

Public Sub Update(ElapsedTime As Double)
    If BodyType = BodyDynamic And FreezePos = False Then
        pElapsedTime = ElapsedTime
        Vel.X = Vel.X + (pForce.X * pInvMass + World.Gravity.X * GravityScale.X) * pElapsedTime
        Vel.Y = Vel.Y + (pForce.Y * pInvMass + World.Gravity.Y * GravityScale.Y) * pElapsedTime
    End If
End Sub

Public Sub ApplyImpulse(Impulse As Vec2)
    If FreezePos Then Exit Sub
    If Mass = 0 Then
        Vel.X = Vel.X + Impulse.X
        Vel.Y = Vel.Y + Impulse.Y
    Else
        Vel.X = Vel.X + Impulse.X * pInvMass
        Vel.Y = Vel.Y + Impulse.Y * pInvMass
    End If
End Sub

Public Sub ApplyForce(Force As Vec2)
    pForce.X = pForce.X + Force.X
    pForce.Y = pForce.Y + Force.Y
End Sub

Public Sub ClearForces()
    pForce.X = 0
    pForce.Y = 0
End Sub

Public Sub SetContact(Index As Integer, Body As Body)
    Dim Sides(3) As CollisionSide
    Sides(0) = SideTop
    Sides(1) = SideLeft
    Sides(2) = SideBottom
    Sides(3) = SideRight
        
    If Body Is Nothing Then
        If Not pContacts(Index) Is Nothing Then RaiseEvent OnAvert(pContacts(Index), Sides(Index))
        Set pContacts(Index) = Body
    Else
        Set pContacts(Index) = Body
        RaiseEvent OnCollide(Body, Sides(Index))
    End If
End Sub

Public Function GetContact(Index As Integer) As Body
    Set GetContact = pContacts(Index)
End Function
