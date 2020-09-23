VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHuman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name        As String
Public SurName     As String
Public Gender      As Boolean



Public PosX        As Single
Public PosY        As Single
Public vX          As Single
Public vY          As Single

Public ANG         As Single
Public ANGVel      As Single
Public R           As Single

Public ACCLeft         As Single
Public ACCRight         As Single

Public VxHit       As Single
Public VyHit       As Single

Public Walked      As Single


Public EnemyAng    As Single
'Public EnemyDist As Single

Public ReChaAng    As Single
Public ReChaDist   As Single

Public EnemyShotAng As Single
Public EnemyShotDist As Single


Public LastShotTime As Long
Public AvailShots  As Long

Public HitTIME     As Long
Public CantShot    As Boolean

Public EnemyToAttack As Long

Public Hitten      As Boolean


Private SciaX(0 To SciaLen) As Single
Private SciaY(0 To SciaLen) As Single
Private SciaI(0 To SciaLen) As Single
Private SciaC1(0 To SciaLen) As Long
Private SciaC2(0 To SciaLen) As Long

Public SciaPos     As Long

Public BRtoRUN     As Long


Public Function SPEED() As Single
    SPEED = Sqr(vX * vX + vY * vY)
End Function
Public Function DirMoving() As Single
    DirMoving = Atan2(vX, vY)
End Function

Public Sub Move(ByRef Scia As Boolean)
    Dim A          As Single
    Dim A2         As Single
    Dim D          As Single
    Dim I          As Long
    Dim cosA       As Single
    Dim SinA       As Single

    Dim xL         As Single
    Dim xR         As Single
    Dim yL         As Single
    Dim yR         As Single

    Dim NewPosX    As Single
    Dim NewposY    As Single

    Dim InvSciaLenPuls1 As Single
    
InvSciaLenPuls1 = 1 / (SciaLen + 1)
    If Scia Then
        'If CurTime Mod 8 = 0 Then
        For I = 0 To SciaLen
            SciaI(I) = SciaI(I) - (150 * InvSciaLenPuls1)
            If SciaI(I) < 0 Then SciaI(I) = 0
            SciaC1(I) = RGB(SciaI(I), 0, 0)
            SciaC2(I) = RGB(0, SciaI(I), SciaI(I))
        Next
        SciaPos = (SciaPos + 1) Mod (SciaLen + 1)
        SciaX(SciaPos) = PosX
        SciaY(SciaPos) = PosY
        SciaI(SciaPos) = 255
        SciaC1(SciaPos) = RGB(255, 0, 0)
        SciaC2(SciaPos) = RGB(0, 255, 255)
        'End If
    End If


    If CurTime - HitTIME < FrozenTime Then
        CantShot = True
        'Exit Sub
        'ACC = -ACC
        ACCLeft = ACCLeft * 0.25
        ACCRight = ACCRight * 0.25
    Else
        CantShot = False
    End If



    'PosX = PosX + Cos(ANG) * Vel + VxHit
    'PosY = PosY + Sin(ANG) * Vel + VyHit
    '**********************************************************
    '    If CurTime - HitTIME >= InvisibleTime Then
    '        vX = vX + (Cos(ANG) * ACC + VxHit) * 0.1
    '        vY = vY + (Sin(ANG) * ACC + VyHit) * 0.1
    '        Hitten = False
    '    Else
    '        vX = vX + (Cos(ANG) * ACC + VxHit) * 0.05
    '        vY = vY + (Sin(ANG) * ACC + VyHit) * 0.05
    '        Hitten = True
    '    End If
    '**********************************************************



    cosA = Cos(ANG - PIh)
    SinA = Sin(ANG - PIh)

    xL = PosX + cosA * R
    yL = PosY + SinA * R
    xR = PosX - cosA * R
    yR = PosY - SinA * R

    cosA = Cos(ANG)
    SinA = Sin(ANG)

    xL = xL + cosA * ACCLeft
    yL = yL + SinA * ACCLeft
    xR = xR + cosA * ACCRight
    yR = yR + SinA * ACCRight



    ANGVel = ANGVel + AngleDiff(ANG, Atan2(xR - xL, yR - yL) - PIh) * 0.15 '0.125
    NewPosX = (xL + xR) * 0.5
    NewposY = (yL + yR) * 0.5

    If CurTime - HitTIME >= InvisibleTime Then
        vX = vX + (NewPosX - PosX) * 0.09 + VxHit
        vY = vY + (NewposY - PosY) * 0.09 + VyHit
        Hitten = False
    Else
        vX = vX + (NewPosX - PosX) * 0.01 + VxHit
        vY = vY + (NewposY - PosY) * 0.01 + VyHit
        Hitten = True
    End If

    
    If CurTime - HitTIME >= FrozenTime Then
        ANG = ANG + ANGVel
    Else
        ANG = ANG + ANGVel * 0.25
    End If
    
    If ANG > PI2 Then ANG = ANG - PI2
    If ANG < -PI2 Then ANG = ANG + PI2
    
    ANGVel = ANGVel * 0.95 '0.95 '0.8
    '**********************************************************

    PosX = PosX + vX
    PosY = PosY + vY
    vX = vX * 0.98 ' 0.99
    vY = vY * 0.98 + Gravity

    D = vX * vX + vY * vY
    
    If D > MaxHSpeedSq Then
    
        D = Sqr(D)
        vX = (vX / D) * MaxHSpeed
        vY = (vY / D) * MaxHSpeed
    End If



    VxHit = VxHit * 0 '.75
    VyHit = VyHit * 0 '.75


    If PosX < 0 Then PosX = MaxX
    'If PosX < 0 Then PosX = 1: vX = -vX    ': ANG = ANG - PI

    'If PosY < 0 Then PosY = MaxY
    If PosY < 0 Then PosY = 1: vY = -vY    ': ANG = ANG - PI

    If PosX > MaxX Then PosX = 0
    'If PosX > MaxX Then PosX = MaxX - 1: vX = -vX    ': ANG = ANG - PI

    'If PosY > MaxY Then PosY = 0
    If PosY > MaxY Then PosY = MaxY - 1: vY = -vY    ': ANG = ANG - PI


End Sub

Public Sub DRAW(ByVal pHDC As Long, fit As Single, Team As Long, ByRef Scia As Boolean, Optional IsTheBest As Boolean = False)

    Dim I          As Long
    Dim SciaRad    As Long
    Dim SciaRad2   As Long

    Dim A          As Single
    Dim C          As Long
    Dim C2         As Long
    Dim S          As Integer

    Dim cosA       As Single
    Dim SinA       As Single
    Dim CosA1      As Single
    Dim SinA1      As Single
    Dim CosA2      As Single
    Dim SinA2      As Single

    Dim X0         As Long
    Dim Y0         As Long
    Dim x1         As Long
    Dim y1         As Long
    Dim x2         As Long
    Dim y2         As Long

    Dim pa(0 To 2) As POINTAPI

    Dim ScreenX    As Long
    Dim ScreenY    As Long

    Dim XX         As Single
    Dim YY         As Single


    ScreenX = XtoScreen(PosX)
    ScreenY = YtoScreen(PosY)

    If IsInsideScreen(ScreenX, ScreenY) Then
    
              S = IIf(Team = 1, 1, -1)
              
        If Scia Then
            SciaRad = SciaR * ZOOM
            SciaRad2 = SciaRad * 2
            If S = 1 Then
                For I = SciaLen To 0 Step -1
                    MyCircle pHDC, XtoScreen(SciaX(I)), YtoScreen(SciaY(I)), SciaRad, SciaRad2, SciaC1(I)
                Next
            Else
                For I = SciaLen To 0 Step -1
                    MyCircle pHDC, XtoScreen(SciaX(I)), YtoScreen(SciaY(I)), SciaRad, SciaRad2, SciaC2(I)
                Next
            End If
        End If

        If IsTheBest Then
            C2 = IIf(Team = 1, RGB(255, 200, 100), RGB(0, 255, 100))
        Else
            C2 = IIf(Team = 1, vbRed, vbCyan)
            'C2 = IIf(team = 1, RGB(255, 0, 0), RGB(0, 255, 255))
        End If

  


        cosA = Cos(ANG)
        SinA = Sin(ANG)
        CosA1 = -Cos(ANG - 1)
        SinA1 = -Sin(ANG - 1)
        CosA2 = -Cos(ANG + 1)
        SinA2 = -Sin(ANG + 1)

        'Stop

        '    pa(0).x = ZOOM * (PosX + cosA * r * 2)
        '    pa(0).y = ZOOM * (PosY + SinA * r * 2)
        '    pa(1).x = ZOOM * (PosX + CosA1 * r)
        '    pa(1).y = ZOOM * (PosY + SinA1 * r)
        '    pa(2).x = ZOOM * (PosX + CosA2 * r)
        '    pa(2).y = ZOOM * (PosY + SinA2 * r)

        pa(0).X = XtoScreen(PosX + cosA * R * 2)
        pa(0).Y = YtoScreen(PosY + SinA * R * 2)
        pa(1).X = XtoScreen(PosX + CosA1 * R)
        pa(1).Y = YtoScreen(PosY + SinA1 * R)
        pa(2).X = XtoScreen(PosX + CosA2 * R)
        pa(2).Y = YtoScreen(PosY + SinA2 * R)


        'FastLine pHDC, X0, Y0, X1, Y1, 1, C2
        'FastLine pHDC, X1, Y1, X2, Y2, 1, C2
        'FastLine pHDC, X2, Y2, X0, Y0, 1, C2

        'bec
        'FastLine pHDC, PosX * Zoom, PosY * Zoom, (PosX + CosA * r * 2) * Zoom, (PosY + SinA * r * 2) * Zoom, 1, C2
        'fit
        FastLine pHDC, XtoScreen(PosX), YtoScreen(PosY + 25), XtoScreen(PosX + S * fit * 0.15), YtoScreen(PosY + 25), 1, C2
        frmM.PIC.FillColor = C2
        Polygon pHDC, pa(0), 3
        'acc

        '        FastLine pHDC, XtoScreen(PosX), YtoScreen(PosY), XtoScreen(PosX - 25 * CosA * ACC), YtoScreen(PosY - 25 * SinA * ACC), 2, RGB(255, 180, 0)
        XX = PosX + CosA2 * R
        YY = PosY + SinA2 * R
        FastLine pHDC, XtoScreen(XX), YtoScreen(YY), _
                 XtoScreen(XX - 50 * cosA * ACCLeft), YtoScreen(YY - 50 * SinA * ACCLeft), 2, RGB(255, 180, 0)
        XX = PosX + CosA1 * R
        YY = PosY + SinA1 * R
        FastLine pHDC, XtoScreen(XX), YtoScreen(YY), _
                 XtoScreen(XX - 50 * cosA * ACCRight), YtoScreen(YY - 50 * SinA * ACCRight), 2, RGB(255, 180, 0)


        If CantShot Then
            FastLine pHDC, XtoScreen(PosX - 12), YtoScreen(PosY - 12), XtoScreen(PosX + 12), YtoScreen(PosY + 12), 1, C2
            FastLine pHDC, XtoScreen(PosX - 12), YtoScreen(PosY + 12), XtoScreen(PosX + 12), YtoScreen(PosY - 12), 1, C2
        End If




    End If

End Sub



