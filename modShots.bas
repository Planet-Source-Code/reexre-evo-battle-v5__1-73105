Attribute VB_Name = "modShots"
Option Explicit

Public Type tShot
    X                  As Single
    Y                  As Single
    vX                 As Single
    vY                 As Single
    Enabled            As Boolean
    TE                 As Long

End Type

Public Const ShotSpeed As Single = 3.5    '3.5     '4    ' 4.5

Public Sho(1 To 2, NH, MaxShots) As tShot

Public Function FireShot(Team, wI, TE)
    Dim I              As Long
    Dim wF             As Long
    Dim DoShot         As Boolean
    Dim VV             As Single
    Dim vvX            As Single
    Dim vvY            As Single
    Dim cosA           As Single
    Dim SinA           As Single

    DoShot = False
    I = 1
    Do
        If Sho(Team, wI, I).Enabled = False Then
            wF = I
            I = MaxShots
            DoShot = True
        End If
        I = I + 1
    Loop While I <= MaxShots

    If DoShot Then
        'Stop

        With Sho(Team, wI, wF)

            If Team = 1 Then

                .X = H1(wI).PosX
                .Y = H1(wI).PosY
                cosA = Cos(H1(wI).ANG)
                SinA = Sin(H1(wI).ANG)
                Projection H1(wI).vX, H1(wI).vY, cosA, SinA, vvX, vvY
                .vX = vvX + cosA * ShotSpeed
                .vY = vvY + SinA * ShotSpeed
                .TE = TE
                .Enabled = True
                H1(wI).AvailShots = H1(wI).AvailShots - 1
                H1(wI).LastShotTime = CurTime
            Else
                .X = H2(wI).PosX
                .Y = H2(wI).PosY
                cosA = Cos(H2(wI).ANG)
                SinA = Sin(H2(wI).ANG)
                Projection H2(wI).vX, H2(wI).vY, cosA, SinA, vvX, vvY
                .vX = vvX + cosA * ShotSpeed
                .vY = vvY + SinA * ShotSpeed
                .Enabled = True
                .TE = TE
                H2(wI).AvailShots = H2(wI).AvailShots - 1
                H2(wI).LastShotTime = CurTime
            End If

        End With
    End If
End Function


Public Sub InitShots()
    Dim I              As Long
    Dim wI             As Long
    Dim wS             As Long


    For I = 1 To 2
        For wI = 1 To NH
            For wS = 1 To MaxShots

                Sho(I, wI, wS).Enabled = False
            Next
        Next
    Next
End Sub

Public Sub MoveAndDrawShots(pHDC As Long, ZOOM As Single)
    Dim I              As Long
    Dim wI             As Long
    Dim wS             As Long
    Dim x1             As Long
    Dim y1             As Long
    Dim x2             As Long
    Dim y2             As Long


    For I = 1 To 2
        For wI = 1 To NH
            For wS = 1 To MaxShots

                With Sho(I, wI, wS)
                    If .Enabled Then
                        x1 = XtoScreen(.X)
                        y1 = YtoScreen(.Y)
                        
                        '  Stop
                        .vY = .vY + Gravity
                        .X = .X + .vX
                        .Y = .Y + .vY

                       
                        'SetPixel pHDC, x, Y, IIf(I = 1, vbRed, vbCyan)
                        x2 = XtoScreen(.X)
                        y2 = YtoScreen(.Y)
                        ' SetPixel pHDC, x, Y, IIf(I = 1, vbRed, vbCyan)
                        If IsInsideScreen(x2, y2) Then FastLine pHDC, x1, y1, x2, y2, 2, IIf(I = 1, vbRed, vbCyan)

                        If .X < 0 Or .Y < 0 Or .X > MaxX Or .Y > MaxY Then

                            .Enabled = False
                            If I = 1 Then
                                H1(wI).AvailShots = H1(wI).AvailShots + 1
                            Else
                                H2(wI).AvailShots = H2(wI).AvailShots + 1

                            End If

                        End If
                    End If
                End With
            Next
        Next
    Next


End Sub


Public Function NearestShot(Team, wI, ByRef RetwE, ByRef RetwS) As Single
    Dim Dmin           As Single
    Dim D              As Single

    Dim RetTeam        As Long
    Dim E              As Long
    Dim S              As Long


    Dmin = 9999999999#
    If Team = 1 Then
        RetTeam = 2
        For E = 1 To NH
            For S = 1 To MaxShots
                If Sho(RetTeam, E, S).Enabled Then
                    D = DistanceSQ(H1(wI).PosX, H1(wI).PosY, Sho(RetTeam, E, S).X, Sho(RetTeam, E, S).Y)
                    If D < Dmin Then
                        Dmin = D
                        RetwE = E
                        RetwS = S
                    End If
                End If
            Next
        Next
    Else
        RetTeam = 1
        For E = 1 To NH
            For S = 1 To MaxShots
                If Sho(RetTeam, E, S).Enabled Then
                    D = DistanceSQ(H2(wI).PosX, H2(wI).PosY, Sho(RetTeam, E, S).X, Sho(RetTeam, E, S).Y)
                    If D < Dmin Then
                        Dmin = D
                        RetwE = E
                        RetwS = S
                    End If
                End If
            Next
        Next


    End If

    NearestShot = Sqr(Dmin)

    If NearestShot = 9999999999# Then NearestShot = MaxX

End Function


