Attribute VB_Name = "modWorld"
Option Explicit


Public GA1             As New SimplyGA2
Public GA2             As New SimplyGA2

Public Const PI2 As Single = 6.28318530717959
Public Const PI As Single = 3.14159265358979
Public Const PIh As Single = 1.5707963267949

Public H1()            As New clsHuman
Public H2()            As New clsHuman

Public Const NH        As Long = 7 '8 '12    '14 ' 12    '15 '12


Public F()             As New clsRecharge
Public NF              As Long




Public Const StartFit As Single = 10000
Public Const MaxFitDist As Single = 600   ' 400


Public Const MaxShots  As Long = 1 '1    '2    '2    ' 3
Public Const ShotDelay As Long = 25 '25    '30 ' 20 '6    '80    '50    '8    '15


Public EnemyDist(1 To NH, 1 To NH) As Single

Public ShotDist1(1 To MaxShots) As Single
Public ShotDist2(1 To MaxShots) As Single

Public CurTime         As Long

Public Const MaxHSpeed As Single = 1.8    '2 '2.5
Public Const MaxHSpeedSq As Single = 1.8 * 1.8


Public Const FrozenTime As Long = 200 '100    '100
Public Const InvisibleTime As Long = 250 '200    '100*2

Public Const HitPTS    As Long = 40    '38 '55 '38 '34 '38 '55
Public Const HittenPTS As Long = -35 '-30 '-25 '34 '20 '34    '34 '
Public Const FirePTS   As Long = -2

Public Const MatchLenght As Long = 5000 ' 4000 '3000    '3000    '2600

Public Const Gravity   As Single = 0    ' 0.015

Public Const NofInputs As Long = 11

Public Const DefaultRadius As Single = 8

Public Const SciaR As Single = 1.8
Public Const SciaFreq As Long = 6
Public Const SciaLen As Long = 50 ' 25

Public Const MaxEYEdist As Single = 700


Public Function Atan2(X As Single, Y As Single) As Single
    If X Then
        Atan2 = -PI + Atn(Y / X) - (X > 0) * PI
    Else
        Atan2 = -PIh - (Y > 0) * PI
    End If

    ' While Atan2 < 0: Atan2 = Atan2 + Pi2: Wend
    ' While Atan2 > Pi2: Atan2 = Atan2 - Pi2: Wend


End Function
Public Sub ComputeEnemiesDistances()
    Dim I              As Long
    Dim J              As Long
    For J = 1 To NH
        For I = 1 To NH
            EnemyDist(I, J) = DistanceSQ(H1(I).PosX, H1(I).PosY, H2(J).PosX, H2(J).PosY)
        Next
    Next

    'Stop

End Sub

Public Function NearestEnemy(Team, wI) As Long
    Dim D              As Single
    Dim Dmin           As Single
    Dim I              As Long
    Dim J              As Long


    NearestEnemy = IIf(Team = 1, GA2.STAT_GenerBestFitINDX, GA1.STAT_GenerBestFitINDX)


    Dmin = 99999999999#
    If Team = 1 Then
        For J = 1 To NH
            '            If CurTime - H2(J).HitTIME >= InvisibleTime Then
            If H2(J).Hitten = False Then
                H2(J).R = DefaultRadius
                If EnemyDist(wI, J) < Dmin Then
                    Dmin = EnemyDist(wI, J)
                    NearestEnemy = J

                End If
            Else
                H2(J).R = DefaultRadius * 0.4 '6
            End If

        Next
    Else
        For I = 1 To NH
            '            If CurTime - H1(I).HitTIME >= InvisibleTime Then
            If H1(I).Hitten = False Then

                H1(I).R = DefaultRadius
                If EnemyDist(I, wI) < Dmin Then
                    Dmin = EnemyDist(I, wI)
                    NearestEnemy = I

                End If
            Else
                H1(I).R = DefaultRadius * 0.6
            End If
        Next

    End If


End Function

Public Function ComputeAngle(Team, wI, wEnem) As Single
    Dim A              As Single

    If Team = 1 Then

        A = Atan2(H1(wI).PosX - H2(wEnem).PosX, H1(wI).PosY - H2(wEnem).PosY)
        ComputeAngle = AngleDiff01(H1(wI).ANG, A)
    Else
        A = Atan2(H2(wI).PosX - H1(wEnem).PosX, H2(wI).PosY - H1(wEnem).PosY)

        ComputeAngle = AngleDiff01(H2(wI).ANG, A)
    End If
End Function

Public Function AngleDiff01(A1 As Single, A2 As Single) As Single
'double difference = secondAngle - firstAngle;
'while (difference < -180) difference += 360;
'while (difference > 180) difference -= 360;
'return difference;

    AngleDiff01 = A2 - A1
    While AngleDiff01 < -PI
        AngleDiff01 = AngleDiff01 + PI2
    Wend
    While AngleDiff01 > PI
        AngleDiff01 = AngleDiff01 - PI2
    Wend



    '''' this is to have values between 0 and 1
    AngleDiff01 = AngleDiff01 + PI
    AngleDiff01 = AngleDiff01 / (PI2)


End Function
Public Function AngleDiff(A1 As Single, A2 As Single) As Single
'double difference = secondAngle - firstAngle;
'while (difference < -180) difference += 360;
'while (difference > 180) difference -= 360;
'return difference;

    AngleDiff = A2 - A1
    While AngleDiff < -PI
        AngleDiff = AngleDiff + PI2
    Wend
    While AngleDiff > PI
        AngleDiff = AngleDiff - PI2
    Wend



End Function

Public Function AngToLeftRight(A0to1 As Single, ByRef AL As Single, AR As Single)
    If A0to1 < 0.5 Then

        AR = 0: AL = (0.5 - A0to1) * 2
    Else


        AL = 0: AR = (A0to1 - 0.5) * 2
    End If

End Function


Public Sub NotOverlap()
    Dim I1             As Long
    Dim I2             As Long
    Dim dX             As Single
    Dim dY             As Single
    Dim L              As Single
    Dim R              As Single

    Dim R2             As Single

    R = DefaultRadius * 3 '2.5        '*2 pure (correct)
    R2 = R * R


    For I1 = 1 To NH - 1
        For I2 = I1 + 1 To NH
            dX = H1(I2).PosX - H1(I1).PosX

            If Abs(dX) < R Then
                dY = H1(I2).PosY - H1(I1).PosY

                If Abs(dY) < R Then

                    L = (dX * dX + dY * dY)
                    If L < R2 Then
                        L = Sqr(L)
                        If L <> 0 Then
                        dX = dX / L
                        dY = dY / L
                        dX = dX * (R - L) * 0.5
                        dY = dY * (R - L) * 0.5
                        H1(I2).PosX = H1(I2).PosX + dX
                        H1(I2).PosY = H1(I2).PosY + dY
                        H1(I1).PosX = H1(I1).PosX - dX
                        H1(I1).PosY = H1(I1).PosY - dY
                        End If
                    End If
                End If
            End If

            dX = H2(I2).PosX - H2(I1).PosX
            If Abs(dX) < R Then
                dY = H2(I2).PosY - H2(I1).PosY
                If Abs(dY) < R Then
                    L = (dX * dX + dY * dY)
                    If L < R2 Then
                        L = Sqr(L)
                        If L <> 0 Then
                        dX = dX / L
                        dY = dY / L
                        dX = dX * (R - L) * 0.5
                        dY = dY * (R - L) * 0.5
                        H2(I2).PosX = H2(I2).PosX + dX
                        H2(I2).PosY = H2(I2).PosY + dY
                        H2(I1).PosX = H2(I1).PosX - dX
                        H2(I1).PosY = H2(I1).PosY - dY
                        End If
                    End If

                End If
            End If

        Next
    Next

    For I1 = 1 To NH
        For I2 = 1 To NH
            dX = H2(I2).PosX - H1(I1).PosX
            If Abs(dX) < R Then
                dY = H2(I2).PosY - H1(I1).PosY
                If Abs(dY) < R Then

                    L = (dX * dX + dY * dY)
                    If L < R2 Then
                        L = Sqr(L)
                        If L <> 0 Then
                        dX = dX / L
                        dY = dY / L
                        dX = dX * (R - L) * 0.5
                        dY = dY * (R - L) * 0.5
                        H2(I2).PosX = H2(I2).PosX + dX
                        H2(I2).PosY = H2(I2).PosY + dY
                        H1(I1).PosX = H1(I1).PosX - dX
                        H1(I1).PosY = H1(I1).PosY - dY
                        Else
                        H1(I1).PosX = H1(I1).PosX + (Rnd - 0.5) * 0.01
                        H1(I1).PosY = H1(I1).PosY + (Rnd - 0.5) * 0.01
                        End If
                        
                    End If
                End If
            End If

        Next
    Next


End Sub



' ' def projection(self, vector):
'    '        k = (self.dot(vector)) / vector.length()
'    '        return k * vector.unit()
'
'    Dim K As Single
'
'    Set Projection = New cls2DVector
'
'    K = (x * V.x + Y * V.Y) / Sqr(V.x * V.x + V.Y * V.Y)
'    V.Normaliz
'    V.MUL K
'
'    Set Projection = V

Public Sub Projection(x1, y1, toX, toY, ByRef X3, ByRef Y3)
    Dim K              As Single
    Dim L              As Single

    L = Sqr(toX * toX + toY * toY)

    K = ((x1 * toX + y1 * toY) / L)

    X3 = (toX / L) * K
    Y3 = (toY / L) * K

End Sub

