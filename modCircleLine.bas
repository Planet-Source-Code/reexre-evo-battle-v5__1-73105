Attribute VB_Name = "modCircleLine"
Option Explicit



Public Function Line_Circle(ByVal LX1 As Single, ByVal LY1 As Single, _
                            ByVal LX2 As Single, ByVal LY2 As Single, _
                            ByVal CX As Single, ByVal CY As Single, _
                            ByVal Radius As Single, ByRef desx1, ByRef desy1, ByRef desx2, ByRef desy2) As Integer
    Dim A              As Single
    Dim x1             As Single
    Dim x2             As Single
    Dim Y1             As Single
    Dim Y2             As Single
    Dim dX             As Single
    Dim dY             As Single
    Dim dR             As Single
    Dim dRQ            As Single

    Dim D              As Single

    Dim point1_exist   As Boolean
    Dim point2_exist   As Boolean
    Dim points         As Integer    ' points of intersection

    Dim Intersection   As Boolean
    'Stop

    LX2 = LX2 - LX1
    LY2 = LY2 - LY1



    x1 = LX1 - CX
    x2 = (LX1 + LX2) - CX
    Y1 = LY1 - CY
    Y2 = (LY1 + LY2) - CY

    dX = x2 - x1
    dY = Y2 - Y1

    dRQ = (dX * dX + dY * dY)
    dR = Sqr(dRQ)
    D = x1 * Y2 - x2 * Y1

    If (Radius * Radius * dRQ - D * D) >= 0 Then Intersection = True


    If Intersection = True Then

        'If (Radius ^ 2 * dR ^ 2 - D ^ 2) < 0 Then
        If (Radius * Radius * dRQ - D * D) < 0 Then

            A = 0
        Else
            A = Sqr(Radius * Radius * dRQ - D * D)
        End If

        desx1 = (D * dY + My_Sgn(dY) * dX * A) / dRQ + CX
        desy1 = (-D * dX + Abs(dY) * A) / dRQ + CY
        desx2 = (D * dY - My_Sgn(dY) * dX * A) / dRQ + CX
        desy2 = (-D * dX - Abs(dY) * A) / dRQ + CY

    End If

    point1_exist = Point_Line(desx1, desy1, LX1, LY1, LX2, LY2)
    point2_exist = Point_Line(desx2, desy2, LX1, LY1, LX2, LY2)

    If point1_exist And point2_exist Then
        points = 2
    Else
        points = 0
        If point1_exist Then points = 1
        If point2_exist Then points = 1: desx1 = desx2: desy1 = desy2
    End If

    Line_Circle = points

End Function

Private Function My_Sgn(X) As Integer

    If X < 0 Then
        My_Sgn = -1
    Else
        My_Sgn = 1
    End If

End Function

Private Function Point_Line(ByVal X, ByVal Y, ByVal pxn, ByVal pyn, ByVal pxt, ByVal pyt) As Boolean
    Dim t1             As Single
    Dim t2             As Single
    Dim op             As Boolean
    Dim T              As Single

    'pxt = pxt - pxn
    'pyt = pyt - pyn

    If pxt = 0 Then
        T = (Y - pyn) / pyt
        If T <= 1 And T >= 0 And X = pxn Then op = True
    End If

    If pyt = 0 Then
        T = (X - pxn) / pxt
        If T <= 1 And T >= 0 And Y = pyn Then op = True
    End If

    If pxt <> 0 And pyt <> 0 Then
        t1 = (X - pxn) / pxt
        t2 = (Y - pyn) / pyt
        If Abs(t1 - t2) <= 0.001 And t1 <= 1 And t1 >= 0 Then op = True
    End If

    Point_Line = op

End Function


Public Function Distance(x1, Y1, x2, Y2) As Single
    Dim dX             As Single
    Dim dY             As Single
    dX = x2 - x1
    dY = Y2 - Y1
    Distance = Sqr(dX * dX + dY * dY)
End Function
Public Function DistanceSQ(x1, Y1, x2, Y2) As Single
    Dim dX             As Single
    Dim dY             As Single
    dX = x2 - x1
    dY = Y2 - Y1
    DistanceSQ = (dX * dX + dY * dY)
End Function
