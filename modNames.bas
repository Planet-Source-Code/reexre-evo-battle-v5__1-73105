Attribute VB_Name = "modNames"
Option Explicit

Const V = "AEIOUaeiou"


Public Function GenName(ByVal lengMin, ByVal lengMax) As String
    Dim I              As Integer
    Dim I2             As Integer

    Dim S              As String
    Dim C              As String

    If lengMax < lengMin Then lengMax = lengMin

    lengMax = lengMax - lengMin
    If lengMax < 0 Then lengMax = 0


    If Rnd < 0.5 Then
        I2 = Int(Rnd * 5) + 1:
        C = Mid$(V, I2, 1)
    Else
        Do
            I2 = Int(Rnd * 25)
            I2 = Asc("A") + I2
            C = Chr$(I2)
        Loop While InStr(1, V, C) <> 0
    End If
    S = S + C

    For I = 2 To lengMin + Int(Rnd * (lengMax + 1))
        If InStr(1, V, Mid$(S, Len(S))) = 0 And Rnd < 0.8 Then
            I2 = Int(Rnd * 5) + 1:
            C = Mid$(V, I2, 1)
        Else
            Do
                I2 = Int(Rnd * 25)
                I2 = Asc("A") + I2
                C = Chr$(I2)
            Loop While InStr(1, V, C) <> 0
        End If
        S = S + LCase(C)


    Next


    GenName = S




End Function

