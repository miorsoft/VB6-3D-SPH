Attribute VB_Name = "mVectors"
Option Explicit

Public Const PI   As Double = 3.14159265358979
Public Const InvPI As Double = 1 / 3.14159265358979

Public Const PIh  As Double = 1.5707963267949
Public Const PI2  As Double = 6.8318530717959


Private Const INV2 As Double = 0.5
Private Const INV9 As Double = 0.111111111111111
Private Const INV72 As Double = 1.38888888888889E-02
Private Const INV1008 As Double = 9.92063492063492E-04
Private Const INV30240 As Double = 3.30687830687831E-05

Public Type tVec3
    x             As Double
    Y             As Double
    z             As Double
End Type

Public Function Atan2(ByVal x As Double, ByVal Y As Double) As Double
    If x Then
        Atan2 = -PI + Atn(Y / x) - (x > 0!) * PI
    Else
        Atan2 = -PIh - (Y > 0!) * PI
    End If
End Function

Public Function Vec3(x As Double, Y As Double, z As Double) As tVec3
    Vec3.x = x
    Vec3.Y = Y
    Vec3.z = z
End Function

Public Function Length3(V As tVec3) As Double
    With V
        Length3 = Sqr(.x * .x + .Y * .Y + .z * .z)
    End With
End Function

Public Function Length32(V As tVec3) As Double
    With V
        Length32 = .x * .x + .Y * .Y + .z * .z
    End With
End Function



Public Function SUM3(v1 As tVec3, V2 As tVec3) As tVec3
    SUM3.x = v1.x + V2.x
    SUM3.Y = v1.Y + V2.Y
    SUM3.z = v1.z + V2.z
End Function



Public Function Normalize3(V As tVec3) As tVec3
    Dim D         As Double

    D = (V.x * V.x + V.Y * V.Y + V.z * V.z)
    If D Then
        D = 1# / Sqr(D)
        Normalize3.x = V.x * D
        Normalize3.Y = V.Y * D
        Normalize3.z = V.z * D
    End If

End Function

Public Function MUL3(V As tVec3, ByVal A As Double) As tVec3
    MUL3.x = V.x * A
    MUL3.Y = V.Y * A
    MUL3.z = V.z * A

End Function

Public Function DOT3(v1 As tVec3, V2 As tVec3) As Double

    DOT3 = (v1.x * V2.x) + _
           (v1.Y * V2.Y) + _
           (v1.z * V2.z)

End Function

Public Function CROSS3(A As tVec3, B As tVec3) As tVec3
    CROSS3.x = A.Y * B.z - A.z * B.Y
    CROSS3.Y = A.z * B.x - A.x * B.z
    CROSS3.z = A.x * B.Y - A.Y * B.x
End Function

Public Function DIFF3(v1 As tVec3, V2 As tVec3) As tVec3
    DIFF3.x = v1.x - V2.x
    DIFF3.Y = v1.Y - V2.Y
    DIFF3.z = v1.z - V2.z
End Function


Public Function Project3(V As tVec3, N As tVec3) As tVec3
    Dim DOT       As Double

    DOT = DOT3(V, N)
    Project3 = MUL3(N, DOT)

End Function





Public Function Rotate3(V As tVec3, XAxe As tVec3, YAxe As tVec3, ZAxe As tVec3) As tVec3

' TO TEST and ADJUST

    Rotate3.x = DOT3(V, XAxe)
    Rotate3.Y = DOT3(V, YAxe)
    Rotate3.z = DOT3(V, ZAxe)

End Function

''''******************************************************************
''''   TODO:
''''   http://paulbourke.net/geometry/rotate/
''''   http://paulbourke.net/geometry/rotate/source.c
''''******************************************************************
'''Public Function Rotate3Axe(P As tVec3, Axe As tVec3, theta As Double) As tVec3
'''    Dim L         As Double
'''    Dim CosTheta  As Double
'''    Dim SinTheta  As Double
'''
'''    'http://paulbourke.net/geometry/rotate/source.c
'''
'''    'Normalize AXE
'''    L = Sqr(Axe.x * Axe.x + Axe.Y * Axe.Y + Axe.z * Axe.z)
'''    If L <> 1 Then
'''        L = 1 / L
'''        Axe = MUL3(Axe, L)
'''    End If
'''
'''
''''       CosTheta = Cos(-theta)
''''       SinTheta = Sin(-theta)
'''
'''    'TODO :    MUST TO FIND LEFT HAND RULES   .... seems to works for Y axe, steel to test
'''    CosTheta = Cos(-theta)
'''    SinTheta = -Sin(-theta)
'''
'''    With Rotate3Axe
''''        .x = 0: .Y = 0: .z = 0
'''        .x = .x + (CosTheta + (1 - CosTheta) * Axe.x * Axe.x) * P.x
'''        .x = .x + ((1 - CosTheta) * Axe.x * Axe.Y - Axe.z * SinTheta) * P.Y
'''        .x = .x + ((1 - CosTheta) * Axe.x * Axe.z + Axe.Y * SinTheta) * P.z
'''
'''        .Y = .Y + ((1 - CosTheta) * Axe.x * Axe.Y + Axe.z * SinTheta) * P.x
'''        .Y = .Y + (CosTheta + (1 - CosTheta) * Axe.Y * Axe.Y) * P.Y
'''        .Y = .Y + ((1 - CosTheta) * Axe.Y * Axe.z - Axe.x * SinTheta) * P.z
'''
'''        .z = .z + ((1 - CosTheta) * Axe.x * Axe.z - Axe.Y * SinTheta) * P.x
'''        .z = .z + ((1 - CosTheta) * Axe.Y * Axe.z + Axe.x * SinTheta) * P.Y
'''        .z = .z + (CosTheta + (1 - CosTheta) * Axe.z * Axe.z) * P.z
'''
'''    End With
'''
'''End Function



Public Function Rotate3xz(V As tVec3, XAxe As tVec3) As tVec3


' POSITIVE Z = Left Side (front in POSITIVE X , Up Negative Y)

'    Rotate3xz.x = DOT3(V, XAxe)    'Parallel
'    Rotate3xz.Y = V.Y * 1
'    Rotate3xz.z = DOT3(V, Vec3(XAxe.z, 0, -XAxe.x))    'Perpendicular clockwise
'   This (up) is a change of coordinete system


' POSITIVE Z = RIGHT Side (front in POSITIVE X , Up Negative Y)

' Corrected ONE:
'http://www.vbforums.com/showthread.php?874965-Rotation-using-DOT-product

    Rotate3xz.x = DOT3(V, Vec3(XAxe.x, 0, -XAxe.z))
    Rotate3xz.Y = V.Y * 1
    Rotate3xz.z = DOT3(V, Vec3(XAxe.z, 0, XAxe.x))

End Function




Public Function RayPlaneIntersect(rayVector As tVec3, rayPoint As tVec3, PlaneNormal As tVec3, planePoint As tVec3) As tVec3
'https://rosettacode.org/wiki/Find_the_intersection_of_a_line_with_a_plane#C.23

    Dim Diff      As tVec3
    Dim prod1     As Double
    Dim prod2     As Double
    Dim prod3     As Double

    Diff = DIFF3(rayPoint, planePoint)
    prod1 = DOT3(Diff, PlaneNormal)
    prod2 = DOT3(rayVector, PlaneNormal)
    prod3 = prod1 / prod2
    RayPlaneIntersect = DIFF3(rayPoint, MUL3(rayVector, prod3))

End Function





'Public Function fastEXP(ByVal x As Double) As Double
''https://en.wikipedia.org/wiki/Pad%C3%A9_approximant
'    Dim X2        As Double
'    Dim X3        As Double
'    Dim X4        As Double
'    Dim X5        As Double
'
'
'    If x < 5! Then
'
'        If x < -7! Then fastEXP = 0!: Exit Function
'
'        X2 = x * x
'        X3 = X2 * x
'        X4 = X3 * x
'        X5 = X4 * x
'
'        fastEXP = (1! + INV2 * x + INV9 * X2 + INV72 * X3 + INV1008 * X4 + INV30240 * X5) / _
         '                  (1! - INV2 * x + INV9 * X2 - INV72 * X3 + INV1008 * X4 - INV30240 * X5)
'
'    Else
'        fastEXP = Exp(x)
'    End If
'
'
'End Function

