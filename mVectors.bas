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
    X             As Double
    Y             As Double
    Z             As Double
End Type


Public Type tRotor3
    ' scalar part
    a             As Double    '=1
    ' bivector part
    b01           As Double
    b02           As Double
    b12           As Double
End Type



Public Function Atan2(ByVal X As Double, ByVal Y As Double) As Double
    If X Then
        Atan2 = -PI + Atn(Y / X) - (X > 0!) * PI
    Else
        Atan2 = -PIh - (Y > 0!) * PI
    End If
End Function

Public Function Vec3(X As Double, Y As Double, Z As Double) As tVec3
    Vec3.X = X
    Vec3.Y = Y
    Vec3.Z = Z
End Function

Public Function Length3(V As tVec3) As Double
    With V
        Length3 = Sqr(.X * .X + .Y * .Y + .Z * .Z)
    End With
End Function

Public Function Length32(V As tVec3) As Double
    With V
        Length32 = .X * .X + .Y * .Y + .Z * .Z
    End With
End Function



Public Function SUM3(v1 As tVec3, V2 As tVec3) As tVec3
    SUM3.X = v1.X + V2.X
    SUM3.Y = v1.Y + V2.Y
    SUM3.Z = v1.Z + V2.Z
End Function



Public Function Normalize3(V As tVec3) As tVec3
    Dim D         As Double

    D = (V.X * V.X + V.Y * V.Y + V.Z * V.Z)
    If D Then
        D = 1# / Sqr(D)
        Normalize3.X = V.X * D
        Normalize3.Y = V.Y * D
        Normalize3.Z = V.Z * D
    End If

End Function

Public Function MUL3(V As tVec3, ByVal a As Double) As tVec3
    MUL3.X = V.X * a
    MUL3.Y = V.Y * a
    MUL3.Z = V.Z * a

End Function

Public Function DOT3(v1 As tVec3, V2 As tVec3) As Double

    DOT3 = (v1.X * V2.X) + _
           (v1.Y * V2.Y) + _
           (v1.Z * V2.Z)

End Function

Public Function CROSS3(a As tVec3, B As tVec3) As tVec3
    CROSS3.X = a.Y * B.Z - a.Z * B.Y
    CROSS3.Y = a.Z * B.X - a.X * B.Z
    CROSS3.Z = a.X * B.Y - a.Y * B.X

End Function



'// Wedge product
Public Function WEDGE3(a As tVec3, B As tVec3) As tVec3    'BiVector

    WEDGE3.X = a.X * B.Y - a.Y * B.X    ', // XY
    WEDGE3.Y = a.X * B.Z - a.Z * B.X    ', // XZ
    WEDGE3.Z = a.Y * B.Z - a.Z * B.Y    '  // YZ


End Function


Public Function DIFF3(v1 As tVec3, V2 As tVec3) As tVec3
    DIFF3.X = v1.X - V2.X
    DIFF3.Y = v1.Y - V2.Y
    DIFF3.Z = v1.Z - V2.Z
End Function


Public Function Project3(V As tVec3, N As tVec3) As tVec3
    Dim DOT       As Double

    DOT = DOT3(V, N)
    Project3 = MUL3(N, DOT)

End Function





Public Function Rotate3(p As tVec3, Direc As tVec3) As tVec3

' TO TEST and ADJUST
'Stop
'
'    Rotate3.X = DOT3(V, Vec3(Direc.X, 0, -Direc.Z))
'    Rotate3.Y = DOT3(V, Vec3(0, 0, 0))
'    Rotate3.Z = DOT3(V, Vec3(Direc.Z, 0, Direc.X))
'
'    Rotate3.X = Rotate3.X + DOT3(V, Vec3(Direc.X, -Direc.Y, 0))
'    Rotate3.Y = Rotate3.Y + DOT3(V, Vec3(Direc.Y, Direc.X, 0))
'    Rotate3.Z = Rotate3.Z + DOT3(V, Vec3(0, 0, 0))
'
'    Rotate3.X = Rotate3.X + DOT3(V, Vec3(0, 0, 0))
'    Rotate3.Y = Rotate3.Y + DOT3(V, Vec3(0, Direc.Y, -Direc.Z))
'    Rotate3.Z = Rotate3.Z + DOT3(V, Vec3(0, Direc.Z, Direc.Y))

    Dim U         As tVec3
    Dim V         As tVec3
    Dim W         As tVec3

Dim R As tRotor3


'    'ALMOST RIGHT
'
'    W = Normalize3(Vec3(Direc.x, Direc.Y, -Direc.Z))
'    U = Normalize3(CROSS3(Vec3(0, -1, 0), W))
'    V = Normalize3(CROSS3(U, W))
'
'    Rotate3.x = DOT3(p, W)
'    Rotate3.Y = DOT3(p, V)
'    Rotate3.Z = DOT3(p, U)


    'http://marctenbosch.com/quaternions/


R = Rotor3FT(Vec3(0, 1, 0), Normalize3(Vec3(Direc.X, Direc.Y, Direc.Z)))

Rotate3 = Rotate3WithRotor(p, R)


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



Public Function Rotate3xz(V As tVec3, DirectionXZ As tVec3) As tVec3


' Corrected ONE:
'http://www.vbforums.com/showthread.php?874965-Rotation-using-DOT-product

    Rotate3xz.X = DOT3(V, Vec3(DirectionXZ.X, 0, -DirectionXZ.Z))
    'Rotate3xz.Y = V.Y * 1
    Rotate3xz.Y = DOT3(V, Vec3(0, 1, 0))
    Rotate3xz.Z = DOT3(V, Vec3(DirectionXZ.Z, 0, DirectionXZ.X))

End Function


Public Function Rotate3xy(V As tVec3, DirectionXY As tVec3) As tVec3

    Rotate3xy.X = DOT3(V, Vec3(DirectionXY.X, DirectionXY.Y, 0))
    Rotate3xy.Y = DOT3(V, Vec3(-DirectionXY.Y, DirectionXY.X, 0))
    Rotate3xy.Z = DOT3(V, Vec3(0, 0, 1))

End Function

Public Function Rotate3yz(V As tVec3, DirectionYZ As tVec3) As tVec3

    Rotate3yz.X = DOT3(V, Vec3(1, 0, 0))
    Rotate3yz.Y = DOT3(V, Vec3(0, DirectionYZ.Y, DirectionYZ.Y))
    Rotate3yz.Z = DOT3(V, Vec3(0, -DirectionYZ.Z, DirectionYZ.Y))

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




Public Function Rotor3(a As Double, b01 As Double, b02 As Double, b12 As Double) As tRotor3
    With Rotor3
        .a = a
        .b01 = b01
        .b02 = b02
        .b12 = b12
    End With
End Function


Public Function Rotor3Normalize(R As tRotor3) As tRotor3
    Dim L         As Double

    With R
        L = .a * .a + .b01 * .b01 + .b02 * .b02 + .b12 * .b12

        If L Then
            L = 1 / Sqr(L)
            Rotor3Normalize.a = .a * L
            Rotor3Normalize.b01 = .b01 * L
            Rotor3Normalize.b02 = .b02 * L
            Rotor3Normalize.b12 = .b12 * L
        End If

    End With

End Function

' construct the rotor that rotates one vector to another
 'uses the usual trick to get the half angle
Public Function Rotor3FT(vFrom As tVec3, vTo As tVec3) As tRotor3
    Dim minusb    As tVec3

    With Rotor3FT

        .a = 1 + DOT3(vTo, vFrom)
        ' the left side of the products have b a, not a b, so flip
        minusb = WEDGE3(vTo, vFrom)
        .b01 = minusb.X    '.b01
        .b02 = minusb.Y    '.b02
        .b12 = minusb.Z    '.b12
    End With

    Rotor3FT = Rotor3Normalize(Rotor3FT)

End Function


' angle+plane, plane must be normalized
Public Function Rotor3AP(angleRadian As Double, BiVectorPlane As tVec3) As tRotor3
    Dim SinA      As Double

    With Rotor3AP
        SinA = Sin(angleRadian * 0.5)
        .a = Cos(angleRadian * 0.5)
        ' the left side of the products have b a, not a b
        .b01 = -SinA * BiVectorPlane.X    '.b01
        .b02 = -SinA * BiVectorPlane.Y    '.b02
        .b12 = -SinA * BiVectorPlane.Z    '.b12
    End With

End Function

Public Function Rotor3Product(p As tRotor3, Q As tRotor3) As tRotor3
' Rotor3-Rotor3 product
' non-optimized
    With Rotor3Product
        .a = p.a * Q.a _
             - p.b01 * Q.b01 - p.b02 * Q.b02 - p.b12 * Q.b12

        .b01 = p.b01 * Q.a + p.a * Q.b01 _
               + p.b12 * Q.b02 - p.b02 * Q.b12

        .b02 = p.b02 * Q.a + p.a * Q.b02 _
               - p.b12 * Q.b01 + p.b01 * Q.b12

        .b12 = p.b12 * Q.a + p.a * Q.b12 _
               + p.b02 * Q.b01 - p.b01 * Q.b02

    End With

End Function



Public Function Rotate3WithRotor(V As tVec3, R As tRotor3) As tVec3

    Dim Q         As tVec3
Dim q012 As Double

    ' q = R V
    Q.X = R.a * V.X + V.Y * R.b01 + V.Z * R.b02
    Q.Y = R.a * V.Y - V.X * R.b01 + V.Z * R.b12
    Q.Z = R.a * V.Z - V.X * R.b02 - V.Y * R.b12

    q012 = -V.X * R.b12 + V.Y * R.b02 - V.Z * R.b01   ' trivector

    ' r = q R*
    With Rotate3WithRotor
        .X = R.a * Q.X + Q.Y * R.b01 + Q.Z * R.b02 - q012 * R.b12
        .Y = R.a * Q.Y - Q.X * R.b01 + q012 * R.b02 + Q.Z * R.b12
        .Z = R.a * Q.Z - q012 * R.b01 - Q.X * R.b02 - Q.Y * R.b12
    End With

End Function

'Public Function fastEXP(ByVal V As Double) As Double
''https://en.wikipedia.org/wiki/Pad%C3%A9_approximant
'    Dim X2        As Double
'    Dim X3        As Double
'    Dim X4        As Double
'    Dim X5        As Double
'
'
'    If V < 5! Then
'
'        If V < -7! Then fastEXP = 0!: Exit Function
'
'        X2 = V * V
'        X3 = X2 * V
'        X4 = X3 * V
'        X5 = X4 * V
'
'        fastEXP = (1! + INV2 * V + INV9 * X2 + INV72 * X3 + INV1008 * X4 + INV30240 * X5) / _
         '                  (1! - INV2 * V + INV9 * X2 - INV72 * X3 + INV1008 * X4 - INV30240 * X5)
'
'    Else
'        fastEXP = Exp(V)
'    End If
'
'
'End Function

