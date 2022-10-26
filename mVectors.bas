Attribute VB_Name = "mVectors"
Option Explicit

Public Const PI   As Double = 3.14159265358979
Public Const InvPI As Double = 1 / 3.14159265358979

Public Const PIh  As Double = 1.5707963267949
Public Const PI2  As Double = 6.28318530717959


'Pade Approximant-----------****************************************


'SIN
Private Const C12671_D_4363920 As Single = 12671 / 4363920
Private Const C2363_D_18183 As Single = 2363 / 18183
Private Const C121_D_16662240 As Single = 121 / 16662240
Private Const C601_D_872784 As Single = 601 / 872784
Private Const C445_D_12122 As Single = 445 / 12122

'EXP
Private Const INV2 As Single = 0.5
Private Const INV9 As Single = 0.111111111111111
Private Const INV72 As Single = 1.38888888888889E-02
Private Const INV1008 As Single = 9.92063492063492E-04
Private Const INV30240 As Single = 3.30687830687831E-05


Public Type tVec3
    x             As Double
    y             As Double
    z             As Double
End Type


Public Type tRotor3
    ' scalar part
    A             As Double  '=1
    ' bivector part
    b01           As Double
    b02           As Double
    b12           As Double
End Type



Public Function Atan2(ByVal x As Double, ByVal y As Double) As Double
    If x Then
        '        Stop
        '        Atan2 = -PI + Atn(Y / X) - (X > 0!) * PI
        '        Stop
        Atan2 = Atn(y / x) + PI * (x < 0!)
    Else
        Atan2 = -PIh - (y > 0!) * PI
    End If
End Function

Public Function ArcCos(x As Double) As Double
    ArcCos = Atn(-x / Sqr(-x * x + 1)) + PIh
End Function
Public Function ArcSin(ByVal x As Single) As Single
    ArcSin = Atn(x / Sqr(-x * x + 1))
End Function

Public Function Vec3(x As Double, y As Double, z As Double) As tVec3
    Vec3.x = x
    Vec3.y = y
    Vec3.z = z
End Function

Public Function Length3(V As tVec3) As Double
    With V
        Length3 = Sqr(.x * .x + .y * .y + .z * .z)
    End With
End Function

Public Function Length32(V As tVec3) As Double
    With V
        Length32 = .x * .x + .y * .y + .z * .z
    End With
End Function



Public Function SUM3(V1 As tVec3, V2 As tVec3) As tVec3
    SUM3.x = V1.x + V2.x
    SUM3.y = V1.y + V2.y
    SUM3.z = V1.z + V2.z
End Function



Public Function Normalize3(V As tVec3) As tVec3
    Dim D         As Double

    D = (V.x * V.x + V.y * V.y + V.z * V.z)
    If D Then
        D = 1# / Sqr(D)
        Normalize3.x = V.x * D
        Normalize3.y = V.y * D
        Normalize3.z = V.z * D
    End If

End Function

Public Function MUL3(V As tVec3, ByVal A As Double) As tVec3
    MUL3.x = V.x * A
    MUL3.y = V.y * A
    MUL3.z = V.z * A

End Function

Public Function DOT3(V1 As tVec3, V2 As tVec3) As Double

    DOT3 = (V1.x * V2.x) + _
           (V1.y * V2.y) + _
           (V1.z * V2.z)

End Function

Public Function CROSS3(A As tVec3, B As tVec3) As tVec3
    CROSS3.x = A.y * B.z - A.z * B.y
    CROSS3.y = A.z * B.x - A.x * B.z
    CROSS3.z = A.x * B.y - A.y * B.x

End Function



'// Wedge product
Public Function WEDGE3(A As tVec3, B As tVec3) As tVec3    'BiVector

    WEDGE3.x = A.x * B.y - A.y * B.x    ', // XY
    WEDGE3.y = A.x * B.z - A.z * B.x    ', // XZ
    WEDGE3.z = A.y * B.z - A.z * B.y    '  // YZ


End Function


Public Function DIFF3(V1 As tVec3, V2 As tVec3) As tVec3
    DIFF3.x = V1.x - V2.x
    DIFF3.y = V1.y - V2.y
    DIFF3.z = V1.z - V2.z
End Function


Public Function Project3(V As tVec3, V2nrmlzd As tVec3) As tVec3

    Project3 = MUL3(V2nrmlzd, DOT3(V, V2nrmlzd))

End Function


Public Function ProjectToPlane3(V As tVec3, PlaneN As tVec3) As tVec3
    Dim DOT       As Double
    ' unsure !
    ProjectToPlane3 = DIFF3(V, Project3(V, PlaneN))
    'https://www.physicsforums.com/threads/projecting-a-vector-onto-a-plane.496184/
    'Project3 = CROSS3(V2nrmlzd, CROSS3(v, V2nrmlzd))
End Function




Public Function Rotate3(P As tVec3, Direc As tVec3) As tVec3

' TO TEST and ADJUST
'Stop

    Dim U         As tVec3
    Dim V         As tVec3
    Dim W         As tVec3

    Dim R         As tRotor3

    Dim A         As tVec3
    Dim B         As tVec3
    Dim C         As tVec3

    Dim PL1       As tVec3
    Dim PL2       As tVec3
    Dim PL3       As tVec3
    Dim D         As tVec3




    D = Normalize3(Direc)
    '    D = MUL3(D, -1)


    Rotate3 = Rotate3yz(P, D)
    Rotate3 = Rotate3xy(Rotate3, D)




    '    D = Normalize3(Direc)
    '
    '    A = MUL3(D, 1)
    '    B = CROSS3(Vec3(0, -1, 0), A)
    '    C = CROSS3(B, A)
    '
    '
    '    U = Project3(p, A)
    '    v = Project3(p, B)
    '    W = Project3(p, C)
    '
    '    Rotate3 = U 'SUM3(U, SUM3(v, W))
    '
    ''    Rotate3.Z = W.Z
    ''    Rotate3.X = U.X
    ''    Rotate3.Y = v.Y
    '------------------------------------------------------------------
    'PL1 = Normalize3(Direc)
    'PL2 = CROSS3(Vec3(0, 1, 0), PL1)
    'PL3 = CROSS3(PL2, PL1)
    '
    'A = ProjectToPlane3(p, PL1)
    'B = ProjectToPlane3(p, PL2)
    'C = ProjectToPlane3(p, PL3)
    '
    'A = Normalize3(A)
    'B = Normalize3(B)
    'C = Normalize3(C)
    '
    '
    'Rotate3.X = DOT3(p, A)
    'Rotate3.Y = DOT3(p, B)
    'Rotate3.Z = DOT3(p, C)


    '------------------------------------------------------------------

    '    'ALMOST RIGHT
    '
    ''        W = Normalize3(Vec3(Direc.X, -Direc.Y, Direc.Z))
    ''        U = Normalize3(CROSS3(W, Vec3(0, -1, 0)))
    ''        v = Normalize3(CROSS3(W, U))
    ''
    ''        Rotate3.X = DOT3(p, U)
    ''        Rotate3.Y = DOT3(p, v)
    ''        Rotate3.Z = DOT3(p, W)


    '    'http://marctenbosch.com/quaternions/
    '    R = Rotor3FT(Vec3(0, 1, 0), Normalize3(Vec3(Direc.X, Direc.Y, Direc.Z)))
    '    Rotate3 = Rotate3WithRotor(p, R)


End Function



Public Function ToString3(V As tVec3) As String
    ToString3 = V.x & "   " & V.y & "   " & V.z
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

    Rotate3xz.x = DOT3(V, Vec3(DirectionXZ.x, 0, -DirectionXZ.z))
    'Rotate3xz.Y = V.Y * 1
    Rotate3xz.y = DOT3(V, Vec3(0, 1, 0))
    Rotate3xz.z = DOT3(V, Vec3(DirectionXZ.z, 0, DirectionXZ.x))

End Function


Public Function Rotate3xy(V As tVec3, DirectionXY As tVec3) As tVec3

    Rotate3xy.x = DOT3(V, Vec3(DirectionXY.x, DirectionXY.y, 0))
    Rotate3xy.y = DOT3(V, Vec3(-DirectionXY.y, DirectionXY.x, 0))
    Rotate3xy.z = DOT3(V, Vec3(0, 0, 1))

End Function

Public Function Rotate3yz(V As tVec3, DirectionYZ As tVec3) As tVec3

    Rotate3yz.x = DOT3(V, Vec3(1, 0, 0))
    Rotate3yz.y = DOT3(V, Vec3(0, DirectionYZ.y, DirectionYZ.z))
    Rotate3yz.z = DOT3(V, Vec3(0, -DirectionYZ.z, DirectionYZ.y))

End Function

Public Function Rotate3zx(V As tVec3, DirectionZX As tVec3) As tVec3

    Rotate3zx.z = DOT3(V, Vec3(0, DirectionZX.x, DirectionZX.z))
    Rotate3zx.y = DOT3(V, Vec3(0, 1, 0))
    Rotate3zx.x = DOT3(V, Vec3(0, -DirectionZX.z, DirectionZX.x))


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


Public Function XZPerp(V As tVec3) As tVec3
'LHR
    XZPerp.x = V.z
    XZPerp.y = V.y
    XZPerp.z = -V.x

    ''Right Hand
    '    XZPerp.X = -v.Z
    '    XZPerp.Y = v.Y
    '    XZPerp.Z = v.X

End Function




Public Function Rotor3(A As Double, b01 As Double, b02 As Double, b12 As Double) As tRotor3
    With Rotor3
        .A = A
        .b01 = b01
        .b02 = b02
        .b12 = b12
    End With
End Function


Public Function Rotor3Normalize(R As tRotor3) As tRotor3
    Dim L         As Double

    With R
        L = .A * .A + .b01 * .b01 + .b02 * .b02 + .b12 * .b12

        If L Then
            L = 1 / Sqr(L)
            Rotor3Normalize.A = .A * L
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

        .A = 1 + DOT3(vTo, vFrom)
        ' the left side of the products have b a, not a b, so flip
        minusb = WEDGE3(vTo, vFrom)
        .b01 = minusb.x      '.b01
        .b02 = minusb.y      '.b02
        .b12 = minusb.z      '.b12
    End With

    Rotor3FT = Rotor3Normalize(Rotor3FT)

End Function


' angle+plane, plane must be normalized
Public Function Rotor3AP(angleRadian As Double, BiVectorPlane As tVec3) As tRotor3
    Dim SinA      As Double

    With Rotor3AP
        SinA = Sin(angleRadian * 0.5)
        .A = Cos(angleRadian * 0.5)
        ' the left side of the products have b a, not a b
        .b01 = -SinA * BiVectorPlane.x    '.b01
        .b02 = -SinA * BiVectorPlane.y    '.b02
        .b12 = -SinA * BiVectorPlane.z    '.b12
    End With

End Function

Public Function Rotor3Product(P As tRotor3, Q As tRotor3) As tRotor3
' Rotor3-Rotor3 product
' non-optimized
    With Rotor3Product
        .A = P.A * Q.A _
             - P.b01 * Q.b01 - P.b02 * Q.b02 - P.b12 * Q.b12

        .b01 = P.b01 * Q.A + P.A * Q.b01 _
               + P.b12 * Q.b02 - P.b02 * Q.b12

        .b02 = P.b02 * Q.A + P.A * Q.b02 _
               - P.b12 * Q.b01 + P.b01 * Q.b12

        .b12 = P.b12 * Q.A + P.A * Q.b12 _
               + P.b02 * Q.b01 - P.b01 * Q.b02

    End With

End Function



Public Function Rotate3WithRotor(V As tVec3, R As tRotor3) As tVec3

    Dim Q         As tVec3
    Dim q012      As Double

    ' q = R V
    Q.x = R.A * V.x + V.y * R.b01 + V.z * R.b02
    Q.y = R.A * V.y - V.x * R.b01 + V.z * R.b12
    Q.z = R.A * V.z - V.x * R.b02 - V.y * R.b12

    q012 = -V.x * R.b12 + V.y * R.b02 - V.z * R.b01    ' trivector

    ' r = q R*
    With Rotate3WithRotor
        .x = R.A * Q.x + Q.y * R.b01 + Q.z * R.b02 - q012 * R.b12
        .y = R.A * Q.y - Q.x * R.b01 + q012 * R.b02 + Q.z * R.b12
        .z = R.A * Q.z - q012 * R.b01 - Q.x * R.b02 - Q.y * R.b12
    End With

End Function














Public Function fastEXP(ByVal V As Double) As Double
'https://en.wikipedia.org/wiki/Pad%C3%A9_approximant
    Dim X2        As Double
    Dim X3        As Double
    Dim X4        As Double
    Dim X5        As Double


    If V < 5! Then

        If V < -7! Then fastEXP = 0!: Exit Function

        X2 = V * V
        X3 = X2 * V
        X4 = X3 * V
        X5 = X4 * V

        fastEXP = (1! + INV2 * V + INV9 * X2 + INV72 * X3 + INV1008 * X4 + INV30240 * X5) / _
                  (1! - INV2 * V + INV9 * X2 - INV72 * X3 + INV1008 * X4 - INV30240 * X5)

    Else
        fastEXP = Exp(V)
    End If


End Function



Public Function FastSIN(ByVal x As Single) As Single
'https://math.stackexchange.com/questions/2196371/how-to-approximate-sinx-using-pad%C3%A9-approximation
' ORDER 13 K 4

    Dim X2        As Single
    Dim X3        As Single
    Dim X4        As Single
    Dim X5        As Single
    Dim X6        As Single
    '
    While x > PI: x = x - PI2: Wend
    While x < -PI: x = x + PI2: Wend

    X2 = x * x
    X3 = X2 * x
    X4 = X3 * x
    X5 = X4 * x
    X6 = X5 * x

    FastSIN = (C12671_D_4363920 * X5 - C2363_D_18183 * X3 + x) / _
              (C121_D_16662240 * X6 + C601_D_872784 * X4 + C445_D_12122 * X2 + 1!)

End Function

Public Function FastCOS(ByVal x As Single) As Single
    FastCOS = FastSIN(x + PIh)
End Function


Public Function AngleDIFF(ByRef A1 As Double, ByRef A2 As Double) As Double

    AngleDIFF = A1 - A2
    While AngleDIFF < -PI
        AngleDIFF = AngleDIFF + PI2
    Wend
    While AngleDIFF > PI
        AngleDIFF = AngleDIFF - PI2
    Wend
End Function


''https://github.com/processing/processing/blob/349f413a3fb63a75e0b096097a5b0ba7f5565198/core/src/processing/core/PVector.java
'Public Function AngleBetween(V1 As tVec3, v2 As tVec3) As Double
'    Dim Dot       As Double
'    Dim v1Mag     As Double
'    Dim v2Mag     As Double
'    Dim amt       As Double
'
'    Dot = V1.X * v2.X + V1.Y * v2.Y + V1.Z * v2.Z
'    v1Mag = Sqr(V1.X * V1.X + V1.Y * V1.Y + V1.Z * V1.Z)
'    v2Mag = Sqr(v2.X * v2.X + v2.Y * v2.Y + v2.Z * v2.Z)
'    '  This should be a number between -1 and 1, since it's "normalized"
'    amt = Dot / (v1Mag * v2Mag)
'    '  But if it's not due to rounding error, then we need to fix it
'    '  http://code.google.com/p/processing/issues/detail?id=340
'    '  Otherwise if outside the range, acos() will return NaN
'    '  http://www.cppreference.com/wiki/c/math/acos
'    If (amt <= -1) Then
'        AngleBetween = PI: Exit Function
'    ElseIf (amt >= 1) Then
'        '  http://code.google.com/p/processing/issues/detail?id=435
'        AngleBetween = 0: Exit Function
'    End If
'
'    AngleBetween = ArcCos(amt)
'
'
'End Function

Public Function AngleBetween(V1 As tVec3, V2 As tVec3) As Double
'http://www.dotnetframework.org/default.aspx/Net/Net/3@5@50727@3053/DEVDIV/depot/DevDiv/releases/Orcas/SP/wpf/src/Core/CSharp/System/Windows/Media3D/Vector3D@cs/1/Vector3D@cs


    Dim Ratio     As Double
    Dim nV1       As tVec3
    Dim nV2       As tVec3

    nV1 = Normalize3(V1)
    nV2 = Normalize3(V2)

    Ratio = DOT3(nV1, nV2)

    If (Ratio < 0) Then
        '   Math.PI - 2.0 * Math.Asin((-vector1 - vector2).Length / 2.0);
        AngleBetween = PI - 2# * ArcSin(Length3(SUM3(nV1, nV2)) * 0.5)
    Else
        '   2.0 * Math.Asin((vector1 - vector2).Length / 2.0);
        AngleBetween = 2# * ArcSin(Length3(DIFF3(nV1, nV2)) * 0.5)
    End If

End Function

