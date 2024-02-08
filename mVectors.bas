Attribute VB_Name = "mVectors"
Option Explicit

Public Const PI   As Single = 3.14159265358979
Attribute PI.VB_VarUserMemId = 1073938432
Public Const InvPI As Single = 1 / 3.14159265358979
Attribute InvPI.VB_VarUserMemId = 1073938433

Public Const PIh  As Single = 1.5707963267949
Public Const PI2  As Single = 6.28318530717959
Attribute PI2.VB_VarUserMemId = 1073938434


'Pade Approximant-----------****************************************


'SIN
Private Const C12671_D_4363920 As Single = 12671 / 4363920
Attribute C12671_D_4363920.VB_VarUserMemId = 1073938433
Private Const C2363_D_18183 As Single = 2363 / 18183
Attribute C2363_D_18183.VB_VarUserMemId = 1610809345
Private Const C121_D_16662240 As Single = 121 / 16662240
Attribute C121_D_16662240.VB_VarUserMemId = 1073938435
Private Const C601_D_872784 As Single = 601 / 872784
Attribute C601_D_872784.VB_VarUserMemId = 1073938436
Private Const C445_D_12122 As Single = 445 / 12122
Attribute C445_D_12122.VB_VarUserMemId = 1073938437

'EXP
Private Const INV2 As Single = 0.5
Attribute INV2.VB_VarUserMemId = 1073741830
Private Const INV9 As Single = 0.111111111111111
Attribute INV9.VB_VarUserMemId = 1610809346
Private Const INV72 As Single = 1.38888888888889E-02
Attribute INV72.VB_VarUserMemId = 1073938440
Private Const INV1008 As Single = 9.92063492063492E-04
Attribute INV1008.VB_VarUserMemId = 1073938441
Private Const INV30240 As Single = 3.30687830687831E-05
Attribute INV30240.VB_VarUserMemId = 1073938442


Public Type tVec3
    X             As Single
    Y             As Single
    Z             As Single
End Type


Public Type tRotor3
    ' scalar part
    A             As Single        '=1
    ' bivector part
    b01           As Single
    b02           As Single
    b12           As Single
End Type



Public Function Atan2(ByVal X As Single, ByVal Y As Single) As Single
Attribute Atan2.VB_UserMemId = 1073741848
    If X Then
        '        Stop
        '        Atan2 = -PI + Atn(Y / X) - (X > 0!) * PI
        '        Stop
        Atan2 = Atn(Y / X) + PI * (X < 0!)
    Else
        Atan2 = -PIh - (Y > 0!) * PI
    End If
End Function

Public Function ArcCos(X As Single) As Single
Attribute ArcCos.VB_UserMemId = 1073938457
    ArcCos = Atn(-X / Sqr(-X * X + 1)) + PIh
End Function
Public Function ArcSin(ByVal X As Single) As Single
Attribute ArcSin.VB_UserMemId = 1073741856
    ArcSin = Atn(X / Sqr(-X * X + 1))
End Function

Public Function Vec3(X As Single, Y As Single, Z As Single) As tVec3
Attribute Vec3.VB_UserMemId = 1073741860
    Vec3.X = X
    Vec3.Y = Y
    Vec3.Z = Z
End Function

Public Function Length3(V As tVec3) As Single
Attribute Length3.VB_UserMemId = 1610612740
    With V
        Length3 = Sqr(.X * .X + .Y * .Y + .Z * .Z)
    End With
End Function

Public Function Length32(V As tVec3) As Single
Attribute Length32.VB_UserMemId = 1610612741
    With V
        Length32 = .X * .X + .Y * .Y + .Z * .Z
    End With
End Function



Public Function SUM3(V1 As tVec3, V2 As tVec3) As tVec3
Attribute SUM3.VB_UserMemId = 1073741868
    SUM3.X = V1.X + V2.X
    SUM3.Y = V1.Y + V2.Y
    SUM3.Z = V1.Z + V2.Z
End Function



Public Function Normalize3(V As tVec3) As tVec3
Attribute Normalize3.VB_UserMemId = 1610612743
    Dim D         As Single
    With V
        D = (.X * .X + .Y * .Y + .Z * .Z)
        If D Then
            D = 1 / Sqr(D)
            Normalize3.X = .X * D
            Normalize3.Y = .Y * D
            Normalize3.Z = .Z * D
        End If
    End With
End Function

Public Function MUL3(V As tVec3, ByVal A As Single) As tVec3
Attribute MUL3.VB_UserMemId = 1610612744
    MUL3.X = V.X * A
    MUL3.Y = V.Y * A
    MUL3.Z = V.Z * A

End Function

Public Function DOT3(V1 As tVec3, V2 As tVec3) As Single
Attribute DOT3.VB_UserMemId = 1610612745

    DOT3 = (V1.X * V2.X) + _
           (V1.Y * V2.Y) + _
           (V1.Z * V2.Z)

End Function

Public Function CROSS3(A As tVec3, B As tVec3) As tVec3
Attribute CROSS3.VB_UserMemId = 1610612746
    CROSS3.X = A.Y * B.Z - A.Z * B.Y
    CROSS3.Y = A.Z * B.X - A.X * B.Z
    CROSS3.Z = A.X * B.Y - A.Y * B.X

End Function



'// Wedge product
Public Function WEDGE3(A As tVec3, B As tVec3) As tVec3    'BiVector
Attribute WEDGE3.VB_UserMemId = 1073741895

    WEDGE3.X = A.X * B.Y - A.Y * B.X    ', // XY
    WEDGE3.Y = A.X * B.Z - A.Z * B.X    ', // XZ
    WEDGE3.Z = A.Y * B.Z - A.Z * B.Y    '  // YZ


End Function


Public Function DIFF3(V1 As tVec3, V2 As tVec3) As tVec3
    DIFF3.X = V1.X - V2.X
    DIFF3.Y = V1.Y - V2.Y
    DIFF3.Z = V1.Z - V2.Z
End Function


Public Function Project3(V As tVec3, V2nrmlzd As tVec3) As tVec3
Attribute Project3.VB_UserMemId = 1610612749

    Project3 = MUL3(V2nrmlzd, DOT3(V, V2nrmlzd))

End Function


Public Function ProjectToPlane3(V As tVec3, PlaneN As tVec3) As tVec3
Attribute ProjectToPlane3.VB_UserMemId = 1610612750
    ' unsure !
    ProjectToPlane3 = DIFF3(V, Project3(V, PlaneN))
    'https://www.physicsforums.com/threads/projecting-a-vector-onto-a-plane.496184/
    'Project3 = CROSS3(V2nrmlzd, CROSS3(v, V2nrmlzd))
End Function




Public Function Rotate3(P As tVec3, Direc As tVec3) As tVec3
Attribute Rotate3.VB_UserMemId = 1610612751

    ' TO TEST and ADJUST
    'Stop

    Dim U         As tVec3
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
Attribute ToString3.VB_UserMemId = 1610612752
    ToString3 = V.X & "   " & V.Y & "   " & V.Z
End Function




''''******************************************************************
''''   TODO:
''''   http://paulbourke.net/geometry/rotate/
''''   http://paulbourke.net/geometry/rotate/source.c
''''******************************************************************
'''Public Function Rotate3Axe(P As tVec3, Axe As tVec3, theta As single) As tVec3
'''    Dim L         As single
'''    Dim CosTheta  As single
'''    Dim SinTheta  As single
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
Attribute Rotate3xz.VB_UserMemId = 1610612753


    ' Corrected ONE:
    'http://www.vbforums.com/showthread.php?874965-Rotation-using-DOT-product

    Rotate3xz.X = DOT3(V, Vec3(DirectionXZ.X, 0, -DirectionXZ.Z))
    'Rotate3xz.Y = V.Y * 1
    Rotate3xz.Y = DOT3(V, Vec3(0, 1, 0))
    Rotate3xz.Z = DOT3(V, Vec3(DirectionXZ.Z, 0, DirectionXZ.X))

End Function


Public Function Rotate3xy(V As tVec3, DirectionXY As tVec3) As tVec3
Attribute Rotate3xy.VB_UserMemId = 1610809351

    Rotate3xy.X = DOT3(V, Vec3(DirectionXY.X, DirectionXY.Y, 0))
    Rotate3xy.Y = DOT3(V, Vec3(-DirectionXY.Y, DirectionXY.X, 0))
    Rotate3xy.Z = DOT3(V, Vec3(0, 0, 1))

End Function

Public Function Rotate3yz(V As tVec3, DirectionYZ As tVec3) As tVec3
Attribute Rotate3yz.VB_UserMemId = 1610612755

    Rotate3yz.X = DOT3(V, Vec3(1, 0, 0))
    Rotate3yz.Y = DOT3(V, Vec3(0, DirectionYZ.Y, DirectionYZ.Z))
    Rotate3yz.Z = DOT3(V, Vec3(0, -DirectionYZ.Z, DirectionYZ.Y))

End Function

Public Function Rotate3zx(V As tVec3, DirectionZX As tVec3) As tVec3
Attribute Rotate3zx.VB_UserMemId = 1610612756

    Rotate3zx.Z = DOT3(V, Vec3(0, DirectionZX.X, DirectionZX.Z))
    Rotate3zx.Y = DOT3(V, Vec3(0, 1, 0))
    Rotate3zx.X = DOT3(V, Vec3(0, -DirectionZX.Z, DirectionZX.X))


End Function


Public Function RayPlaneIntersect(rayVector As tVec3, rayPoint As tVec3, PlaneNormal As tVec3, planePoint As tVec3) As tVec3
Attribute RayPlaneIntersect.VB_UserMemId = 1610612757
    'https://rosettacode.org/wiki/Find_the_intersection_of_a_line_with_a_plane#C.23

    Dim Diff      As tVec3
    Dim prod1     As Single
    Dim prod2     As Single
    Dim prod3     As Single

    Diff = DIFF3(rayPoint, planePoint)
    prod1 = DOT3(Diff, PlaneNormal)
    prod2 = DOT3(rayVector, PlaneNormal)
    prod3 = prod1 / prod2
    RayPlaneIntersect = DIFF3(rayPoint, MUL3(rayVector, prod3))

End Function


Public Function XZPerp(V As tVec3) As tVec3
Attribute XZPerp.VB_UserMemId = 1610612758
    'LHR
    XZPerp.X = V.Z
    XZPerp.Y = V.Y
    XZPerp.Z = -V.X

    ''Right Hand
    '    XZPerp.X = -v.Z
    '    XZPerp.Y = v.Y
    '    XZPerp.Z = v.X

End Function




Public Function Rotor3(A As Single, b01 As Single, b02 As Single, b12 As Single) As tRotor3
Attribute Rotor3.VB_UserMemId = 1610612759
    With Rotor3
        .A = A
        .b01 = b01
        .b02 = b02
        .b12 = b12
    End With
End Function


Public Function Rotor3Normalize(R As tRotor3) As tRotor3
Attribute Rotor3Normalize.VB_UserMemId = 1610612760
    Dim L         As Single

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
Attribute Rotor3FT.VB_UserMemId = 1610612761
    Dim minusb    As tVec3

    With Rotor3FT

        .A = 1 + DOT3(vTo, vFrom)
        ' the left side of the products have b a, not a b, so flip
        minusb = WEDGE3(vTo, vFrom)
        .b01 = minusb.X            '.b01
        .b02 = minusb.Y            '.b02
        .b12 = minusb.Z            '.b12
    End With

    Rotor3FT = Rotor3Normalize(Rotor3FT)

End Function


' angle+plane, plane must be normalized
Public Function Rotor3AP(angleRadian As Single, BiVectorPlane As tVec3) As tRotor3
Attribute Rotor3AP.VB_UserMemId = 1610612762
    Dim SinA      As Single

    With Rotor3AP
        SinA = Sin(angleRadian * 0.5)
        .A = Cos(angleRadian * 0.5)
        ' the left side of the products have b a, not a b
        .b01 = -SinA * BiVectorPlane.X    '.b01
        .b02 = -SinA * BiVectorPlane.Y    '.b02
        .b12 = -SinA * BiVectorPlane.Z    '.b12
    End With

End Function

Public Function Rotor3Product(P As tRotor3, Q As tRotor3) As tRotor3
Attribute Rotor3Product.VB_UserMemId = 1610612763
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
Attribute Rotate3WithRotor.VB_UserMemId = 1610612764

    Dim Q         As tVec3
    Dim q012      As Single

    ' q = R V
    Q.X = R.A * V.X + V.Y * R.b01 + V.Z * R.b02
    Q.Y = R.A * V.Y - V.X * R.b01 + V.Z * R.b12
    Q.Z = R.A * V.Z - V.X * R.b02 - V.Y * R.b12

    q012 = -V.X * R.b12 + V.Y * R.b02 - V.Z * R.b01    ' trivector

    ' r = q R*
    With Rotate3WithRotor
        .X = R.A * Q.X + Q.Y * R.b01 + Q.Z * R.b02 - q012 * R.b12
        .Y = R.A * Q.Y - Q.X * R.b01 + q012 * R.b02 + Q.Z * R.b12
        .Z = R.A * Q.Z - q012 * R.b01 - Q.X * R.b02 - Q.Y * R.b12
    End With

End Function














Public Function fastEXP(ByVal V As Single) As Single
Attribute fastEXP.VB_UserMemId = 1610612765
    'https://en.wikipedia.org/wiki/Pad%C3%A9_approximant
    Dim X2        As Single
    Dim X3        As Single
    Dim X4        As Single
    Dim X5        As Single


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



Public Function FastSIN(ByVal X As Single) As Single
Attribute FastSIN.VB_UserMemId = 1610612766
    'https://math.stackexchange.com/questions/2196371/how-to-approximate-sinx-using-pad%C3%A9-approximation
    ' ORDER 13 K 4

    Dim X2        As Single
    Dim X3        As Single
    Dim X4        As Single
    Dim X5        As Single
    Dim X6        As Single
    '
    While X > PI: X = X - PI2: Wend
    While X < -PI: X = X + PI2: Wend

    X2 = X * X
    X3 = X2 * X
    X4 = X3 * X
    X5 = X4 * X
    X6 = X5 * X

    FastSIN = (C12671_D_4363920 * X5 - C2363_D_18183 * X3 + X) / _
              (C121_D_16662240 * X6 + C601_D_872784 * X4 + C445_D_12122 * X2 + 1!)

End Function

Public Function FastCOS(ByVal X As Single) As Single
Attribute FastCOS.VB_UserMemId = 1610612767
    FastCOS = FastSIN(X + PIh)
End Function


Public Function AngleDIFF(ByRef A1 As Single, ByRef A2 As Single) As Single
Attribute AngleDIFF.VB_UserMemId = 1610612768

    AngleDIFF = A1 - A2
    While AngleDIFF < -PI
        AngleDIFF = AngleDIFF + PI2
    Wend
    While AngleDIFF > PI
        AngleDIFF = AngleDIFF - PI2
    Wend
End Function


''https://github.com/processing/processing/blob/349f413a3fb63a75e0b096097a5b0ba7f5565198/core/src/processing/core/PVector.java
'Public Function AngleBetween(V1 As tVec3, v2 As tVec3) As single
'    Dim Dot       As single
'    Dim v1Mag     As single
'    Dim v2Mag     As single
'    Dim amt       As single
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

Public Function AngleBetween(V1 As tVec3, V2 As tVec3) As Single
Attribute AngleBetween.VB_UserMemId = 1610612769
    'http://www.dotnetframework.org/default.aspx/Net/Net/3@5@50727@3053/DEVDIV/depot/DevDiv/releases/Orcas/SP/wpf/src/Core/CSharp/System/Windows/Media3D/Vector3D@cs/1/Vector3D@cs


    Dim Ratio     As Single
    Dim nV1       As tVec3
    Dim nV2       As tVec3

    nV1 = Normalize3(V1)
    nV2 = Normalize3(V2)

    Ratio = DOT3(nV1, nV2)

    If (Ratio < 0) Then
        '   Math.PI - 2.0 * Math.Asin((-vector1 - vector2).Length / 2.0);
        AngleBetween = PI - 2 * ArcSin(Length3(SUM3(nV1, nV2)) * 0.5)
    Else
        '   2.0 * Math.Asin((vector1 - vector2).Length / 2.0);
        AngleBetween = 2 * ArcSin(Length3(DIFF3(nV1, nV2)) * 0.5)
    End If

End Function

