Attribute VB_Name = "mMain"
Option Explicit


Public Const ExpectedMaxDensity As Double = 64
Public Const INVExpectedMaxDensity As Double = 1 / ExpectedMaxDensity


Public SpatialGRID As cSpatialGrid3D
'Public HASH3D As cSpatialHash3D


Public PIChDC     As Long

Public WW         As Long
Public HH         As Long
Public ZZ         As Long
Public invZZ      As Double


Public pX()       As Double
Public pY()       As Double
Public pZ()       As Double

Public vX()       As Double
Public vY()       As Double
Public vZ()       As Double

Public pvX()      As Double
Public pvY()      As Double
Public pvZ()      As Double


Public NP         As Long

Public DrawPairs  As Boolean
Public DoLOOP     As Boolean
Public Frame      As Long
Public DoSaveFrames As Boolean
Public rndGravity As Boolean

Public DoFaucet1  As Boolean
Public DoFaucet2  As Boolean

Public COMGravity As Boolean
Public GravScale  As Double


Public FPS        As Double
Public CNT        As Long
Public OldCNT     As Long

Public mTime      As Double
Public OldmTime   As Double


Public RenderEvery As Long

Public h          As Double  'smoothing radius
Public invH       As Double
Public InvH2      As Double
Public SQR_Table() As Double
Public Normalize_Table() As Double
Public SmoothKernel_Table() As Double
Public InvDensity_Table() As Double
Public Visco_Table() As Double


Public Const TABLESLength As Double = 2 ^ 14 - 1    ' 2 ^ 14 - 1

Public Const kRestitution As Double = 0.65    ' 0.65    '0.75    ' 0.75 '0.5    '0.66
Public Const kWallFriction As Double = 0.99    '0.98 '0.996    '0.995
Public Const kFakeDensity As Double = 2    '0.3  '2022
Public Const kFakeVel As Double = 0.005



Public gX         As Double  'GRAVITY
Public gY         As Double
Public gZ         As Double

Public gTOX       As Double
Public gTOY       As Double
Public gTOZ       As Double


' SPH ---------------------------------------------------------

Public DT         As Double
Public invDT      As Double


Private RestDensity As Double
Private INVRestDensity As Double
Private PressureLimit As Double
Private KAttraction As Double
Private KPressure As Double
Private KViscosity As Double


Public Density()  As Double
Public Pressure() As Double
Public Phase()    As Long
Public INVDensity() As Double



Public VXChange() As Double
Public VYChange() As Double
Public VZChange() As Double


Public P1()       As Long
Public P2()       As Long
Public arrDX()    As Double
Public arrDY()    As Double
Public arrDZ()    As Double

Public arrD()     As Double
Public RetNofPairs As Long
Public MaxNofPairs As Long


Public CAMERA     As c3DEasyCam

Public COMx       As Double
Public COMy       As Double
Public COMz       As Double

Public CamRot     As Boolean


Public TestMaxDens As Double




Public Sub SPH_InitConst()
    Dim R         As Double
    Dim kernelWeight As Double
    Dim I         As Double

    R = 0.33333333333
    '    RestDensity = SmoothKernel_3(r) * 6    ' 2D
    '    RestDensity = SmoothKernel_3(R) * 6#    ' 3D
    RestDensity = SmoothKernel_3(R) * 5#     '2023
    
'    RestDensity = SmoothKernel_3(R) * 5 * 3 'Without Attraction

    For R = 0 To 1 Step 0.001
        I = I + 1
        kernelWeight = kernelWeight + SmoothKernel_3(R)
    Next

    'kernelWeight = kernelWeight / (I + 1)
    kernelWeight = kernelWeight / (I)


    INVRestDensity = 1 / RestDensity
    '    PressureLimit = 400      '200 '100    '50    '45 '20
    PressureLimit = 500     '800 '2022

    DT = 0.25
    invDT = 1 / DT

    'KAttraction = 0.0128 * invDT
    KAttraction = 0.0128 * invDT * 0.72    '0.75    ' 0.85    '2023
'KAttraction = 0 'Without Attraction
'    'KPressure = kernelWeight * 0.08 * invDT
    KPressure = kernelWeight * 0.15 ' 0.12 '0.15    '0.12    '0.15 * invDT '2022

KPressure = KPressure * 1.5
'KPressure = KPressure * 80 'Without Attraction

    'KViscosity = 0.018 * 0.8
    KViscosity = 0.018 * 0.5     '1 '1.5 ' 0.66    '0.5    '2022

'KViscosity = KViscosity * 1 'Without Attraction

    ReDim VXChange(NP)
    ReDim VYChange(NP)
    ReDim VZChange(NP)

    ReDim Density(NP)
    ReDim Pressure(NP)
    ReDim INVDensity(NP)

    ReDim Phase(NP)

    ReDim SQR_Table(TABLESLength)
    ReDim Normalize_Table(TABLESLength)
    ReDim SmoothKernel_Table(TABLESLength)
    ReDim InvDensity_Table(TABLESLength)
    ReDim Visco_Table(TABLESLength)


    For I = 0 To TABLESLength
        SQR_Table(I) = Sqr(I / TABLESLength)


        If I Then Normalize_Table(I) = 1 / (h * I / TABLESLength)
        SmoothKernel_Table(I) = SmoothKernel_3(I / TABLESLength)
        If I Then InvDensity_Table(I) = 1 / (ExpectedMaxDensity * (I / TABLESLength))    '

        R = I / TABLESLength
'        If R Then Visco_Table(I) = -0.5 * R * R * R + R * R + 0.5 / R - 1#
        If R Then Visco_Table(I) = (1 - R) ^ 5 '2023

    Next

End Sub

Public Sub SPH_MOVE()
    Dim I         As Long





    Dim invNP     As Double
    Dim DX        As Double
    Dim DY        As Double
    Dim DZ        As Double
    Dim D         As Double
    Dim F         As Double

    Dim wwH#, HHH#, zzH#
    Dim S#

    wwH = WW - h
    HHH = HH - h
    zzH = ZZ - h


    For I = 1 To NP
        vX(I) = vX(I) + VXChange(I) + gX * DT
        vY(I) = vY(I) + VYChange(I) + gY * DT
        vZ(I) = vZ(I) + VZChange(I) + gZ * DT

        VXChange(I) = 0#
        VYChange(I) = 0#
        VZChange(I) = 0#

        vX(I) = vX(I) * 0.9995    ' 0.998#
        vY(I) = vY(I) * 0.9995    ' 0.998#
        vZ(I) = vZ(I) * 0.9995    ' 0.998#

        pX(I) = pX(I) + vX(I) * DT
        pY(I) = pY(I) + vY(I) * DT
        pZ(I) = pZ(I) + vZ(I) * DT


        Density(I) = 0#


        If pX(I) < 0# Then pX(I) = -pX(I): vX(I) = -vX(I) * kRestitution: vY(I) = vY(I) * kWallFriction: vZ(I) = vZ(I) * kWallFriction
        If pY(I) < 0# Then pY(I) = -pY(I): vY(I) = -vY(I) * kRestitution: vX(I) = vX(I) * kWallFriction: vZ(I) = vZ(I) * kWallFriction
        If pZ(I) < 0# Then pZ(I) = -pZ(I): vZ(I) = -vZ(I) * kRestitution: vX(I) = vX(I) * kWallFriction: vY(I) = vY(I) * kWallFriction

        If pX(I) > WW Then pX(I) = WW - (pX(I) - WW): vX(I) = -vX(I) * kRestitution: vY(I) = vY(I) * kWallFriction: vZ(I) = vZ(I) * kWallFriction
        If pY(I) > HH Then pY(I) = HH - (pY(I) - HH): vY(I) = -vY(I) * kRestitution: vX(I) = vX(I) * kWallFriction: vZ(I) = vZ(I) * kWallFriction
        If pZ(I) > ZZ Then pZ(I) = ZZ - (pZ(I) - ZZ): vZ(I) = -vZ(I) * kRestitution: vX(I) = vX(I) * kWallFriction: vY(I) = vY(I) * kWallFriction



        ' -------------------------------- FAKE boundary (density)  and (VEL)
        If pX(I) < h Then S = SmoothKernel_3(pX(I) * invH): Density(I) = Density(I) + S * kFakeDensity: VXChange(I) = VXChange(I) + S * invDT * kFakeVel
        If pY(I) < h Then S = SmoothKernel_3(pY(I) * invH): Density(I) = Density(I) + S * kFakeDensity: VYChange(I) = VYChange(I) + S * invDT * kFakeVel
        If pZ(I) < h Then S = SmoothKernel_3(pZ(I) * invH): Density(I) = Density(I) + S * kFakeDensity: VZChange(I) = VZChange(I) + S * invDT * kFakeVel

        If pX(I) > wwH Then S = SmoothKernel_3((WW - pX(I)) * invH): Density(I) = Density(I) + S * kFakeDensity: VXChange(I) = VXChange(I) - S * invDT * kFakeVel
        If pY(I) > HHH Then S = SmoothKernel_3((HH - pY(I)) * invH): Density(I) = Density(I) + S * kFakeDensity: VYChange(I) = VYChange(I) - S * invDT * kFakeVel
        If pZ(I) > zzH Then S = SmoothKernel_3((ZZ - pZ(I)) * invH): Density(I) = Density(I) + S * kFakeDensity: VZChange(I) = VZChange(I) - S * invDT * kFakeVel
        '----------------------------------------


        COMx = COMx + pX(I)
        COMy = COMy + pY(I)
        COMz = COMz + pZ(I)

    Next


    '        invNP = 1 / NP
    '        COMx = COMx * invNP
    '        COMy = COMy * invNP
    '        COMz = COMz * invNP



    If COMGravity Then
        invNP = 1# / NP
        COMx = COMx * invNP
        COMy = COMy * invNP
        COMz = COMz * invNP

        For I = 1 To NP
            DX = pX(I) - COMx
            DY = pY(I) - COMy
            DZ = pZ(I) - COMz
            D = DX * DX + DY * DY + DZ * DZ
            D = 1# / (1# + D)
            F = D * GravScale * NP * 0.0002
            vX(I) = vX(I) - DX * F
            vY(I) = vY(I) - DY * F
            vZ(I) = vZ(I) - DZ * F
        Next
    End If


End Sub



Public Sub SPH_ComputePAIRS()
    Dim pair      As Long

    Dim D         As Double
    Dim I         As Long
    Dim J         As Long
    Dim DX        As Double
    Dim DY        As Double
    Dim DZ        As Double

    Dim NormalizedDX As Double
    Dim NormalizedDY As Double
    Dim NormalizedDZ As Double

    Dim InvD      As Double
    Dim R         As Double
    Dim Smooth    As Double
    Dim VXcI      As Double
    Dim VYcI      As Double
    Dim VZcI      As Double

    Dim VXcJ      As Double
    Dim VYcJ      As Double
    Dim VZcJ      As Double

    Dim K         As Double
    Dim iX        As Double
    Dim iY        As Double
    Dim IZ        As Double

    Dim SmoothPRESS As Double
    Dim Pij       As Double

    Dim vDX       As Double
    Dim vDY       As Double
    Dim vDZ       As Double

    Dim OmR       As Double

Dim DOTvel As Double




    'PRE comute pairs .... only for DENSITY / Pressure
    '------------------------------------------- DENSITY
    For pair = 1 To RetNofPairs
        '        D = Sqr(arrD(pair))    ''''''''''''''<<<<<<<<<<  SQR
        '        arrD(pair) = D
        '-------------------

        D = h * SQR_Table(TABLESLength * arrD(pair) * InvH2)    '   Avoid SQR using a table

        '-------------------
        '        D = TABLESLength * arrD(pair) * InvH2
        '        If D >= 1.10492178673608E-02 Then
        '            D = h * SQR_Table(D)    '   Avoid SQR using a table
        '        Else
        '            D = Sqr(arrD(pair))
        '        End If
        '-------------------

        arrD(pair) = D

        '        If D Then
        I = P1(pair)
        J = P2(pair)
        R = D * invH
        '        Smooth = SmoothKernel_3(R)
        Smooth = SmoothKernel_Table(R * TABLESLength)
        Density(I) = Density(I) + Smooth
        Density(J) = Density(J) + Smooth

        '        End If
    Next
    '------------------------------------------- PRESSURE
    For I = 1 To NP
        Pressure(I) = (Density(I) - RestDensity) * INVRestDensity
        If Pressure(I) > PressureLimit Then
            Pressure(I) = PressureLimit
        ElseIf Pressure(I) < -PressureLimit Then
            Pressure(I) = -PressureLimit
        End If
        'Reset Density
        '        Density(I) = 0#  'move to SPH_MOVE

        If Density(I) > 0.0005 Then
            If Density(I) > TestMaxDens Then TestMaxDens = Density(I)
            '  INVDensity(I) = 1# / Density(I)
            If Density(I) > ExpectedMaxDensity Then Density(I) = ExpectedMaxDensity
            INVDensity(I) = InvDensity_Table(TABLESLength * Density(I) * INVExpectedMaxDensity)
        Else
            INVDensity(I) = 0#
        End If

    Next
    '---------------------------------------------


    ' main PAIRS computation

    For pair = 1 To RetNofPairs
        I = P1(pair)
        J = P2(pair)

        DX = arrDX(pair)
        DY = arrDY(pair)
        DZ = arrDZ(pair)

        D = arrD(pair)

        If D Then

            R = D * invH     ' the distance between particles in range 0-1
            OmR = 1# - R

            '            InvD = 1# / D ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            InvD = Normalize_Table(R * TABLESLength)    ' AVOID Division
            NormalizedDX = DX * InvD
            NormalizedDY = DY * InvD
            NormalizedDZ = DZ * InvD

            '----------------------------------------------------------------

            VXcI = VXChange(I)
            VYcI = VYChange(I)
            VZcI = VZChange(I)

            VXcJ = VXChange(J)
            VYcJ = VYChange(J)
            VZcJ = VZChange(J)


            If Phase(I) = Phase(J) Then

                ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<  ATTRACTION
                K = OmR * OmR * KAttraction
                iX = NormalizedDX * K
                iY = NormalizedDY * K
                IZ = NormalizedDZ * K

                VXcI = VXcI + iX
                VYcI = VYcI + iY
                VZcI = VZcI + IZ

                VXcJ = VXcJ - iX
                VYcJ = VYcJ - iY
                VZcJ = VZcJ - IZ

                '                Smooth = SmoothKernel_3(R)
                Smooth = SmoothKernel_Table(R * TABLESLength)

                ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<  PRESSURE

                SmoothPRESS = Smooth * OmR    ' V 1
                Pij = 0.5 * (Pressure(I) + Pressure(J)) * SmoothPRESS * KPressure
                iX = NormalizedDX * Pij
                iY = NormalizedDY * Pij
                IZ = NormalizedDZ * Pij

                VXcI = VXcI - iX
                VYcI = VYcI - iY
                VZcI = VZcI - IZ

                VXcJ = VXcJ + iX
                VYcJ = VYcJ + iY
                VZcJ = VZcJ + IZ


                ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<  VISCOSITY 1
                ''                ' Without Densities sum division (1st Version)
                ''                vDX = vX(J) - vX(I)
                ''                vDY = vY(J) - vY(I)
                ''                K = -0.5# * r * r * r + r * r + 0.5# * InvD * H - 1#
                ''                K = K * KViscosity
                ''                'particles are Separating  ?
                ''                If (dX * vDX + dY * vDY) < 0# Then K = K * 0.005#               '025#
                ''                If K > 1# Then K = 1#
                ''                iX = vDX * K
                ''                iY = vDY * K
                ''                VXcI = VXcI + iX
                ''                VYcI = VYcI + iY
                ''                VXcJ = VXcJ - iX
                ''                VYcJ = VYcJ - iY

                ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<  VISCOSITY 2
                ' Inverse proportional to Densities Sum
                ' K = KViscosity * OmR * OmR
                ' K = -0.5# * r * r * r + r * r + 1# / (2# * r) - 1#
                ' Same but without division:

                vDX = vX(J) - vX(I)
                vDY = vY(J) - vY(I)
                vDZ = vZ(J) - vZ(I)

                ' K = (-0.5 * R * R * R + R * R + 0.5 * InvD * h - 1#)* KViscosity
                K = Visco_Table(R * TABLESLength) * KViscosity

                ' MODE 2 -----------<<<<<<< difference from above
'''''Before 2023 OK
'''                K = K * 8.3 * (INVDensity(I) + INVDensity(J))
'''
'''                'particles are Separating  ?
'''                'If (DX * vDX + DY * vDY + DZ * vDZ) < 0# Then K = K * 0.001    '025#
'''
'''''                If (NormalizedDX * vDX + NormalizedDY * vDY + NormalizedDZ * vDZ) < 0# Then K = K * 0.001    '025#
'''''                If K > 0.5 Then K = 0.5

'DOTvel = (NormalizedDX * vDX + NormalizedDY * vDY + NormalizedDZ * vDZ)

                K = K * 15:             If K > 1# Then K = 1
      
                iX = vDX * K
                iY = vDY * K
                IZ = vDZ * K
'If DOTvel < 0 Then
                VXcI = VXcI + iX
                VYcI = VYcI + iY
                VZcI = VZcI + IZ
'Else
                VXcJ = VXcJ - iX
                VYcJ = VYcJ - iY
                VZcJ = VZcJ - IZ
'End If
            Else
                K = OmR * OmR * KAttraction * 26#
                iX = NormalizedDX * K
                iY = NormalizedDY * K
                IZ = NormalizedDZ * K

                VXcI = VXcI - iX
                VYcI = VYcI - iY
                VZcI = VZcI - IZ

                VXcJ = VXcJ + iX
                VYcJ = VYcJ + iY
                VZcJ = VZcJ + IZ

            End If


            '----------------------------------------------------------------

            VXChange(I) = VXcI
            VYChange(I) = VYcI
            VZChange(I) = VZcI

            VXChange(J) = VXcJ
            VYChange(J) = VYcJ
            VZChange(J) = VZcJ

            '----------------------------------------------------------------
        Else
            '        Beep
            '
            '
            '            MsgBox I & "   " & J & "   Same position"
            '            VXChange(I) = VXChange(I) + (Rnd * 2 - 1) * 0.1 * h * 9
            '            VYChange(I) = VYChange(I) + (Rnd * 2 - 1) * 0.1 * h * 9
            '            VZChange(I) = VZChange(I) + (Rnd * 2 - 1) * 0.1 * h * 9
            '
            '            VXChange(J) = VXChange(J) + (Rnd * 2 - 1) * 0.1 * h * 9
            '            VYChange(J) = VYChange(J) + (Rnd * 2 - 1) * 0.1 * h * 9
            '            VZChange(J) = VZChange(J) + (Rnd * 2 - 1) * 0.1 * h * 9

            pX(I) = pX(I) + (Rnd * 2 - 1) * 0.001 * h    '* 9
            pY(I) = pY(I) + (Rnd * 2 - 1) * 0.001 * h    '* 9
            pZ(I) = pZ(I) + (Rnd * 2 - 1) * 0.001 * h    '* 9

            pX(J) = pX(J) + (Rnd * 2 - 1) * 0.001 * h    '* 9
            pY(J) = pY(J) + (Rnd * 2 - 1) * 0.001 * h    '* 9
            pZ(J) = pZ(J) + (Rnd * 2 - 1) * 0.001 * h    '* 9


        End If


    Next

End Sub



Private Function SmoothKernel_1(ByVal R As Double) As Double
    SmoothKernel_1 = 1# - R * R * (3# - 2# * R)
End Function

Private Function SmoothKernel_2(ByVal R As Double) As Double
'A new kernel function for SPH with applications to free surfaceflowsqX.F. Yanga, S.L. Pengb, M.B. Liu
    SmoothKernel_2 = (4# * Cos(PI * R) + Cos(PI2 * R) + 3#) * 0.125
End Function

Public Function SmoothKernel_3(ByVal R As Double) As Double
''http://www.astro.lu.se/~david/teaching/SPH/notes/annurev.aa.30.090192.pdf
    R = R * 2#
    If R <= 1# Then
        SmoothKernel_3 = 1# - 1.5 * R * R + 0.75 * R * R * R
    Else
        R = 2# - R
        SmoothKernel_3 = 0.25 * R * R * R
    End If

    'SmoothKernel_3 = Exp(-R * R * 6.5)

End Function

Public Function SmoothKernel_4(ByVal R As Double) As Double
'https://www.desmos.com/calculator/o3hktwyuo5
    R = 1# - R

    SmoothKernel_4 = R * R * R * (6# * R * R - 15# * R + 10#)
End Function

