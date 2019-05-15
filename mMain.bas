Attribute VB_Name = "mMain"
Option Explicit




Public SpatialGRID As cSpatialGrid3D

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


Public FPS        As Long
Public CNT        As Long
Public OldCNT     As Long

Public RenderEvery As Long

Public H          As Double    'smoothing radius
Public invH       As Double



Public gX         As Double    'GRACITY
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



Public CAMERA     As c3DEasyCam



Public Sub SPH_InitConst()
    Dim r         As Double
    Dim kernelWeight As Double
    Dim I         As Double

    r = 0.33333333333
    '    RestDensity = SmoothKernel_3(r) * 6    ' 2D

    RestDensity = SmoothKernel_3(r) * 6    ' 3D


    For r = 0 To 1 Step 0.01
        I = I + 1
        kernelWeight = kernelWeight + SmoothKernel_3(r)
    Next
    kernelWeight = kernelWeight / I

    INVRestDensity = 1 / RestDensity
    PressureLimit = 50    '45 '20

    DT = 0.25
    invDT = 1 / DT

    KAttraction = 0.0128 * invDT
    KPressure = kernelWeight * 0.08 * invDT
    KViscosity = 0.018

    ReDim VXChange(NP)
    ReDim VYChange(NP)
    ReDim VZChange(NP)

    ReDim Density(NP)
    ReDim Pressure(NP)
    ReDim INVDensity(NP)

    ReDim Phase(NP)


End Sub

Public Sub SPH_MOVE()
    Dim I         As Long

    Const kRestitution As Double = 0.85
    Const kFakeDensity As Double = 0.3
    Const kFakeVel As Double = 0.005


    Dim COMx      As Double
    Dim COMy      As Double
    Dim COMz      As Double
    Dim invNP     As Double
    Dim dx        As Double
    Dim DY        As Double
    Dim dz        As Double
    Dim D         As Double
    Dim F         As Double

    Dim wwH!, hhH!, zzH!
    Dim S!

    wwH = WW - H
    hhH = HH - H
    zzH = ZZ - H


    For I = 1 To NP
        vX(I) = vX(I) + VXChange(I) + gX * DT
        vY(I) = vY(I) + VYChange(I) + gY * DT
        vZ(I) = vZ(I) + VZChange(I) + gZ * DT

        VXChange(I) = 0!
        VYChange(I) = 0!
        VZChange(I) = 0!

        vX(I) = vX(I) * 0.999     ' 0.998!
        vY(I) = vY(I) * 0.999     ' 0.998!
        vZ(I) = vZ(I) * 0.999     ' 0.998!

        pX(I) = pX(I) + vX(I) * DT
        pY(I) = pY(I) + vY(I) * DT
        pZ(I) = pZ(I) + vZ(I) * DT


        Density(I) = 0

        If pX(I) < 0! Then pX(I) = -pX(I): vX(I) = -vX(I) * kRestitution
        If pY(I) < 0! Then pY(I) = -pY(I): vY(I) = -vY(I) * kRestitution
        If pZ(I) < 0! Then pZ(I) = -pZ(I): vZ(I) = -vZ(I) * kRestitution

        If pX(I) > WW Then pX(I) = WW - (pX(I) - WW): vX(I) = -vX(I) * kRestitution
        If pY(I) > HH Then pY(I) = HH - (pY(I) - HH): vY(I) = -vY(I) * kRestitution
        If pZ(I) > ZZ Then pZ(I) = ZZ - (pZ(I) - ZZ): vZ(I) = -vZ(I) * kRestitution



        ' -------------------------------- FAKE boundary (density)  and (VEL)
        If pX(I) < H Then S = SmoothKernel_3(pX(I) * invH): Density(I) = Density(I) + S * kFakeDensity: VXChange(I) = VXChange(I) + S * invDT * kFakeVel
        If pY(I) < H Then S = SmoothKernel_3(pY(I) * invH): Density(I) = Density(I) + S * kFakeDensity: VYChange(I) = VYChange(I) + S * invDT * kFakeVel
        If pZ(I) < H Then S = SmoothKernel_3(pZ(I) * invH): Density(I) = Density(I) + S * kFakeDensity: VZChange(I) = VZChange(I) + S * invDT * kFakeVel

        If pX(I) > wwH Then S = SmoothKernel_3((WW - pX(I)) * invH): Density(I) = Density(I) + S * kFakeDensity: VXChange(I) = VXChange(I) - S * invDT * kFakeVel
        If pY(I) > hhH Then S = SmoothKernel_3((HH - pY(I)) * invH): Density(I) = Density(I) + S * kFakeDensity: VYChange(I) = VYChange(I) - S * invDT * kFakeVel
        If pZ(I) > zzH Then S = SmoothKernel_3((ZZ - pZ(I)) * invH): Density(I) = Density(I) + S * kFakeDensity: VZChange(I) = VZChange(I) - S * invDT * kFakeVel
        '----------------------------------------




        COMx = COMx + pX(I)
        COMy = COMy + pY(I)
        COMz = COMz + pZ(I)

    Next


    If COMGravity Then
        invNP = 1 / NP
        COMx = COMx * invNP
        COMy = COMy * invNP
        COMz = COMz * invNP

        For I = 1 To NP
            dx = pX(I) - COMx
            DY = pY(I) - COMy
            dz = pZ(I) - COMz
            D = dx * dx + DY * DY + dz * dz
            D = 1 / (1 + D)
            F = D * GravScale * NP * 0.0002
            vX(I) = vX(I) - dx * F
            vY(I) = vY(I) - DY * F
            vZ(I) = vZ(I) - dz * F
        Next
    End If


End Sub



Public Sub SPH_ComputePAIRS()
    Dim pair      As Long

    Dim D         As Double
    Dim I         As Long
    Dim J         As Long
    Dim dx        As Double
    Dim DY        As Double
    Dim dz        As Double

    Dim NormalizedDX As Double
    Dim NormalizedDY As Double
    Dim NormalizedDZ As Double

    Dim InvD      As Double
    Dim r         As Double
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



    'PRE comute pairs .... only for DENSITY / Pressure
    '------------------------------------------- DENSITY
    For pair = 1 To RetNofPairs
        arrD(pair) = Sqr(arrD(pair))    ''''''''''''''<<<<<<<<<<  SQR
        D = arrD(pair)
        If D Then
            I = P1(pair)
            J = P2(pair)
            r = D * invH
            Smooth = SmoothKernel_3(r)
            Density(I) = Density(I) + Smooth
            Density(J) = Density(J) + Smooth
        End If
    Next
    '------------------------------------------- PRESSURE
    For I = 1 To NP
        Pressure(I) = (Density(I) - RestDensity) * INVRestDensity
        If Pressure(I) > PressureLimit Then
            Pressure(I) = PressureLimit
        ElseIf Pressure(I) < -0 Then
            Pressure(I) = -0
        End If
        'Reset Density
        '        Density(I) = 0!  'move to SPH_MOVE

        If Density(I) > 0.001 Then
            INVDensity(I) = 1 / (Density(I))

            '            INVDensity(I) = 1 / (Density(I) * Density(I))

        Else
            INVDensity(I) = 0!
        End If

    Next
    '---------------------------------------------


    ' main PAIRS computation

    For pair = 1 To RetNofPairs
        I = P1(pair)
        J = P2(pair)

        dx = arrDX(pair)
        DY = arrDY(pair)
        dz = arrDZ(pair)

        D = arrD(pair)

        If D Then

            r = D * invH    ' the distance between particles in range 0-1
            OmR = 1! - r

            InvD = 1! / D
            NormalizedDX = dx * InvD
            NormalizedDY = DY * InvD
            NormalizedDZ = dz * InvD

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

                Smooth = SmoothKernel_3(r)


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
                ''                K = -0.5! * r * r * r + r * r + 0.5! * InvD * H - 1!
                ''                K = K * KViscosity
                ''                'particles are Separating  ?
                ''                If (dX * vDX + dY * vDY) < 0! Then K = K * 0.005!               '025!
                ''                If K > 1! Then K = 1!
                ''                iX = vDX * K
                ''                iY = vDY * K
                ''                VXcI = VXcI + iX
                ''                VYcI = VYcI + iY
                ''                VXcJ = VXcJ - iX
                ''                VYcJ = VYcJ - iY

                ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<  VISCOSITY 2
                ' Inverse proportional to Densities Sum
                ' K = KViscosity * OmR * OmR
                ' K = -0.5! * r * r * r + r * r + 1! / (2! * r) - 1!
                ' Same but without division:

                vDX = vX(J) - vX(I)
                vDY = vY(J) - vY(I)
                vDZ = vZ(J) - vZ(I)

                K = -0.5 * r * r * r + r * r + 0.5 * InvD * H - 1!
                K = K * KViscosity

                ' MODE 2 -----------<<<<<<< difference from above

                K = K * 8.3 * (INVDensity(I) + INVDensity(J))


                'particles are Separating  ?
                If (dx * vDX + DY * vDY + dz * vDZ) < 0! Then K = K * 0.001            '025!
                If K > 0.5 Then K = 0.5
                iX = vDX * K
                iY = vDY * K
                IZ = vDZ * K

                VXcI = VXcI + iX
                VYcI = VYcI + iY
                VZcI = VZcI + IZ

                VXcJ = VXcJ - iX
                VYcJ = VYcJ - iY
                VZcJ = VZcJ - IZ

            Else
                K = OmR * OmR * KAttraction * 26!
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
            VXChange(I) = VXChange(I) + (Rnd * 2 - 1) * 0.0001 * H
            VYChange(I) = VYChange(I) + (Rnd * 2 - 1) * 0.0001 * H
            VZChange(I) = VZChange(I) + (Rnd * 2 - 1) * 0.0001 * H

            VXChange(J) = VXChange(J) + (Rnd * 2 - 1) * 0.0001 * H
            VYChange(J) = VYChange(J) + (Rnd * 2 - 1) * 0.0001 * H
            VZChange(J) = VZChange(J) + (Rnd * 2 - 1) * 0.0001 * H

        End If


    Next

End Sub



Private Function SmoothKernel_1(ByVal r As Double) As Double
    SmoothKernel_1 = 1! - r * r * (3! - 2! * r)
End Function

Private Function SmoothKernel_2(ByVal r As Double) As Double
'A new kernel function for SPH with applications to free surfaceflowsqX.F. Yanga, S.L. Pengb, M.B. Liu
    SmoothKernel_2 = (4! * Cos(PI * r) + Cos(PI2 * r) + 3!) * 0.125
End Function

Public Function SmoothKernel_3(ByVal r As Double) As Double
'http://www.astro.lu.se/~david/teaching/SPH/notes/annurev.aa.30.090192.pdf
    r = r * 2!
    If r <= 1! Then
        SmoothKernel_3 = 1! - 1.5 * r * r + 0.75 * r * r * r
    Else
        r = 2! - r
        SmoothKernel_3 = 0.25 * r * r * r
    End If
End Function



