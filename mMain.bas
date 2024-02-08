Attribute VB_Name = "mMain"
Option Explicit


Public Const ExpectedMaxDensity As Single = 64
Attribute ExpectedMaxDensity.VB_VarUserMemId = 1073938433
Public Const INVExpectedMaxDensity As Single = 1 / ExpectedMaxDensity


Public OCTTREE    As cOctTree
Attribute OCTTREE.VB_VarUserMemId = 1073938435
'Public SpatialGRID As cSpatialGrid3D
''Public HASH3D As cSpatialHash3D


Public PIChDC     As Long
Attribute PIChDC.VB_VarUserMemId = 1073741825

Public WW         As Long
Attribute WW.VB_VarUserMemId = 1610809345
Public HH         As Long
Attribute HH.VB_VarUserMemId = 1073938435
Public ZZ         As Long
Attribute ZZ.VB_VarUserMemId = 1073938436
Public invZZ      As Single
Attribute invZZ.VB_VarUserMemId = 1073938437


Public px()       As Single
Attribute px.VB_VarUserMemId = 1073741830
Public py()       As Single
Attribute py.VB_VarUserMemId = 1610809346
Public pz()       As Single
Attribute pz.VB_VarUserMemId = 1073938440

Public vX()       As Single
Attribute vX.VB_VarUserMemId = 1073938442
Public vY()       As Single
Attribute vY.VB_VarUserMemId = 1073938443
Public vZ()       As Single
Attribute vZ.VB_VarUserMemId = 1073938444

Public pvX()      As Single
Attribute pvX.VB_VarUserMemId = 1073938445
Public pvY()      As Single
Attribute pvY.VB_VarUserMemId = 1073741837
Public pvZ()      As Single
Attribute pvZ.VB_VarUserMemId = 1073741838


Public NP         As Long
Attribute NP.VB_VarUserMemId = 1073938448

Public DrawPairs  As Boolean
Attribute DrawPairs.VB_VarUserMemId = 1610809349
Public DoLOOP     As Boolean
Attribute DoLOOP.VB_VarUserMemId = 1073741841
Public Frame      As Long
Attribute Frame.VB_VarUserMemId = 1073938451
Public DoSaveFrames As Boolean
Attribute DoSaveFrames.VB_VarUserMemId = 1073938452
Public rndGravity As Boolean
Attribute rndGravity.VB_VarUserMemId = 1610809350

Public DoFaucet1  As Boolean
Attribute DoFaucet1.VB_VarUserMemId = 1073938454
Public DoFaucet2  As Boolean
Attribute DoFaucet2.VB_VarUserMemId = 1073741846

Public COMGravity As Boolean
Attribute COMGravity.VB_VarUserMemId = 1073741847
Public GravScale  As Single
Attribute GravScale.VB_VarUserMemId = 1073741848


Public FPS        As Single
Attribute FPS.VB_VarUserMemId = 1073741849
Public CNT        As Long
Attribute CNT.VB_VarUserMemId = 1073741850
Public OldCNT     As Long
Attribute OldCNT.VB_VarUserMemId = 1073741851

Public mTime      As Single
Attribute mTime.VB_VarUserMemId = 1073741852
Public OldmTime   As Single
Attribute OldmTime.VB_VarUserMemId = 1073938455


Public RenderEvery As Long
Attribute RenderEvery.VB_VarUserMemId = 1073938457

Public H          As Single        'smoothing radius
Attribute H.VB_VarUserMemId = 1073741855
Public invH       As Single
Attribute invH.VB_VarUserMemId = 1073741856
Public InvH2      As Single
Attribute InvH2.VB_VarUserMemId = 1879244800
Public SQR_Table() As Single
Attribute SQR_Table.VB_VarUserMemId = 1879244836
Public Normalize_Table() As Single
Attribute Normalize_Table.VB_VarUserMemId = 1879244864
Public SmoothKernel_Table() As Single
Attribute SmoothKernel_Table.VB_VarUserMemId = 1073741860
Public InvDensity_Table() As Single
Attribute InvDensity_Table.VB_VarUserMemId = 1073741861
Public Visco_Table() As Single
Attribute Visco_Table.VB_VarUserMemId = 1073741862


Public Const TABLESLength As Single = 2 ^ 14 - 1    ' 2 ^ 14 - 1

Public Const kRestitution As Single = 0.7    ' 0.65    ' 0.65    '0.75    ' 0.75 '0.5    '0.66
Public Const kWallFriction As Single = 0.99    '0.98 '0.996    '0.995
Public Const kFakeDensity As Single = 2    ' 2    '0.3  '2022
Public Const kFakeVel As Single = 0.01    ' 0.005



Public gX         As Single        'GRAVITY
Attribute gX.VB_VarUserMemId = 1073741863
Public gY         As Single
Attribute gY.VB_VarUserMemId = 1073741864
Public gZ         As Single
Attribute gZ.VB_VarUserMemId = 1073741865

Public gTOX       As Single
Attribute gTOX.VB_VarUserMemId = 1073741866
Public gTOY       As Single
Attribute gTOY.VB_VarUserMemId = 1073741867
Public gTOZ       As Single
Attribute gTOZ.VB_VarUserMemId = 1073741868


' SPH ---------------------------------------------------------

Public DT         As Single
Attribute DT.VB_VarUserMemId = 1073741869
Public invDT      As Single
Attribute invDT.VB_VarUserMemId = 1073741870


Private RestDensity As Single
Attribute RestDensity.VB_VarUserMemId = 1073741871
Private INVRestDensity As Single
Attribute INVRestDensity.VB_VarUserMemId = 1073741872
Private PressureLimit As Single
Attribute PressureLimit.VB_VarUserMemId = 1073741873
Private KAttraction As Single
Attribute KAttraction.VB_VarUserMemId = 1073741874
Private KPressure As Single
Attribute KPressure.VB_VarUserMemId = 1073741875
Private KViscosity As Single
Attribute KViscosity.VB_VarUserMemId = 1610809345


Public Density()  As Single
Attribute Density.VB_VarUserMemId = 1610809351
Public Pressure() As Single
Attribute Pressure.VB_VarUserMemId = 1073741878
Public Phase()    As Long
Attribute Phase.VB_VarUserMemId = 1073741879
Public INVDensity() As Single
Attribute INVDensity.VB_VarUserMemId = 1073741880



Public VXChange() As Single
Attribute VXChange.VB_VarUserMemId = 1073741881
Public VYChange() As Single
Attribute VYChange.VB_VarUserMemId = 1073741882
Public VZChange() As Single
Attribute VZChange.VB_VarUserMemId = 1073741883


Public P1()       As Long
Attribute P1.VB_VarUserMemId = 1073741884
Public P2()       As Long
Attribute P2.VB_VarUserMemId = 1073741885
Public arrDX()    As Single
Attribute arrDX.VB_VarUserMemId = 1073741886
Public arrDY()    As Single
Attribute arrDY.VB_VarUserMemId = 1073741887
Public arrDZ()    As Single
Attribute arrDZ.VB_VarUserMemId = 1073741888

Public arrD()     As Single
Attribute arrD.VB_VarUserMemId = 1073741889
Public RetNofPairs As Long
Attribute RetNofPairs.VB_VarUserMemId = 1610809352
Public MaxNofPairs As Long
Attribute MaxNofPairs.VB_VarUserMemId = 1073741891


Public CAMERA     As c3DEasyCam
Attribute CAMERA.VB_VarUserMemId = 1073741892

Public COMx       As Single
Attribute COMx.VB_VarUserMemId = 1073741893
Public COMy       As Single
Attribute COMy.VB_VarUserMemId = 1073741894
Public COMz       As Single
Attribute COMz.VB_VarUserMemId = 1073741895

Public CamRot     As Boolean
Attribute CamRot.VB_VarUserMemId = 1073741896


Public TestMaxDens As Single
Attribute TestMaxDens.VB_VarUserMemId = 1073741897




Public Sub SPH_InitConst()
    Dim R         As Single
    Dim kernelWeight As Single
    Dim I         As Single

    R = 0.33333333333
    '    RestDensity = SmoothKernel_3(r) * 6    ' 2D
    '    RestDensity = SmoothKernel_3(R) * 6     ' 3D
    RestDensity = SmoothKernel_3(R) * 4.8    '5    '5      '2023

    '    RestDensity = SmoothKernel_3(R) * 5 * 3 'Without Attraction

    For R = 0 To 1 Step 0.001
        I = I + 1
        kernelWeight = kernelWeight + SmoothKernel_3(R)
    Next

    'kernelWeight = kernelWeight / (I + 1)
    kernelWeight = kernelWeight / (I)


    INVRestDensity = 1 / RestDensity
    '    PressureLimit = 400      '200 '100    '50    '45 '20
    PressureLimit = 500            '800 '2022

    DT = 0.25
    invDT = 1 / DT

    'KAttraction = 0.0128 * invDT
    KAttraction = 0.0128 * invDT * 0.65    ' 0.72    '0.75    ' 0.85    '2023
    'KAttraction = 0 'Without Attraction
    '    'KPressure = kernelWeight * 0.08 * invDT
    KPressure = kernelWeight * 0.15    ' 0.12 '0.15    '0.12    '0.15 * invDT '2022

    KPressure = KPressure * 2.5    '* 1.5
    'KPressure = KPressure * 80 'Without Attraction

    'KViscosity = 0.018 * 0.8
    KViscosity = 0.018 * 0.5       '1 '1.5 ' 0.66    '0.5    '2022

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


        If I Then Normalize_Table(I) = 1 / (H * I / TABLESLength)
        SmoothKernel_Table(I) = SmoothKernel_3(I / TABLESLength)
        If I Then InvDensity_Table(I) = 1 / (ExpectedMaxDensity * (I / TABLESLength))    '

        R = I / TABLESLength
        '        If R Then Visco_Table(I) = -0.5 * R * R * R + R * R + 0.5 / R - 1
        If R Then Visco_Table(I) = (1 - R) ^ 5    '2023

    Next

End Sub

Public Sub SPH_MOVE()
    Dim I         As Long





    Dim invNP     As Single
    Dim DX        As Single
    Dim DY        As Single
    Dim DZ        As Single
    Dim D         As Single
    Dim F         As Single

    Dim wwH As Single, HHH As Single, zzH As Single
    Dim S         As Single

    wwH = WW - H
    HHH = HH - H
    zzH = ZZ - H


    For I = 1 To NP
        vX(I) = vX(I) + VXChange(I) + gX * DT
        vY(I) = vY(I) + VYChange(I) + gY * DT
        vZ(I) = vZ(I) + VZChange(I) + gZ * DT

        VXChange(I) = 0
        VYChange(I) = 0
        VZChange(I) = 0

        vX(I) = vX(I) * 0.9995     ' 0.998
        vY(I) = vY(I) * 0.9995     ' 0.998
        vZ(I) = vZ(I) * 0.9995     ' 0.998

        px(I) = px(I) + vX(I) * DT
        py(I) = py(I) + vY(I) * DT
        pz(I) = pz(I) + vZ(I) * DT


        Density(I) = 0


        If px(I) < 0 Then px(I) = -px(I): vX(I) = -vX(I) * kRestitution: vY(I) = vY(I) * kWallFriction: vZ(I) = vZ(I) * kWallFriction
        If py(I) < 0 Then py(I) = -py(I): vY(I) = -vY(I) * kRestitution: vX(I) = vX(I) * kWallFriction: vZ(I) = vZ(I) * kWallFriction
        If pz(I) < 0 Then pz(I) = -pz(I): vZ(I) = -vZ(I) * kRestitution: vX(I) = vX(I) * kWallFriction: vY(I) = vY(I) * kWallFriction

        If px(I) > WW Then px(I) = WW - (px(I) - WW): vX(I) = -vX(I) * kRestitution: vY(I) = vY(I) * kWallFriction: vZ(I) = vZ(I) * kWallFriction
        If py(I) > HH Then py(I) = HH - (py(I) - HH): vY(I) = -vY(I) * kRestitution: vX(I) = vX(I) * kWallFriction: vZ(I) = vZ(I) * kWallFriction
        If pz(I) > ZZ Then pz(I) = ZZ - (pz(I) - ZZ): vZ(I) = -vZ(I) * kRestitution: vX(I) = vX(I) * kWallFriction: vY(I) = vY(I) * kWallFriction



        ' -------------------------------- FAKE boundary (density)  and (VEL)
        If px(I) < H Then S = SmoothKernel_3(px(I) * invH): Density(I) = Density(I) + S * kFakeDensity: VXChange(I) = VXChange(I) + S * invDT * kFakeVel
        If py(I) < H Then S = SmoothKernel_3(py(I) * invH): Density(I) = Density(I) + S * kFakeDensity: VYChange(I) = VYChange(I) + S * invDT * kFakeVel
        If pz(I) < H Then S = SmoothKernel_3(pz(I) * invH): Density(I) = Density(I) + S * kFakeDensity: VZChange(I) = VZChange(I) + S * invDT * kFakeVel

        If px(I) > wwH Then S = SmoothKernel_3((WW - px(I)) * invH): Density(I) = Density(I) + S * kFakeDensity: VXChange(I) = VXChange(I) - S * invDT * kFakeVel
        If py(I) > HHH Then S = SmoothKernel_3((HH - py(I)) * invH): Density(I) = Density(I) + S * kFakeDensity: VYChange(I) = VYChange(I) - S * invDT * kFakeVel
        If pz(I) > zzH Then S = SmoothKernel_3((ZZ - pz(I)) * invH): Density(I) = Density(I) + S * kFakeDensity: VZChange(I) = VZChange(I) - S * invDT * kFakeVel
        '----------------------------------------


        COMx = COMx + px(I)
        COMy = COMy + py(I)
        COMz = COMz + pz(I)

    Next


    '        invNP = 1 / NP
    '        COMx = COMx * invNP
    '        COMy = COMy * invNP
    '        COMz = COMz * invNP



    If COMGravity Then
        invNP = 1 / NP
        COMx = COMx * invNP
        COMy = COMy * invNP
        COMz = COMz * invNP

        For I = 1 To NP
            DX = px(I) - COMx
            DY = py(I) - COMy
            DZ = pz(I) - COMz
            D = DX * DX + DY * DY + DZ * DZ
            D = 1 / (1 + D)
            F = D * GravScale * NP * 0.0002
            vX(I) = vX(I) - DX * F
            vY(I) = vY(I) - DY * F
            vZ(I) = vZ(I) - DZ * F
        Next
    End If



End Sub



Public Sub SPH_ComputePAIRS()
    Dim pair      As Long

    Dim D         As Single
    Dim I         As Long
    Dim J         As Long
    Dim DX        As Single
    Dim DY        As Single
    Dim DZ        As Single

    Dim NormalizedDX As Single
    Dim NormalizedDY As Single
    Dim NormalizedDZ As Single

    Dim InvD      As Single
    Dim R         As Single
    Dim Smooth    As Single
    Dim VXcI      As Single
    Dim VYcI      As Single
    Dim VZcI      As Single

    Dim VXcJ      As Single
    Dim VYcJ      As Single
    Dim VZcJ      As Single

    Dim K         As Single
    Dim iX        As Single
    Dim iY        As Single
    Dim IZ        As Single

    Dim SmoothPRESS As Single
    Dim Pij       As Single

    Dim vDX       As Single
    Dim vDY       As Single
    Dim vDZ       As Single

    Dim OmR       As Single

    Dim DOTvel    As Single




    'PRE comute pairs .... only for DENSITY / Pressure
    '------------------------------------------- DENSITY
    For pair = 1 To RetNofPairs
        '        D = Sqr(arrD(pair))    ''''''''''''''<<<<<<<<<<  SQR
        '        arrD(pair) = D
        '-------------------

        D = H * SQR_Table(TABLESLength * arrD(pair) * InvH2)    '   Avoid SQR using a table

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
        '        Density(I) = 0   'move to SPH_MOVE

        If Density(I) > 0.0005 Then
            If Density(I) > TestMaxDens Then TestMaxDens = Density(I)
            '  INVDensity(I) = 1  / Density(I)
            If Density(I) > ExpectedMaxDensity Then Density(I) = ExpectedMaxDensity
            INVDensity(I) = InvDensity_Table(TABLESLength * Density(I) * INVExpectedMaxDensity)
        Else
            INVDensity(I) = 0
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

            R = D * invH           ' the distance between particles in range 0-1
            OmR = 1 - R

            '            InvD = 1  / D ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
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
                ''                K = -0.5  * r * r * r + r * r + 0.5  * InvD * H - 1
                ''                K = K * KViscosity
                ''                'particles are Separating  ?
                ''                If (dX * vDX + dY * vDY) < 0  Then K = K * 0.005                '025
                ''                If K > 1  Then K = 1
                ''                iX = vDX * K
                ''                iY = vDY * K
                ''                VXcI = VXcI + iX
                ''                VYcI = VYcI + iY
                ''                VXcJ = VXcJ - iX
                ''                VYcJ = VYcJ - iY

                ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<  VISCOSITY 2
                ' Inverse proportional to Densities Sum
                ' K = KViscosity * OmR * OmR
                ' K = -0.5  * r * r * r + r * r + 1  / (2  * r) - 1
                ' Same but without division:

                vDX = vX(J) - vX(I)
                vDY = vY(J) - vY(I)
                vDZ = vZ(J) - vZ(I)

                ' K = (-0.5 * R * R * R + R * R + 0.5 * InvD * h - 1 )* KViscosity
                K = Visco_Table(R * TABLESLength) * KViscosity

                ' MODE 2 -----------<<<<<<< difference from above
                '''''Before 2023 OK
                '''                K = K * 8.3 * (INVDensity(I) + INVDensity(J))
                '''
                '''                'particles are Separating  ?
                '''                'If (DX * vDX + DY * vDY + DZ * vDZ) < 0  Then K = K * 0.001    '025
                '''
                '''''                If (NormalizedDX * vDX + NormalizedDY * vDY + NormalizedDZ * vDZ) < 0  Then K = K * 0.001    '025
                '''''                If K > 0.5 Then K = 0.5

                'DOTvel = (NormalizedDX * vDX + NormalizedDY * vDY + NormalizedDZ * vDZ)

                '* 15
                K = K * 15: If K > 1 Then K = 1

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
                K = OmR * OmR * KAttraction * 26
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

            px(I) = px(I) + (Rnd * 2 - 1) * 0.001 * H    '* 9
            py(I) = py(I) + (Rnd * 2 - 1) * 0.001 * H    '* 9
            pz(I) = pz(I) + (Rnd * 2 - 1) * 0.001 * H    '* 9

            px(J) = px(J) + (Rnd * 2 - 1) * 0.001 * H    '* 9
            py(J) = py(J) + (Rnd * 2 - 1) * 0.001 * H    '* 9
            pz(J) = pz(J) + (Rnd * 2 - 1) * 0.001 * H    '* 9


        End If


    Next

End Sub



Private Function SmoothKernel_1(ByVal R As Single) As Single
    SmoothKernel_1 = 1 - R * R * (3 - 2 * R)
End Function

Private Function SmoothKernel_2(ByVal R As Single) As Single
    'A new kernel function for SPH with applications to free surfaceflowsqX.F. Yanga, S.L. Pengb, M.B. Liu
    SmoothKernel_2 = (4 * Cos(PI * R) + Cos(PI2 * R) + 3) * 0.125
End Function

Public Function SmoothKernel_3(ByVal R As Single) As Single
    ''http://www.astro.lu.se/~david/teaching/SPH/notes/annurev.aa.30.090192.pdf
    R = R * 2
    If R <= 1 Then
        SmoothKernel_3 = 1 - 1.5 * R * R + 0.75 * R * R * R
    Else
        R = 2 - R
        SmoothKernel_3 = 0.25 * R * R * R
    End If

    'SmoothKernel_3 = Exp(-R * R * 6.5)

End Function

Public Function SmoothKernel_4(ByVal R As Single) As Single
    'https://www.desmos.com/calculator/o3hktwyuo5
    R = 1 - R

    SmoothKernel_4 = R * R * R * (6 * R * R - 15 * R + 10)
End Function

