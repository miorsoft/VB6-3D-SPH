VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSpatialGrid3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type tCELL
    pIDX()        As Long
    NP            As Long
    MaxNP         As Long
    myIDX         As Long

End Type


Private GridSize  As Long

Private MAXCellX  As Long
Private MAXCellY  As Long
Attribute MAXCellY.VB_VarUserMemId = 1610809345
Private MAXCellZ  As Long
Attribute MAXCellZ.VB_VarUserMemId = 1073938435
Private MinCellX  As Long
Attribute MinCellX.VB_VarUserMemId = 1073938436
Private MinCellY  As Long
Attribute MinCellY.VB_VarUserMemId = 1073938437
Private MinCellZ  As Long
Attribute MinCellZ.VB_VarUserMemId = 1073938438


Private CELL()    As tCELL
Attribute CELL.VB_VarUserMemId = 1610809346
Private NP        As Long
Attribute NP.VB_VarUserMemId = 1073938440
Private MaxNP     As Long
Attribute MaxNP.VB_VarUserMemId = 1073938441
Private px()      As Single
Attribute px.VB_VarUserMemId = 1073938442
Private py()      As Single
Attribute py.VB_VarUserMemId = 1073938443
Private pz()      As Single
Attribute pz.VB_VarUserMemId = 1073938444

Private MaxDistance2 As Single
Attribute MaxDistance2.VB_VarUserMemId = 1073938445


Private PairP1()  As Long
Attribute PairP1.VB_VarUserMemId = 1610809348
Private PairP2()  As Long
Attribute PairP2.VB_VarUserMemId = 1073938447
Private PairDX()  As Single
Attribute PairDX.VB_VarUserMemId = 1073938448
Private PairDY()  As Single
Attribute PairDY.VB_VarUserMemId = 1073938449
Private PairDZ()  As Single
Attribute PairDZ.VB_VarUserMemId = 1610809349

Private PairD()   As Single
Attribute PairD.VB_VarUserMemId = 1073938451
Private Npairs    As Long
Attribute Npairs.VB_VarUserMemId = 1073938452
Private MAXNpair  As Long
Attribute MAXNpair.VB_VarUserMemId = 1610809350

Private CountRN   As Long
Attribute CountRN.VB_VarUserMemId = 1073938454



''Private XX()      As Long
''Private YY()      As Long
''Private ZZ()      As Long
''Private Ninc      As Long
''Private ZZpF()    As Long
''Private ZZpT()    As Long


Private iI        As Long
Attribute iI.VB_VarUserMemId = 1073938455

Private PXI       As Single
Attribute PXI.VB_VarUserMemId = 1073938456
Private PYI       As Single
Attribute PYI.VB_VarUserMemId = 1073938457
Private PZI       As Single
Attribute PZI.VB_VarUserMemId = 1073938458


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal ByteLen As Long)
Private Declare Function GetMem4 Lib "MSVBVM60" (ByVal Src As Long, ByVal dest As Long) As Long
Private Declare Function ArrPtr Lib "MSVBVM60" Alias "VarPtr" (arr() As Any) As Long




Private Sub pvArrCopySingle(dest() As Single, Src() As Single)
    Dim Size      As Long
    Dim W         As Long
    '    Dim H         As Long
    '    W = UBound(Src, 1)
    '    H = UBound(Src, 2)
    '    Size = (W + 1) * (H + 1) * LenB(Src(0, 0))
    If pvArrayExists(ArrPtr(Src)) Then

        W = UBound(Src)
        Size = (W + 1) * LenB(Src(0))

        If pvArrayExists(ArrPtr(dest)) Then
            '        If (W - UBound(dest, 1)) Or (H - UBound(dest, 2)) Then
            If (W - UBound(dest)) > 0 Then
                'ReDim dest(W, H)
                ReDim dest(W)
            End If
        Else
            '''Array DEST has No Dimension
            'ReDim dest(W, H)
            ReDim dest(W)
        End If

        '    CopyMemory ByVal VarPtr(dest(0, 0)), ByVal VarPtr(Src(0, 0)), Size
        CopyMemory ByVal VarPtr(dest(0)), ByVal VarPtr(Src(0)), Size
    End If

End Sub

Private Sub pvArrCopyLong(dest() As Long, Src() As Long)
    Dim Size      As Long
    Dim W         As Long
    '    Dim H         As Long
    '    W = UBound(Src, 1)
    '    H = UBound(Src, 2)
    '    Size = (W + 1) * (H + 1) * LenB(Src(0, 0))

    If pvArrayExists(ArrPtr(Src)) Then

        W = UBound(Src)
        Size = (W + 1) * LenB(Src(0))

        If pvArrayExists(ArrPtr(dest)) Then
            '        If (W - UBound(dest, 1)) Or (H - UBound(dest, 2)) Then
            If (W - UBound(dest)) > 0 Then
                '            Debug.Print W
                '            Stop

                'ReDim dest(W, H)
                ReDim dest(W)
            End If
        Else
            '''Array DEST has No Dimension
            'ReDim dest(W, H)
            ReDim dest(W)
        End If

        '    CopyMemory ByVal VarPtr(dest(0, 0)), ByVal VarPtr(Src(0, 0)), Size
        CopyMemory ByVal VarPtr(dest(0)), ByVal VarPtr(Src(0)), Size
    End If

End Sub
Private Function pvArrayExists(ByVal ppArray As Long) As Long
    GetMem4 ppArray, VarPtr(pvArrayExists)
End Function



Public Sub Init(WorldW As Long, WorldH As Long, worldZ As Long, maxDist As Long)

    GridSize = maxDist
    MaxDistance2 = maxDist * maxDist

    MAXCellX = WorldW \ GridSize
    MAXCellY = WorldH \ GridSize
    MAXCellZ = worldZ \ GridSize

    MinCellX = 0
    MinCellY = 0
    MinCellZ = 0

    '-1 / +1 because of pvX:
    MAXCellX = MAXCellX + 1
    MAXCellY = MAXCellY + 1
    MAXCellZ = MAXCellZ + 1
    MinCellX = MinCellX - 1
    MinCellY = MinCellY - 1
    MinCellZ = MinCellZ - 1


    '-1 / +1  GetPairsWDist4 need to enlarge to avoid boundary check)
    MAXCellX = MAXCellX + 1
    MAXCellY = MAXCellY + 1
    MAXCellZ = MAXCellZ + 1
    MinCellX = MinCellX - 1
    MinCellY = MinCellY - 1
    MinCellZ = MinCellZ - 1




    ' ReDim CELL(0 To MAXCellX, 0 To MAXCellY, 0 To MAXCellZ) '
    ReDim CELL(MinCellX To MAXCellX, MinCellY To MAXCellY, MinCellZ To MAXCellZ)

    NP = 0
    Npairs = 0


    '''    '--------------------------------------------------------------
    '''    Dim x1&, Y1&, z1&
    '''    Dim X2&, Y2&, z2&
    '''
    '''    ReDim XX(1 To 8)
    '''    ReDim YY(1 To 8)
    '''    ReDim ZZ(1 To 8)
    '''
    '''
    '''    Ninc = 0
    '''
    '''    For x1 = 0 To 1
    '''        For Y1 = 0 To 1
    '''            For z1 = 0 To 1
    '''                Ninc = Ninc + 1
    '''                XX(Ninc) = x1
    '''                YY(Ninc) = Y1
    '''                ZZ(Ninc) = z1
    '''            Next
    '''        Next
    '''    Next
    '''
    '''    Dim I&, J&
    '''    Ninc = 0
    '''    For I = 1 To 7
    '''        For J = I + 1 To 8
    '''
    '''            If Not (ZZpE(I, J)) Then
    '''                Ninc = Ninc + 1: ReDim Preserve ZZpF(Ninc): ReDim Preserve ZZpT(Ninc)
    '''                ZZpF(Ninc) = I
    '''                ZZpT(Ninc) = J
    '''                Debug.Print Ninc, , XX(I), YY(I), ZZ(I), , XX(J), YY(J), ZZ(J)
    ''''                Stop
    '''            End If
    '''        Next
    '''    Next

    TESTdot


End Sub
'''Private Function ZZpE(I&, J&) As Boolean
'''    Dim K      As Long
'''
'''    For K = 1 To Ninc
'''        If ZZpF(K) = J And ZZpT(K) = I Then ZZpE = True
'''        If ZZpF(K) = I And ZZpT(K) = J Then ZZpE = True
'''
'''        If Ninc >= 7 Then
'''
'''            If Abs(XX(I) - XX(J)) + Abs(YY(I) - YY(J)) + Abs(ZZ(I) - ZZ(J)) < 2 Then ZZpE = True
'''
''''If XX(ZZpF(I)) > XX(ZZpT(J)) Then ZZpE = True: Stop
''''If YY(ZZpF(I)) > YY(ZZpT(J)) Then ZZpE = True: Stop
''''If ZZ(ZZpF(I)) > ZZ(ZZpT(J)) Then ZZpE = True: Stop
'''
'''        End If
'''    Next
'''
'''End Function
Friend Sub ResetPoints()

    Dim X         As Long
    Dim Y         As Long
    Dim Z         As Long
    Dim I&
    For Z = MinCellZ To MAXCellZ
        For Y = MinCellY To MAXCellY
            For X = MinCellX To MAXCellX
                With CELL(X, Y, Z)
                    .NP = 0&
                    .myIDX = I: I = I + 1&
                End With
            Next
        Next
    Next
    NP = 0

    '......... Some memory saving
    If (CNT And 255&) = 0& Then
        For Z = MinCellZ To MAXCellZ
            For Y = MinCellY To MAXCellY
                For X = MinCellX To MAXCellX
                    With CELL(X, Y, Z)
                        .MaxNP = .MaxNP - 1
                        If .MaxNP < -10 Then
                            MaxNP = 0
                            ReDim px(MaxNP)
                            ReDim py(MaxNP)
                            ReDim pz(MaxNP)
                        End If
                    End With
                Next
            Next
        Next
    End If

End Sub

Friend Sub InsertPoint(ByVal X As Single, ByVal Y As Single, ByVal Z As Single)
    NP = NP + 1
    If NP > MaxNP Then
        MaxNP = 1 + NP * 1.25
        ReDim Preserve px(MaxNP)
        ReDim Preserve py(MaxNP)
        ReDim Preserve pz(MaxNP)
    End If

    px(NP) = X
    py(NP) = Y
    pz(NP) = Z

    pvAddToCell X \ GridSize, Y \ GridSize, Z \ GridSize, NP

End Sub

Friend Sub InsertALLpoints(X() As Single, Y() As Single, Z() As Single)
    'Remember to Call ResetPoints First

    Dim I         As Long
    Dim U         As Long
    U = UBound(X)

    For I = 1 To U
        NP = NP + 1&
        If NP > MaxNP Then
            MaxNP = NP + 64        '* 2
            '            MaxNP = 1 + NP * 1.5

            '            ReDim Preserve PX(MaxNP)
            '            ReDim Preserve PY(MaxNP)
        End If
        '        PX(NP) = X(I)
        '        PY(NP) = Y(I)
        pvAddToCell X(I) \ GridSize, Y(I) \ GridSize, Z(I) \ GridSize, NP
    Next
    pvArrCopySingle px, X
    pvArrCopySingle py, Y
    pvArrCopySingle pz, Z

End Sub

Private Sub pvAddToCell(ByVal cellX As Long, ByVal cellY As Long, ByVal CellZ As Long, ByVal CurrP As Long)

    With CELL(cellX, cellY, CellZ)
        .NP = .NP + 1&
        If .NP > .MaxNP Then
            .MaxNP = 1 + .NP * 1.25
            ReDim Preserve .pIDX(.MaxNP)
        End If
        .pIDX(.NP) = CurrP
    End With

End Sub



Friend Sub GetPairsWDist(RP1() As Long, RP2() As Long, _
                         Rdx() As Single, Rdy() As Single, Rdz() As Single, _
                         rD() As Single, rPairsCount As Long)

    'CORE SUB

    Dim X         As Long
    Dim Y         As Long
    Dim Z         As Long

    Dim DX        As Single
    Dim DY        As Single
    Dim DZ        As Single

    Dim I         As Long
    Dim J         As Long
    Dim D         As Single


    Dim Xp1       As Long
    Dim Xm1       As Long

    Dim Yp1       As Long
    Dim Ym1       As Long

    Dim Zp1       As Long
    Dim Zm1       As Long

    Dim iJ        As Long




    Npairs = 0&

    For Z = MinCellZ To MAXCellZ

        Zp1 = Z + 1&
        Zm1 = Z - 1&


        For Y = MinCellY To MAXCellY

            Yp1 = Y + 1&
            Ym1 = Y - 1
            For X = MinCellX To MAXCellX
                Xp1 = X + 1&
                Xm1 = X - 1&

                With CELL(X, Y, Z)

                    For I = 1& To .NP    '- 1& ' Should be -1 to do only SELF but so we can do even others

                        iI = .pIDX(I)

                        PXI = px(iI)
                        PYI = py(iI)
                        PZI = pz(iI)


                        For J = I + 1& To .NP    'SELF
                            iJ = .pIDX(J)
                            DX = px(iJ) - PXI
                            DY = py(iJ) - PYI
                            DZ = pz(iJ) - PZI
                            D = DX * DX + DY * DY + DZ * DZ
                            If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                        Next
                        '--------------------------------------------
                        '--------------------------------------------
                        '--------------------------------------------
                        '--------------------------------------------
                        '--------------------------------------------

                        '-------------------------------- Along XY

                        '''                        If Xp1 <= MAXCellX Then
                        '''                            For J = 1& To CELL(Xp1, y, Z).NP    ' X..
                        '''                                iJ = CELL(Xp1, y, Z).pIDX(J)
                        '''                                dX = PX(iJ) - PXI
                        '''                                dY = PY(iJ) - PYI
                        '''                                dZ = PZ(iJ) - PZI
                        '''                                D = dX * dX + dY * dY + dZ * dZ
                        '''                                If D < MaxDistance2 Then pvAddPairWDist iI, iJ, dX, dY, dZ, D
                        '''                            Next
                        '''                            If Yp1 <= MAXCellY Then
                        '''                                For J = 1& To CELL(Xp1, Yp1, Z).NP    'XY.
                        '''                                    iJ = CELL(Xp1, Yp1, Z).pIDX(J)
                        '''                                    dX = PX(iJ) - PXI
                        '''                                    dY = PY(iJ) - PYI
                        '''                                    dZ = PZ(iJ) - PZI
                        '''                                    D = dX * dX + dY * dY + dZ * dZ
                        '''                                    If D < MaxDistance2 Then pvAddPairWDist iI, iJ, dX, dY, dZ, D
                        '''                                Next
                        '''                            End If
                        '''                        End If
                        '''                        If Yp1 <= MAXCellY Then
                        '''                            For J = 1& To CELL(X, Yp1, Z).NP    '.Y.
                        '''                                iJ = CELL(X, Yp1, Z).pIDX(J)
                        '''                                dX = PX(iJ) - PXI
                        '''                                dY = PY(iJ) - PYI
                        '''                                dZ = PZ(iJ) - PZI
                        '''                                D = dX * dX + dY * dY + dZ * dZ
                        '''                                If D < MaxDistance2 Then pvAddPairWDist iI, iJ, dX, dY, dZ, D
                        '''                            Next
                        '''                        End If
                        '''                        '-------------------------------- Along XY at Z+1
                        '''                        If Zp1 <= MAXCellZ Then
                        '''                            If Xp1 <= MAXCellX Then
                        '''                                For J = 1& To CELL(Xp1, y, Zp1).NP    ' X.Z
                        '''                                    iJ = CELL(Xp1, y, Zp1).pIDX(J)
                        '''                                    dX = PX(iJ) - PXI
                        '''                                    dY = PY(iJ) - PYI
                        '''                                    dZ = PZ(iJ) - PZI
                        '''                                    D = dX * dX + dY * dY + dZ * dZ
                        '''                                    If D < MaxDistance2 Then pvAddPairWDist iI, iJ, dX, dY, dZ, D
                        '''                                Next
                        '''                                If Yp1 <= MAXCellY Then
                        '''                                    For J = 1& To CELL(Xp1, Yp1, Zp1).NP    'XYZ
                        '''                                        iJ = CELL(Xp1, Yp1, Zp1).pIDX(J)
                        '''                                        dX = PX(iJ) - PXI
                        '''                                        dY = PY(iJ) - PYI
                        '''                                        dZ = PZ(iJ) - PZI
                        '''                                        D = dX * dX + dY * dY + dZ * dZ
                        '''                                        If D < MaxDistance2 Then pvAddPairWDist iI, iJ, dX, dY, dZ, D
                        '''                                    Next
                        '''                                End If
                        '''                            End If
                        '''                            If Yp1 <= MAXCellY Then
                        '''                                For J = 1& To CELL(X, Yp1, Zp1).NP    '.YZ
                        '''                                    iJ = CELL(X, Yp1, Zp1).pIDX(J)
                        '''                                    dX = PX(iJ) - PXI
                        '''                                    dY = PY(iJ) - PYI
                        '''                                    dZ = PZ(iJ) - PZI
                        '''                                    D = dX * dX + dY * dY + dZ * dZ
                        '''                                    If D < MaxDistance2 Then pvAddPairWDist iI, iJ, dX, dY, dZ, D
                        '''                                Next
                        '''                            End If
                        '''
                        '''                            For J = 1& To CELL(X, y, Zp1).NP    '..Z
                        '''                                iJ = CELL(X, y, Zp1).pIDX(J)
                        '''                                dX = PX(iJ) - PXI
                        '''                                dY = PY(iJ) - PYI
                        '''                                dZ = PZ(iJ) - PZI
                        '''                                D = dX * dX + dY * dY + dZ * dZ
                        '''                                If D < MaxDistance2 Then pvAddPairWDist iI, iJ, dX, dY, dZ, D
                        '''                            Next
                        '''                        End If
                        '''
                        '''                        '-----------------------UP--------- Along XZ at y-1
                        '''                        If Ym1 >= 0& Then
                        '''                            If Xp1 <= MAXCellX Then
                        '''                                For J = 1& To CELL(Xp1, Ym1, Z).NP    'X -Y .
                        '''                                    iJ = CELL(Xp1, Ym1, Z).pIDX(J)
                        '''                                    dX = PX(iJ) - PXI
                        '''                                    dY = PY(iJ) - PYI
                        '''                                    dZ = PZ(iJ) - PZI
                        '''                                    D = dX * dX + dY * dY + dZ * dZ
                        '''                                    If D < MaxDistance2 Then pvAddPairWDist iI, iJ, dX, dY, dZ, D
                        '''                                Next
                        '''                                If Zp1 <= MAXCellZ Then
                        '''                                    For J = 1& To CELL(Xp1, Ym1, Zp1).NP    'X -Y Z
                        '''                                        iJ = CELL(Xp1, Ym1, Zp1).pIDX(J)
                        '''                                        dX = PX(iJ) - PXI
                        '''                                        dY = PY(iJ) - PYI
                        '''                                        dZ = PZ(iJ) - PZI
                        '''                                        D = dX * dX + dY * dY + dZ * dZ
                        '''                                        If D < MaxDistance2 Then pvAddPairWDist iI, iJ, dX, dY, dZ, D
                        '''                                    Next
                        '''                                End If
                        '''                            End If
                        '''                            If Zp1 <= MAXCellZ Then
                        '''                                For J = 1& To CELL(X, Ym1, Zp1).NP    '. -Y Z
                        '''                                    iJ = CELL(X, Ym1, Zp1).pIDX(J)
                        '''                                    dX = PX(iJ) - PXI
                        '''                                    dY = PY(iJ) - PYI
                        '''                                    dZ = PZ(iJ) - PZI
                        '''                                    D = dX * dX + dY * dY + dZ * dZ
                        '''                                    If D < MaxDistance2 Then pvAddPairWDist iI, iJ, dX, dY, dZ, D
                        '''                                Next
                        '''                            End If
                        '''
                        '''                        End If
                        '''
                        '''
                        '''
                        '''                        '----------------------------------------------
                        '''                        '----------------------------------------------
                        '''
                        '''                        If Xp1 <= MAXCellX Then
                        '''                            If Zm1 >= 0& Then
                        '''                                For J = 1& To CELL(Xp1, y, Zm1).NP  '  X . -Z
                        '''                                    iJ = CELL(Xp1, y, Zm1).pIDX(J)
                        '''                                    dX = PX(iJ) - PXI
                        '''                                    dY = PY(iJ) - PYI
                        '''                                    dZ = PZ(iJ) - PZI
                        '''                                    D = dX * dX + dY * dY + dZ * dZ
                        '''                                    If D < MaxDistance2 Then pvAddPairWDist iI, iJ, dX, dY, dZ, D
                        '''                                Next
                        '''                                If Yp1 <= MAXCellY Then
                        '''                                    For J = 1& To CELL(Xp1, Yp1, Zm1).NP  '  X Y -Z
                        '''                                        iJ = CELL(Xp1, Yp1, Zm1).pIDX(J)
                        '''                                        dX = PX(iJ) - PXI
                        '''                                        dY = PY(iJ) - PYI
                        '''                                        dZ = PZ(iJ) - PZI
                        '''                                        D = dX * dX + dY * dY + dZ * dZ
                        '''                                        If D < MaxDistance2 Then pvAddPairWDist iI, iJ, dX, dY, dZ, D
                        '''                                    Next
                        '''                                End If
                        '''                                If Ym1 >= 0& Then
                        '''                                    For J = 1& To CELL(Xp1, Ym1, Zm1).NP  '  X -Y -Z
                        '''                                        iJ = CELL(Xp1, Ym1, Zm1).pIDX(J)
                        '''                                        dX = PX(iJ) - PXI
                        '''                                        dY = PY(iJ) - PYI
                        '''                                        dZ = PZ(iJ) - PZI
                        '''                                        D = dX * dX + dY * dY + dZ * dZ
                        '''                                        If D < MaxDistance2 Then pvAddPairWDist iI, iJ, dX, dY, dZ, D
                        '''                                    Next
                        '''                                End If
                        '''                            End If
                        '''
                        '''
                        '''                        End If



                        ''''   ' --------- SECONDA PROVA

                        If Xp1 <= MAXCellX Then
                            With CELL(Xp1, Y, Z)
                                For J = 1& To .NP    ' X . .
                                    'iJ = CELL(Xp1, y, z).pIDX(J)
                                    iJ = .pIDX(J)

                                    DX = px(iJ) - PXI
                                    DY = py(iJ) - PYI
                                    DZ = pz(iJ) - PZI
                                    D = DX * DX + DY * DY + DZ * DZ
                                    If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                Next
                            End With
                            If Ym1 >= MinCellY Then
                                With CELL(Xp1, Ym1, Z)
                                    For J = 1& To .NP    ' X -Y .
                                        'iJ = CELL(Xp1, Ym1, z).pIDX(J)
                                        iJ = .pIDX(J)

                                        DX = px(iJ) - PXI
                                        DY = py(iJ) - PYI
                                        DZ = pz(iJ) - PZI
                                        D = DX * DX + DY * DY + DZ * DZ
                                        If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                    Next
                                End With
                            End If
                            If Yp1 <= MAXCellY Then
                                With CELL(Xp1, Yp1, Z)
                                    For J = 1& To .NP    ' X Y .
                                        'iJ = CELL(Xp1, Yp1, z).pIDX(J)
                                        iJ = .pIDX(J)

                                        DX = px(iJ) - PXI
                                        DY = py(iJ) - PYI
                                        DZ = pz(iJ) - PZI
                                        D = DX * DX + DY * DY + DZ * DZ
                                        If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                    Next
                                End With
                            End If
                        End If


                        If Yp1 <= MAXCellY Then
                            With CELL(X, Yp1, Z)
                                For J = 1& To .NP    ' . Y .
                                    'iJ = CELL(x, Yp1, z).pIDX(J)
                                    iJ = .pIDX(J)

                                    DX = px(iJ) - PXI
                                    DY = py(iJ) - PYI
                                    DZ = pz(iJ) - PZI
                                    D = DX * DX + DY * DY + DZ * DZ
                                    If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                Next
                            End With
                        End If
                        '------------------------------------

                        If Zp1 <= MAXCellZ Then
                            If Xp1 <= MAXCellX Then    '                  X + 1
                                With CELL(Xp1, Y, Zp1)
                                    For J = 1& To .NP    ' X . Z
                                        'iJ = CELL(Xp1, y, Zp1).pIDX(J)
                                        iJ = .pIDX(J)

                                        DX = px(iJ) - PXI
                                        DY = py(iJ) - PYI
                                        DZ = pz(iJ) - PZI
                                        D = DX * DX + DY * DY + DZ * DZ
                                        If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                    Next
                                End With
                                If Yp1 <= MAXCellY Then
                                    With CELL(Xp1, Yp1, Zp1)
                                        For J = 1& To .NP    ' X Y Z
                                            'iJ = CELL(Xp1, Yp1, Zp1).pIDX(J)
                                            iJ = .pIDX(J)

                                            DX = px(iJ) - PXI
                                            DY = py(iJ) - PYI
                                            DZ = pz(iJ) - PZI
                                            D = DX * DX + DY * DY + DZ * DZ
                                            If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                        Next
                                    End With
                                End If
                                If Ym1 >= MinCellY Then
                                    With CELL(Xp1, Ym1, Zp1)
                                        For J = 1& To .NP    ' X -Y Z
                                            'iJ = CELL(Xp1, Ym1, Zp1).pIDX(J)
                                            iJ = .pIDX(J)
                                            DX = px(iJ) - PXI
                                            DY = py(iJ) - PYI
                                            DZ = pz(iJ) - PZI
                                            D = DX * DX + DY * DY + DZ * DZ
                                            If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                        Next
                                    End With
                                End If
                            End If


                            '                                                              X + 0
                            With CELL(X, Y, Zp1)
                                For J = 1& To .NP    ' . . Z
                                    'iJ = CELL(x, y, Zp1).pIDX(J)
                                    iJ = .pIDX(J)

                                    DX = px(iJ) - PXI
                                    DY = py(iJ) - PYI
                                    DZ = pz(iJ) - PZI
                                    D = DX * DX + DY * DY + DZ * DZ
                                    If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                Next
                            End With
                            If Yp1 <= MAXCellY Then
                                With CELL(X, Yp1, Zp1)
                                    For J = 1& To .NP    ' . Y Z
                                        'iJ = CELL(x, Yp1, Zp1).pIDX(J)
                                        iJ = .pIDX(J)

                                        DX = px(iJ) - PXI
                                        DY = py(iJ) - PYI
                                        DZ = pz(iJ) - PZI
                                        D = DX * DX + DY * DY + DZ * DZ
                                        If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                    Next
                                End With
                            End If
                            If Ym1 >= MinCellY Then
                                With CELL(X, Ym1, Zp1)
                                    For J = 1& To .NP    ' . -Y Z
                                        'iJ = CELL(x, Ym1, Zp1).pIDX(J)
                                        iJ = .pIDX(J)

                                        DX = px(iJ) - PXI
                                        DY = py(iJ) - PYI
                                        DZ = pz(iJ) - PZI
                                        D = DX * DX + DY * DY + DZ * DZ
                                        If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                    Next
                                End With
                            End If




                            If Xm1 >= MinCellX Then    'MAXCellX Then  2022 BUG FIX  '                  X - 1
                                With CELL(Xm1, Y, Zp1)
                                    For J = 1& To .NP    ' -X . Z
                                        'iJ = CELL(Xm1, y, Zp1).pIDX(J)
                                        iJ = .pIDX(J)

                                        DX = px(iJ) - PXI
                                        DY = py(iJ) - PYI
                                        DZ = pz(iJ) - PZI
                                        D = DX * DX + DY * DY + DZ * DZ
                                        If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                    Next
                                End With
                                If Yp1 <= MAXCellY Then
                                    With CELL(Xm1, Yp1, Zp1)
                                        For J = 1& To .NP    ' -X Y Z
                                            iJ = .pIDX(J)
                                            DX = px(iJ) - PXI
                                            DY = py(iJ) - PYI
                                            DZ = pz(iJ) - PZI
                                            D = DX * DX + DY * DY + DZ * DZ
                                            If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                        Next
                                    End With
                                End If
                                If Ym1 >= MinCellY Then
                                    With CELL(Xm1, Ym1, Zp1)
                                        For J = 1& To .NP    ' -X -Y Z
                                            ' iJ = CELL(Xm1, Ym1, Zp1).pIDX(J)
                                            iJ = .pIDX(J)
                                            DX = px(iJ) - PXI
                                            DY = py(iJ) - PYI
                                            DZ = pz(iJ) - PZI
                                            D = DX * DX + DY * DY + DZ * DZ
                                            If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                        Next
                                    End With
                                End If
                            End If

                        End If

                    Next I

                End With

            Next X
        Next Y
    Next Z



    '--------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------

    rPairsCount = Npairs


    pvArrCopyLong RP1, PairP1
    pvArrCopyLong RP2, PairP2
    pvArrCopySingle Rdx, PairDX
    pvArrCopySingle Rdy, PairDY
    pvArrCopySingle Rdz, PairDZ

    pvArrCopySingle rD, PairD

    '----------------------------------------
    ' Redim to low value if for many times actual Npairs is less than upper bound Pairs Array
    ' Honestly Don't know if somehow could improve perforamnces
    If Npairs < UBound(PairP1) Then
        CountRN = CountRN + 1
        If CountRN > 200 Then
            MAXNpair = Npairs + 1
            ReDim Preserve PairP1(MAXNpair)
            ReDim Preserve PairP2(MAXNpair)
            ReDim Preserve PairDX(MAXNpair)
            ReDim Preserve PairDY(MAXNpair)
            ReDim Preserve PairDZ(MAXNpair)

            ReDim Preserve PairD(MAXNpair)
            CountRN = 0
        End If
    Else
        CountRN = 0
    End If


    ''Check duplicates
    '            For X = 1 To CountRN - 1
    '                For Y = X + 1 To CountRN
    '                    If (PairP1(X) = PairP1(Y)) And (PairP2(X) = PairP2(Y)) Then Stop
    '                    If (PairP1(X) = PairP2(Y)) And (PairP2(X) = PairP1(Y)) Then Stop
    '                Next
    '            Next


End Sub


Private Sub pvAddPairWDist(ByVal P1 As Long, ByVal P2 As Long, _
                           ByVal DX As Single, ByVal DY As Single, ByVal DZ As Single, _
                           ByVal D As Single)
    Npairs = Npairs + 1&

    If Npairs > MAXNpair Then

        MAXNpair = (Npairs + 2) * 1.25

        ReDim Preserve PairP1(MAXNpair)
        ReDim Preserve PairP2(MAXNpair)
        ReDim Preserve PairDX(MAXNpair)
        ReDim Preserve PairDY(MAXNpair)
        ReDim Preserve PairDZ(MAXNpair)
        ReDim Preserve PairD(MAXNpair)
    End If

    PairP1(Npairs) = P1
    PairP2(Npairs) = P2

    PairDX(Npairs) = DX
    PairDY(Npairs) = DY
    PairDZ(Npairs) = DZ

    PairD(Npairs) = D

End Sub




Friend Sub GetPairsWDist2(RP1() As Long, RP2() As Long, _
                          Rdx() As Single, Rdy() As Single, Rdz() As Single, _
                          rD() As Single, rPairsCount As Long)

    'prova ----- non usare....

    Dim X         As Long
    Dim Y         As Long
    Dim Z         As Long

    Dim DX        As Single
    Dim DY        As Single
    Dim DZ        As Single

    Dim I         As Long
    Dim J         As Long
    Dim D         As Single

    Dim cWm1      As Long
    Dim cHm1      As Long
    Dim cZm1      As Long

    Dim Xp1       As Long
    Dim Xm1       As Long

    Dim Yp1       As Long
    Dim Ym1       As Long

    Dim Zp1       As Long
    Dim Zm1       As Long

    Dim iJ        As Long
    Dim iI        As Long

    Dim PXI       As Single
    Dim PYI       As Single
    Dim PZI       As Single

    cWm1 = MAXCellX                '- 1&
    cHm1 = MAXCellY                '- 1&
    cZm1 = MAXCellZ                '- 1&

    Npairs = 0&

    Dim X2        As Long
    Dim Y2        As Long
    Dim z2        As Long

    For X = 0& To MAXCellX
        Xp1 = X + 1&
        Xm1 = X - 1&
        If Xp1 > MAXCellX Then Xp1 = MAXCellX
        If Xm1 < 0& Then Xm1 = 0&


        For Y = 0& To MAXCellY
            Yp1 = Y + 1&
            Ym1 = Y - 1
            If Yp1 > MAXCellY Then Yp1 = MAXCellY
            If Ym1 < 0& Then Ym1 = 0&

            For Z = 0& To MAXCellZ
                Zp1 = Z + 1&
                Zm1 = Z - 1&
                If Zp1 > MAXCellZ Then Zp1 = MAXCellZ
                If Zm1 < 0& Then Zm1 = 0&

                With CELL(X, Y, Z)

                    For I = 1& To .NP    '- 1& ' Should be -1 to do only SELF but so we can do even others

                        iI = .pIDX(I)

                        PXI = px(iI)
                        PYI = py(iI)
                        PZI = pz(iI)


                        For X2 = X To Xp1
                            For Y2 = Y To Yp1
                                For z2 = Z To Zp1
                                    For J = 1& To CELL(X2, Y2, z2).NP    ' X . .
                                        iJ = CELL(X2, Y2, z2).pIDX(J)
                                        If iJ - iI Then
                                            DX = px(iJ) - PXI
                                            DY = py(iJ) - PYI
                                            DZ = pz(iJ) - PZI
                                            D = DX * DX + DY * DY + DZ * DZ
                                            If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                        End If
                                    Next
                                Next
                            Next
                        Next
                    Next I

                End With

                If Xp1 - X Then

                    With CELL(Xp1, Y, Z)

                        For I = 1& To .NP    '- 1& ' Should be -1 to do only SELF but so we can do even others

                            iI = .pIDX(I)

                            PXI = px(iI)
                            PYI = py(iI)
                            PZI = pz(iI)

                            X2 = X
                            Y2 = Yp1
                            z2 = Z
                            If Y - Yp1 Then
                                For J = 1& To CELL(X2, Y2, z2).NP    ' X . .
                                    iJ = CELL(X2, Y2, z2).pIDX(J)
                                    If iJ - iI Then
                                        DX = px(iJ) - PXI
                                        DY = py(iJ) - PYI
                                        DZ = pz(iJ) - PZI
                                        D = DX * DX + DY * DY + DZ * DZ
                                        If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                    End If

                                Next
                            End If
                            X2 = X
                            Y2 = Y
                            z2 = Zp1
                            If Z - Zp1 Then
                                For J = 1& To CELL(X2, Y2, z2).NP    ' Zp1
                                    iJ = CELL(X2, Y2, z2).pIDX(J)
                                    If iJ - iI Then
                                        DX = px(iJ) - PXI
                                        DY = py(iJ) - PYI
                                        DZ = pz(iJ) - PZI
                                        D = DX * DX + DY * DY + DZ * DZ
                                        If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                    End If

                                Next
                            End If
                        Next I
                    End With
                End If
                '---
                If Zp1 - Z Then

                    With CELL(X, Y, Zp1)

                        For I = 1& To .NP    '- 1& ' Should be -1 to do only SELF but so we can do even others

                            iI = .pIDX(I)

                            PXI = px(iI)
                            PYI = py(iI)
                            PZI = pz(iI)


                            X2 = X
                            Y2 = Yp1
                            z2 = Z
                            If Y - Yp1 Then
                                For J = 1& To CELL(X2, Y2, z2).NP    ' X . .
                                    iJ = CELL(X2, Y2, z2).pIDX(J)
                                    If iJ - iI Then
                                        DX = px(iJ) - PXI
                                        DY = py(iJ) - PYI
                                        DZ = pz(iJ) - PZI
                                        D = DX * DX + DY * DY + DZ * DZ
                                        If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                    End If

                                Next
                            End If
                        Next I
                    End With
                End If

            Next Z
        Next Y
    Next X



    '--------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------

    rPairsCount = Npairs


    pvArrCopyLong RP1, PairP1
    pvArrCopyLong RP2, PairP2
    pvArrCopySingle Rdx, PairDX
    pvArrCopySingle Rdy, PairDY
    pvArrCopySingle Rdz, PairDZ

    pvArrCopySingle rD, PairD

    '----------------------------------------
    ' Redim to low value if for many times actual Npairs is less than upper bound Pairs Array
    ' Honestly Don't know if somehow could improve perforamnces
    If Npairs < UBound(PairP1) Then
        CountRN = CountRN + 1
        If CountRN > 200 Then
            MAXNpair = Npairs + 1
            ReDim Preserve PairP1(MAXNpair)
            ReDim Preserve PairP2(MAXNpair)
            ReDim Preserve PairDX(MAXNpair)
            ReDim Preserve PairDY(MAXNpair)
            ReDim Preserve PairDZ(MAXNpair)

            ReDim Preserve PairD(MAXNpair)
            CountRN = 0
        End If
    Else
        CountRN = 0
    End If


    ''Check duplicates
    ''        For X = 1 To CountRN - 1
    ''            For y = X + 1 To CountRN
    ''                If (PairP1(X) = PairP1(y)) And (PairP2(X) = PairP2(y)) Then Stop
    ''                If (PairP1(X) = PairP2(y)) And (PairP2(X) = PairP1(y)) Then Stop
    ''            Next
    ''        Next


End Sub








Friend Sub GetPairsWDist3(RP1() As Long, RP2() As Long, _
                          Rdx() As Single, Rdy() As Single, Rdz() As Single, _
                          rD() As Single, rPairsCount As Long)

    'CORE SUB

    Dim X         As Long
    Dim Y         As Long
    Dim Z         As Long

    Dim DX        As Single
    Dim DY        As Single
    Dim DZ        As Single

    Dim I         As Long
    Dim J         As Long
    Dim D         As Single


    Dim Xp1       As Long
    Dim Xm1       As Long

    Dim Yp1       As Long
    Dim Ym1       As Long

    Dim Zp1       As Long
    Dim Zm1       As Long

    Dim iJ        As Long


    Npairs = 0&

    For X = MinCellX To MAXCellX
        Xp1 = X + 1&
        Xm1 = X - 1&

        For Y = MinCellY To MAXCellY

            Yp1 = Y + 1&
            Ym1 = Y - 1

            For Z = MinCellZ To MAXCellZ

                Zp1 = Z + 1&
                Zm1 = Z - 1&

                With CELL(X, Y, Z)

                    For I = 1& To .NP    '- 1& ' Should be -1 to do only SELF but so we can do even others

                        iI = .pIDX(I)

                        PXI = px(iI)
                        PYI = py(iI)
                        PZI = pz(iI)

                        For J = I + 1& To .NP    'SELF
                            iJ = .pIDX(J)
                            DX = px(iJ) - PXI
                            DY = py(iJ) - PYI
                            DZ = pz(iJ) - PZI
                            D = DX * DX + DY * DY + DZ * DZ
                            If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                        Next
                        '--------------------------------------------
                        '--------------------------------------------
                        '--------------------------------------------

                        '-------------------------------- Along XY
                        If Xp1 <= MAXCellX Then
                            CHECKCELL Xp1, Y, Z
                            If Ym1 >= MinCellY Then CHECKCELL Xp1, Ym1, Z
                            If Yp1 <= MAXCellY Then CHECKCELL Xp1, Yp1, Z
                        End If
                        If Yp1 <= MAXCellY Then CHECKCELL X, Yp1, Z
                        '------------------------------------
                        If Zp1 <= MAXCellZ Then
                            If Xp1 <= MAXCellX Then
                                CHECKCELL Xp1, Y, Zp1
                                If Ym1 >= MinCellY Then CHECKCELL Xp1, Ym1, Zp1
                                If Yp1 <= MAXCellY Then CHECKCELL Xp1, Yp1, Zp1
                            End If
                            If Yp1 <= MAXCellY Then CHECKCELL X, Yp1, Zp1
                            CHECKCELL X, Y, Zp1    '''extra
                        End If
                        '-------------------------------------
                        If Zm1 >= MinCellZ Then
                            If Xp1 <= MAXCellX Then
                                CHECKCELL Xp1, Y, Zm1
                                If Ym1 >= MinCellY Then CHECKCELL Xp1, Ym1, Zm1
                                If Yp1 <= MAXCellY Then CHECKCELL Xp1, Yp1, Zm1
                            End If
                            If Yp1 <= MAXCellY Then CHECKCELL X, Yp1, Zm1
                        End If

                    Next I

                End With

            Next Z
        Next Y
    Next X



    '--------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------

    rPairsCount = Npairs


    pvArrCopyLong RP1, PairP1
    pvArrCopyLong RP2, PairP2
    pvArrCopySingle Rdx, PairDX
    pvArrCopySingle Rdy, PairDY
    pvArrCopySingle Rdz, PairDZ

    pvArrCopySingle rD, PairD

    '----------------------------------------
    ' Redim to low value if for many times actual Npairs is less than upper bound Pairs Array
    ' Honestly Don't know if somehow could improve perforamnces
    If Npairs < UBound(PairP1) Then
        CountRN = CountRN + 1
        If CountRN > 200 Then
            MAXNpair = Npairs + 1
            ReDim Preserve PairP1(MAXNpair)
            ReDim Preserve PairP2(MAXNpair)
            ReDim Preserve PairDX(MAXNpair)
            ReDim Preserve PairDY(MAXNpair)
            ReDim Preserve PairDZ(MAXNpair)

            ReDim Preserve PairD(MAXNpair)
            CountRN = 0
        End If
    Else
        CountRN = 0
    End If


    ''Check duplicates
    ''                For x = 1 To CountRN - 1
    ''                    For y = x + 1 To CountRN
    ''                        If (PairP1(x) = PairP1(y)) And (PairP2(x) = PairP2(y)) Then Stop
    ''                        If (PairP1(x) = PairP2(y)) And (PairP2(x) = PairP1(y)) Then Stop
    ''                    Next
    ''                Next


End Sub





Friend Sub GetPairsWDist4(RP1() As Long, RP2() As Long, _
                          Rdx() As Single, Rdy() As Single, Rdz() As Single, _
                          rD() As Single, rPairsCount As Long)

    'CORE SUB
    ' Without boundaries Check
    ' 2023

    Dim X         As Long
    Dim Y         As Long
    Dim Z         As Long

    Dim DX        As Single
    Dim DY        As Single
    Dim DZ        As Single

    Dim I         As Long
    Dim J         As Long
    Dim D         As Single


    Dim Xp1       As Long
    Dim Xm1       As Long

    Dim Yp1       As Long
    Dim Ym1       As Long

    Dim Zp1       As Long
    Dim Zm1       As Long

    Dim iJ        As Long

    Npairs = 0&

    Dim FX&, TX&
    Dim FY&, TY&
    Dim FZ&, TZ&

    FX = MinCellX + 1
    TX = MAXCellX - 1

    FY = MinCellY + 1
    TY = MAXCellY - 1

    FZ = MinCellZ + 1
    TZ = MAXCellZ - 1



    For Z = FZ To TZ

        Zp1 = Z + 1&
        Zm1 = Z - 1&

        For Y = FY To TY

            Yp1 = Y + 1&
            Yp1 = Y - 1&

            For X = FX To TX
                Xp1 = X + 1&
                Xm1 = X - 1&


                With CELL(X, Y, Z)

                    For I = 1& To .NP    '- 1& ' Should be -1 to do only SELF but so we can do even others

                        iI = .pIDX(I)

                        PXI = px(iI)
                        PYI = py(iI)
                        PZI = pz(iI)

                        For J = I + 1& To .NP    'SELF
                            iJ = .pIDX(J)
                            DX = px(iJ) - PXI
                            DY = py(iJ) - PYI
                            DZ = pz(iJ) - PZI
                            D = DX * DX + DY * DY + DZ * DZ
                            If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                        Next
                        '--------------------------------------------

                        '''  LIKE V3
                        CHECKCELL Xm1, Yp1, Zm1
                        CHECKCELL X, Yp1, Zm1
                        CHECKCELL Xp1, Y, Zm1
                        CHECKCELL Xp1, Yp1, Zm1

                        CHECKCELL Xm1, Yp1, Z
                        CHECKCELL X, Yp1, Z
                        CHECKCELL Xp1, Y, Z
                        CHECKCELL Xp1, Yp1, Z

                        CHECKCELL Xm1, Yp1, Zp1
                        CHECKCELL X, Yp1, Zp1
                        CHECKCELL Xp1, Y, Zp1
                        CHECKCELL Xp1, Yp1, Zp1

                        ' + Missing [ Added (was missing in V3) ]
                        CHECKCELL X, Y, Zp1

                        '


                    Next I

                End With

            Next X
        Next Y
    Next Z

    '--------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------

    rPairsCount = Npairs


    pvArrCopyLong RP1, PairP1
    pvArrCopyLong RP2, PairP2
    pvArrCopySingle Rdx, PairDX
    pvArrCopySingle Rdy, PairDY
    pvArrCopySingle Rdz, PairDZ

    pvArrCopySingle rD, PairD

    '----------------------------------------
    ' Redim to low value if for many times actual Npairs is less than upper bound Pairs Array
    ' Honestly Don't know if somehow could improve perforamnces
    If Npairs < 0.5 * UBound(PairP1) Then
        CountRN = CountRN + 1
        If CountRN > 200 Then
            MAXNpair = Npairs + 1
            ReDim Preserve PairP1(MAXNpair)
            ReDim Preserve PairP2(MAXNpair)
            ReDim Preserve PairDX(MAXNpair)
            ReDim Preserve PairDY(MAXNpair)
            ReDim Preserve PairDZ(MAXNpair)

            ReDim Preserve PairD(MAXNpair)
            CountRN = 0
        End If
    Else
        CountRN = 0
    End If


    ''Check duplicates
    '                    For X = 1 To CountRN - 1
    '                        For Y = X + 1 To CountRN
    '                            If (PairP1(X) = PairP1(Y)) And (PairP2(X) = PairP2(Y)) Then Stop
    '                            If (PairP1(X) = PairP2(Y)) And (PairP2(X) = PairP1(Y)) Then Stop
    '                        Next
    '                    Next


End Sub

Private Sub CHECKCELL(cX&, cY&, cZ&)
    Dim J&, iJ&, DX As Single, DY As Single, DZ As Single, D As Single
    With CELL(cX, cY, cZ)
        For J = 1& To .NP
            iJ = .pIDX(J)
            DX = px(iJ) - PXI
            DY = py(iJ) - PYI
            DZ = pz(iJ) - PZI
            D = DX * DX + DY * DY + DZ * DZ
            If D < MaxDistance2 Then    ' pvAddPairWDist  iI, iJ, DX, DY, DZ, D
                Npairs = Npairs + 1&
                If Npairs > MAXNpair Then
                    MAXNpair = (Npairs + 2) * 1.25
                    ReDim Preserve PairP1(MAXNpair)
                    ReDim Preserve PairP2(MAXNpair)
                    ReDim Preserve PairDX(MAXNpair)
                    ReDim Preserve PairDY(MAXNpair)
                    ReDim Preserve PairDZ(MAXNpair)
                    ReDim Preserve PairD(MAXNpair)
                End If
                PairP1(Npairs) = iI
                PairP2(Npairs) = iJ
                PairDX(Npairs) = DX
                PairDY(Npairs) = DY
                PairDZ(Npairs) = DZ
                PairD(Npairs) = D
            End If
        Next
    End With
End Sub


Private Sub TESTdot()
    Dim X As Single, Y As Single, Z As Single
    Dim sX$, sY$, sZ$

    Dim V         As tVec3
    Dim I&
    V = Normalize3(Vec3(1, 0.01, 0.02))

    For X = -1 To 1
        For Y = -1 To 1
            For Z = -1 To 1

                If DOT3(Vec3(X, Y, Y), V) > 0 Then
                    I = I + 1

                    sX = "X  "
                    sY = "Y  "
                    sZ = "Z  "

                    If X < 0 Then sX = "Xm1"
                    If X > 0 Then sX = "Xp1"

                    If Y < 0 Then sY = "Ym1"
                    If Y > 0 Then sY = "Yp1"

                    If Z < 0 Then sZ = "Zm1"
                    If Z > 0 Then sZ = "Zp1"


Debug.Print I & "      " & sX & ", " & sY & ", " & sZ & "      ", X, Y, Z
                End If
            Next
        Next
    Next
End Sub
