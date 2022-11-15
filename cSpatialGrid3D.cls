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

Private MAXCellX     As Long
Private MAXCellY     As Long
Private MAXCellZ     As Long
Private MinCellX     As Long
Private MinCellY     As Long
Private MinCellZ     As Long


Private CELL()    As tCELL
Private NP        As Long
Private MaxNP     As Long
Private pX()      As Double
Private pY()      As Double
Private pZ()      As Double

Private MaxDistance2 As Double


Private PairP1()  As Long
Private PairP2()  As Long
Private PairDX()  As Double
Private PairDY()  As Double
Private PairDZ()  As Double

Private PairD()   As Double
Private Npairs    As Long
Private MAXNpair  As Long

Private CountRN   As Long



Private XX()      As Long
Private YY()      As Long
Private ZZ()      As Long
Private Ninc      As Long
Private ZZpF()    As Long
Private ZZpT()    As Long


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal ByteLen As Long)
Private Declare Function GetMem4 Lib "MSVBVM60" (ByVal Src As Long, ByVal dest As Long) As Long
Private Declare Function ArrPtr Lib "MSVBVM60" Alias "VarPtr" (arr() As Any) As Long




Private Sub pvArrCopySingle(dest() As Double, Src() As Double)
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

    MAXCellX = WorldW \ GridSize + 1 ' +1 because of pvX
    MAXCellY = WorldH \ GridSize + 1
    MAXCellZ = worldZ \ GridSize + 1
    
    MinCellX = -1 ' -1 because of pvX
    MinCellY = -1
    MinCellZ = -1
   
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

    Dim x         As Long
    Dim y         As Long
    Dim z         As Long
    Dim I&
    For x = MinCellX To MAXCellX
        For y = MinCellY To MAXCellY
            For z = MinCellZ To MAXCellZ
                With CELL(x, y, z)
                    .NP = 0&
                    .myIDX = I: I = I + 1&
                End With
            Next
        Next
    Next

    NP = 0




End Sub

Friend Sub InsertPoint(ByVal x As Double, ByVal y As Double, ByVal z As Double)
    NP = NP + 1
    If NP > MaxNP Then
        MaxNP = 1 + NP * 1.25
        ReDim Preserve pX(MaxNP)
        ReDim Preserve pY(MaxNP)
        ReDim Preserve pZ(MaxNP)

    End If
    pX(NP) = x
    pY(NP) = y
    pZ(NP) = z

    pvAddToCell x \ GridSize, y \ GridSize, z \ GridSize, NP




End Sub

Friend Sub InsertALLpoints(x() As Double, y() As Double, z() As Double)
'Remember to Call ResetPoints First

    Dim I         As Long
    Dim U         As Long
    U = UBound(x)

    For I = 1 To U
        NP = NP + 1&
        If NP > MaxNP Then
            MaxNP = NP + 64  '* 2
            '            MaxNP = 1 + NP * 1.5

            '            ReDim Preserve PX(MaxNP)
            '            ReDim Preserve PY(MaxNP)
        End If
        '        PX(NP) = X(I)
        '        PY(NP) = Y(I)
        pvAddToCell x(I) \ GridSize, y(I) \ GridSize, z(I) \ GridSize, NP
    Next
    pvArrCopySingle pX, x
    pvArrCopySingle pY, y
    pvArrCopySingle pZ, z

End Sub

Private Sub pvAddToCell(ByVal cellX As Long, ByVal cellY As Long, ByVal CellZ As Long, ByVal CurrP As Long)

    With CELL(cellX, cellY, CellZ)
        .NP = .NP + 1&
        If .NP > .MaxNP Then
            '            .MaxNP = .NP + 64    '* 2
            .MaxNP = 1 + .NP * 1.25
            ReDim Preserve .pIDX(.MaxNP)
        End If
        .pIDX(.NP) = CurrP
    End With

End Sub



Friend Sub GetPairsWDist(rP1() As Long, rP2() As Long, _
                         rDX() As Double, rDY() As Double, rDZ() As Double, _
                         rD() As Double, rPairsCount As Long)

'CORE SUB

    Dim x         As Long
    Dim y         As Long
    Dim z         As Long

    Dim DX        As Double
    Dim DY        As Double
    Dim DZ        As Double

    Dim I         As Long
    Dim J         As Long
    Dim D         As Double


    Dim Xp1       As Long
    Dim Xm1       As Long

    Dim Yp1       As Long
    Dim Ym1       As Long

    Dim Zp1       As Long
    Dim Zm1       As Long

    Dim iJ        As Long
    Dim iI        As Long

    Dim PXI       As Double
    Dim PYI       As Double
    Dim PZI       As Double




    Npairs = 0&

    For x = MinCellX To MAXCellX
        Xp1 = x + 1&
        Xm1 = x - 1&

        For y = MinCellY To MAXCellY

            Yp1 = y + 1&
            Ym1 = y - 1

            For z = MinCellZ To MAXCellZ

                Zp1 = z + 1&
                Zm1 = z - 1&

                With CELL(x, y, z)

                    For I = 1& To .NP    '- 1& ' Should be -1 to do only SELF but so we can do even others

                        iI = .pIDX(I)

                        PXI = pX(iI)
                        PYI = pY(iI)
                        PZI = pZ(iI)


                        For J = I + 1& To .NP    'SELF
                            iJ = .pIDX(J)
                            DX = pX(iJ) - PXI
                            DY = pY(iJ) - PYI
                            DZ = pZ(iJ) - PZI
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
                            With CELL(Xp1, y, z)
                                For J = 1& To .NP    ' X . .
                                    'iJ = CELL(Xp1, y, z).pIDX(J)
                                    iJ = .pIDX(J)

                                    DX = pX(iJ) - PXI
                                    DY = pY(iJ) - PYI
                                    DZ = pZ(iJ) - PZI
                                    D = DX * DX + DY * DY + DZ * DZ
                                    If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                Next
                            End With
                            If Ym1 >= MinCellY Then
                                With CELL(Xp1, Ym1, z)
                                    For J = 1& To .NP    ' X -Y .
                                        'iJ = CELL(Xp1, Ym1, z).pIDX(J)
                                        iJ = .pIDX(J)

                                        DX = pX(iJ) - PXI
                                        DY = pY(iJ) - PYI
                                        DZ = pZ(iJ) - PZI
                                        D = DX * DX + DY * DY + DZ * DZ
                                        If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                    Next
                                End With
                            End If
                            If Yp1 <= MAXCellY Then
                                With CELL(Xp1, Yp1, z)
                                    For J = 1& To .NP    ' X Y .
                                        'iJ = CELL(Xp1, Yp1, z).pIDX(J)
                                        iJ = .pIDX(J)

                                        DX = pX(iJ) - PXI
                                        DY = pY(iJ) - PYI
                                        DZ = pZ(iJ) - PZI
                                        D = DX * DX + DY * DY + DZ * DZ
                                        If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                    Next
                                End With
                            End If
                        End If


                        If Yp1 <= MAXCellY Then
                            With CELL(x, Yp1, z)
                                For J = 1& To .NP    ' . Y .
                                    'iJ = CELL(x, Yp1, z).pIDX(J)
                                    iJ = .pIDX(J)

                                    DX = pX(iJ) - PXI
                                    DY = pY(iJ) - PYI
                                    DZ = pZ(iJ) - PZI
                                    D = DX * DX + DY * DY + DZ * DZ
                                    If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                Next
                            End With
                        End If
                        '------------------------------------

                        If Zp1 <= MAXCellZ Then
                            If Xp1 <= MAXCellX Then    '                  X + 1
                                With CELL(Xp1, y, Zp1)
                                    For J = 1& To .NP    ' X . Z
                                        'iJ = CELL(Xp1, y, Zp1).pIDX(J)
                                        iJ = .pIDX(J)

                                        DX = pX(iJ) - PXI
                                        DY = pY(iJ) - PYI
                                        DZ = pZ(iJ) - PZI
                                        D = DX * DX + DY * DY + DZ * DZ
                                        If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                    Next
                                End With
                                If Yp1 <= MAXCellY Then
                                    With CELL(Xp1, Yp1, Zp1)
                                        For J = 1& To .NP    ' X Y Z
                                            'iJ = CELL(Xp1, Yp1, Zp1).pIDX(J)
                                            iJ = .pIDX(J)

                                            DX = pX(iJ) - PXI
                                            DY = pY(iJ) - PYI
                                            DZ = pZ(iJ) - PZI
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
                                            DX = pX(iJ) - PXI
                                            DY = pY(iJ) - PYI
                                            DZ = pZ(iJ) - PZI
                                            D = DX * DX + DY * DY + DZ * DZ
                                            If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                        Next
                                    End With
                                End If
                            End If


                            '                                                              X + 0
                            With CELL(x, y, Zp1)
                                For J = 1& To .NP    ' . . Z
                                    'iJ = CELL(x, y, Zp1).pIDX(J)
                                    iJ = .pIDX(J)

                                    DX = pX(iJ) - PXI
                                    DY = pY(iJ) - PYI
                                    DZ = pZ(iJ) - PZI
                                    D = DX * DX + DY * DY + DZ * DZ
                                    If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                Next
                            End With
                            If Yp1 <= MAXCellY Then
                                With CELL(x, Yp1, Zp1)
                                    For J = 1& To .NP    ' . Y Z
                                        'iJ = CELL(x, Yp1, Zp1).pIDX(J)
                                        iJ = .pIDX(J)

                                        DX = pX(iJ) - PXI
                                        DY = pY(iJ) - PYI
                                        DZ = pZ(iJ) - PZI
                                        D = DX * DX + DY * DY + DZ * DZ
                                        If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                    Next
                                End With
                            End If
                            If Ym1 >= MinCellY Then
                                With CELL(x, Ym1, Zp1)
                                    For J = 1& To .NP    ' . -Y Z
                                        'iJ = CELL(x, Ym1, Zp1).pIDX(J)
                                        iJ = .pIDX(J)

                                        DX = pX(iJ) - PXI
                                        DY = pY(iJ) - PYI
                                        DZ = pZ(iJ) - PZI
                                        D = DX * DX + DY * DY + DZ * DZ
                                        If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                    Next
                                End With
                            End If




                            If Xm1 >= MinCellX Then    'MAXCellX Then  2022 BUG FIX  '                  X - 1
                                With CELL(Xm1, y, Zp1)
                                    For J = 1& To .NP    ' -X . Z
                                        'iJ = CELL(Xm1, y, Zp1).pIDX(J)
                                        iJ = .pIDX(J)

                                        DX = pX(iJ) - PXI
                                        DY = pY(iJ) - PYI
                                        DZ = pZ(iJ) - PZI
                                        D = DX * DX + DY * DY + DZ * DZ
                                        If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                    Next
                                End With
                                If Yp1 <= MAXCellY Then
                                    With CELL(Xm1, Yp1, Zp1)
                                        For J = 1& To .NP    ' -X Y Z
                                            iJ = .pIDX(J)
                                            DX = pX(iJ) - PXI
                                            DY = pY(iJ) - PYI
                                            DZ = pZ(iJ) - PZI
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
                                            DX = pX(iJ) - PXI
                                            DY = pY(iJ) - PYI
                                            DZ = pZ(iJ) - PZI
                                            D = DX * DX + DY * DY + DZ * DZ
                                            If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                        Next
                                    End With
                                End If
                            End If

                        End If

                    Next I

                End With

            Next z
        Next y
    Next x



    '--------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------

    rPairsCount = Npairs


    pvArrCopyLong rP1, PairP1
    pvArrCopyLong rP2, PairP2
    pvArrCopySingle rDX, PairDX
    pvArrCopySingle rDY, PairDY
    pvArrCopySingle rDZ, PairDZ

    pvArrCopySingle rD, PairD

    '----------------------------------------
    ' Redim to low value if for many times actual Npairs is less than upper bound Pairs Array
    ' Honestly Don't know if somehow could improve perforamnces
    If Npairs < UBound(PairP1) Then
        CountRN = CountRN + 1
        If CountRN > 250 Then
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
                           ByVal DX As Double, ByVal DY As Double, ByVal DZ As Double, _
                           ByVal D As Double)
    Npairs = Npairs + 1&

    If Npairs > MAXNpair Then

        MAXNpair = (Npairs + 2) * 1.5

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




Friend Sub GetPairsWDist2(rP1() As Long, rP2() As Long, _
                          rDX() As Double, rDY() As Double, rDZ() As Double, _
                          rD() As Double, rPairsCount As Long)

'prova ----- non usare....

    Dim x         As Long
    Dim y         As Long
    Dim z         As Long

    Dim DX        As Double
    Dim DY        As Double
    Dim DZ        As Double

    Dim I         As Long
    Dim J         As Long
    Dim D         As Double

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

    Dim PXI       As Double
    Dim PYI       As Double
    Dim PZI       As Double

    cWm1 = MAXCellX             '- 1&
    cHm1 = MAXCellY             '- 1&
    cZm1 = MAXCellZ             '- 1&

    Npairs = 0&

    Dim X2        As Long
    Dim Y2        As Long
    Dim z2        As Long

    For x = 0& To MAXCellX
        Xp1 = x + 1&
        Xm1 = x - 1&
        If Xp1 > MAXCellX Then Xp1 = MAXCellX
        If Xm1 < 0& Then Xm1 = 0&


        For y = 0& To MAXCellY
            Yp1 = y + 1&
            Ym1 = y - 1
            If Yp1 > MAXCellY Then Yp1 = MAXCellY
            If Ym1 < 0& Then Ym1 = 0&

            For z = 0& To MAXCellZ
                Zp1 = z + 1&
                Zm1 = z - 1&
                If Zp1 > MAXCellZ Then Zp1 = MAXCellZ
                If Zm1 < 0& Then Zm1 = 0&

                With CELL(x, y, z)

                    For I = 1& To .NP    '- 1& ' Should be -1 to do only SELF but so we can do even others

                        iI = .pIDX(I)

                        PXI = pX(iI)
                        PYI = pY(iI)
                        PZI = pZ(iI)


                        For X2 = x To Xp1
                            For Y2 = y To Yp1
                                For z2 = z To Zp1
                                    For J = 1& To CELL(X2, Y2, z2).NP    ' X . .
                                        iJ = CELL(X2, Y2, z2).pIDX(J)
                                        If iJ - iI Then
                                            DX = pX(iJ) - PXI
                                            DY = pY(iJ) - PYI
                                            DZ = pZ(iJ) - PZI
                                            D = DX * DX + DY * DY + DZ * DZ
                                            If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                        End If
                                    Next
                                Next
                            Next
                        Next
                    Next I

                End With

                If Xp1 - x Then

                    With CELL(Xp1, y, z)

                        For I = 1& To .NP    '- 1& ' Should be -1 to do only SELF but so we can do even others

                            iI = .pIDX(I)

                            PXI = pX(iI)
                            PYI = pY(iI)
                            PZI = pZ(iI)

                            X2 = x
                            Y2 = Yp1
                            z2 = z
                            If y - Yp1 Then
                                For J = 1& To CELL(X2, Y2, z2).NP    ' X . .
                                    iJ = CELL(X2, Y2, z2).pIDX(J)
                                    If iJ - iI Then
                                        DX = pX(iJ) - PXI
                                        DY = pY(iJ) - PYI
                                        DZ = pZ(iJ) - PZI
                                        D = DX * DX + DY * DY + DZ * DZ
                                        If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                    End If

                                Next
                            End If
                            X2 = x
                            Y2 = y
                            z2 = Zp1
                            If z - Zp1 Then
                                For J = 1& To CELL(X2, Y2, z2).NP    ' Zp1
                                    iJ = CELL(X2, Y2, z2).pIDX(J)
                                    If iJ - iI Then
                                        DX = pX(iJ) - PXI
                                        DY = pY(iJ) - PYI
                                        DZ = pZ(iJ) - PZI
                                        D = DX * DX + DY * DY + DZ * DZ
                                        If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                    End If

                                Next
                            End If
                        Next I
                    End With
                End If
                '---
                If Zp1 - z Then

                    With CELL(x, y, Zp1)

                        For I = 1& To .NP    '- 1& ' Should be -1 to do only SELF but so we can do even others

                            iI = .pIDX(I)

                            PXI = pX(iI)
                            PYI = pY(iI)
                            PZI = pZ(iI)


                            X2 = x
                            Y2 = Yp1
                            z2 = z
                            If y - Yp1 Then
                                For J = 1& To CELL(X2, Y2, z2).NP    ' X . .
                                    iJ = CELL(X2, Y2, z2).pIDX(J)
                                    If iJ - iI Then
                                        DX = pX(iJ) - PXI
                                        DY = pY(iJ) - PYI
                                        DZ = pZ(iJ) - PZI
                                        D = DX * DX + DY * DY + DZ * DZ
                                        If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                    End If

                                Next
                            End If
                        Next I
                    End With
                End If

            Next z
        Next y
    Next x



    '--------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------

    rPairsCount = Npairs


    pvArrCopyLong rP1, PairP1
    pvArrCopyLong rP2, PairP2
    pvArrCopySingle rDX, PairDX
    pvArrCopySingle rDY, PairDY
    pvArrCopySingle rDZ, PairDZ

    pvArrCopySingle rD, PairD

    '----------------------------------------
    ' Redim to low value if for many times actual Npairs is less than upper bound Pairs Array
    ' Honestly Don't know if somehow could improve perforamnces
    If Npairs < UBound(PairP1) Then
        CountRN = CountRN + 1
        If CountRN > 250 Then
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

