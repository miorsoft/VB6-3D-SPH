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
End Type


Private mW        As Long
Private mH        As Long
Private mZ        As Long

Private GridSize  As Long

Private CellW     As Long
Private CellH     As Long
Private cellZ     As Long


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
            If (W - UBound(dest, 1)) Then
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
            If (W - UBound(dest, 1)) Then
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



Public Sub Init(WorldW As Long, WorldH As Long, worldZ As Long, MaxDist As Long)
    mW = WorldW
    mH = WorldH
    mZ = worldZ

    GridSize = MaxDist
    MaxDistance2 = MaxDist * MaxDist

    CellW = mW \ GridSize
    CellH = mH \ GridSize
    cellZ = mZ \ GridSize

    ReDim CELL(CellW, CellH, cellZ)

    NP = 0
    Npairs = 0

End Sub

Friend Sub ResetPoints()

    Dim X         As Long
    Dim Y         As Long
    Dim Z         As Long

    For X = 0 To CellW
        For Y = 0 To CellH
            For Z = 0 To cellZ
                CELL(X, Y, Z).NP = 0&
            Next
        Next
    Next

    NP = 0

End Sub

Friend Sub InsertPoint(ByVal X As Double, ByVal Y As Double, ByVal Z As Double)
    NP = NP + 1
    If NP > MaxNP Then
        MaxNP = NP * 2
        ReDim Preserve pX(MaxNP)
        ReDim Preserve pY(MaxNP)
        ReDim Preserve pZ(MaxNP)

    End If
    pX(NP) = X
    pY(NP) = Y
    pZ(NP) = Z

    pvAddToCell X \ GridSize, Y \ GridSize, Z \ GridSize, NP

End Sub

Friend Sub InsertALLpoints(X() As Double, Y() As Double, Z() As Double)
'Remember to Call ResetPoints First

    Dim I         As Long
    Dim U         As Long
    U = UBound(X)

    For I = 1 To U
        NP = NP + 1
        If NP > MaxNP Then
            MaxNP = NP + 64    '* 2
            '            ReDim Preserve PX(MaxNP)
            '            ReDim Preserve PY(MaxNP)
        End If
        '        PX(NP) = X(I)
        '        PY(NP) = Y(I)
        pvAddToCell X(I) \ GridSize, Y(I) \ GridSize, Z(I) \ GridSize, NP
    Next
    pvArrCopySingle pX, X
    pvArrCopySingle pY, Y
    pvArrCopySingle pZ, Z

End Sub

Private Sub pvAddToCell(ByVal cellX As Long, ByVal cellY As Long, ByVal cellZ As Long, ByVal CurrP As Long)

    With CELL(cellX, cellY, cellZ)
        .NP = .NP + 1
        If .NP > .MaxNP Then
            .MaxNP = .NP + 64    '* 2
            ReDim Preserve .pIDX(.MaxNP)
        End If
        .pIDX(.NP) = CurrP
    End With

End Sub



Friend Sub GetPairsWDist(rP1() As Long, rP2() As Long, _
                         rDX() As Double, rDY() As Double, rDZ() As Double, _
                         rD() As Double, rPairsCount As Long)

'CORE SUB

    Dim X         As Long
    Dim Y         As Long
    Dim Z         As Long

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

    cWm1 = CellW - 1&
    cHm1 = CellH - 1&
    cZm1 = cellZ - 1&

    Npairs = 0&

    For X = 0& To CellW
        Xp1 = X + 1&
        Xm1 = X - 1&

        For Y = 0& To CellH
            Yp1 = Y + 1&
            Ym1 = Y - 1

            For Z = 0& To cellZ
                Zp1 = Z + 1&
                Zm1 = Z - 1&

                With CELL(X, Y, Z)

                    For I = 1& To .NP    '- 1& ' Should be -1 to do only SELF but so we can do even others

                        iI = .pIDX(I)

                        PXI = pX(iI)
                        PYI = pY(iI)
                        PZI = pZ(iI)


                        For J = I + 1& To .NP   'SELF
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

                        '''                        If Xp1 <= CellW Then
                        '''                            For J = 1& To CELL(Xp1, y, Z).NP    ' X..
                        '''                                iJ = CELL(Xp1, y, Z).pIDX(J)
                        '''                                dX = PX(iJ) - PXI
                        '''                                dY = PY(iJ) - PYI
                        '''                                dZ = PZ(iJ) - PZI
                        '''                                D = dX * dX + dY * dY + dZ * dZ
                        '''                                If D < MaxDistance2 Then pvAddPairWDist iI, iJ, dX, dY, dZ, D
                        '''                            Next
                        '''                            If Yp1 <= CellH Then
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
                        '''                        If Yp1 <= CellH Then
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
                        '''                        If Zp1 <= cellZ Then
                        '''                            If Xp1 <= CellW Then
                        '''                                For J = 1& To CELL(Xp1, y, Zp1).NP    ' X.Z
                        '''                                    iJ = CELL(Xp1, y, Zp1).pIDX(J)
                        '''                                    dX = PX(iJ) - PXI
                        '''                                    dY = PY(iJ) - PYI
                        '''                                    dZ = PZ(iJ) - PZI
                        '''                                    D = dX * dX + dY * dY + dZ * dZ
                        '''                                    If D < MaxDistance2 Then pvAddPairWDist iI, iJ, dX, dY, dZ, D
                        '''                                Next
                        '''                                If Yp1 <= CellH Then
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
                        '''                            If Yp1 <= CellH Then
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
                        '''                            If Xp1 <= CellW Then
                        '''                                For J = 1& To CELL(Xp1, Ym1, Z).NP    'X -Y .
                        '''                                    iJ = CELL(Xp1, Ym1, Z).pIDX(J)
                        '''                                    dX = PX(iJ) - PXI
                        '''                                    dY = PY(iJ) - PYI
                        '''                                    dZ = PZ(iJ) - PZI
                        '''                                    D = dX * dX + dY * dY + dZ * dZ
                        '''                                    If D < MaxDistance2 Then pvAddPairWDist iI, iJ, dX, dY, dZ, D
                        '''                                Next
                        '''                                If Zp1 <= cellZ Then
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
                        '''                            If Zp1 <= cellZ Then
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
                        '''                        If Xp1 <= CellW Then
                        '''                            If Zm1 >= 0& Then
                        '''                                For J = 1& To CELL(Xp1, y, Zm1).NP  '  X . -Z
                        '''                                    iJ = CELL(Xp1, y, Zm1).pIDX(J)
                        '''                                    dX = PX(iJ) - PXI
                        '''                                    dY = PY(iJ) - PYI
                        '''                                    dZ = PZ(iJ) - PZI
                        '''                                    D = dX * dX + dY * dY + dZ * dZ
                        '''                                    If D < MaxDistance2 Then pvAddPairWDist iI, iJ, dX, dY, dZ, D
                        '''                                Next
                        '''                                If Yp1 <= CellH Then
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

                        If Xp1 <= CellW Then
                            For J = 1& To CELL(Xp1, Y, Z).NP    ' X . .
                                iJ = CELL(Xp1, Y, Z).pIDX(J)
                                DX = pX(iJ) - PXI
                                DY = pY(iJ) - PYI
                                DZ = pZ(iJ) - PZI
                                D = DX * DX + DY * DY + DZ * DZ
                                If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                            Next
                            If Ym1 >= 0 Then
                                For J = 1& To CELL(Xp1, Ym1, Z).NP    ' X -Y .
                                    iJ = CELL(Xp1, Ym1, Z).pIDX(J)
                                    DX = pX(iJ) - PXI
                                    DY = pY(iJ) - PYI
                                    DZ = pZ(iJ) - PZI
                                    D = DX * DX + DY * DY + DZ * DZ
                                    If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                Next
                            End If
                            If Yp1 <= CellH Then
                                For J = 1& To CELL(Xp1, Yp1, Z).NP    ' X Y .
                                    iJ = CELL(Xp1, Yp1, Z).pIDX(J)
                                    DX = pX(iJ) - PXI
                                    DY = pY(iJ) - PYI
                                    DZ = pZ(iJ) - PZI
                                    D = DX * DX + DY * DY + DZ * DZ
                                    If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                Next
                            End If

                        End If
                        If Yp1 <= CellH Then
                            For J = 1& To CELL(X, Yp1, Z).NP    ' . Y .
                                iJ = CELL(X, Yp1, Z).pIDX(J)
                                DX = pX(iJ) - PXI
                                DY = pY(iJ) - PYI
                                DZ = pZ(iJ) - PZI
                                D = DX * DX + DY * DY + DZ * DZ
                                If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                            Next
                        End If
                        '------------------------------------

                        If Zp1 <= cellZ Then
                            If Xp1 <= CellW Then    '                  X + 1
                                For J = 1& To CELL(Xp1, Y, Zp1).NP    ' X . Z
                                    iJ = CELL(Xp1, Y, Zp1).pIDX(J)
                                    DX = pX(iJ) - PXI
                                    DY = pY(iJ) - PYI
                                    DZ = pZ(iJ) - PZI
                                    D = DX * DX + DY * DY + DZ * DZ
                                    If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                Next
                                If Yp1 <= CellH Then
                                    For J = 1& To CELL(Xp1, Yp1, Zp1).NP    ' X Y Z
                                        iJ = CELL(Xp1, Yp1, Zp1).pIDX(J)
                                        DX = pX(iJ) - PXI
                                        DY = pY(iJ) - PYI
                                        DZ = pZ(iJ) - PZI
                                        D = DX * DX + DY * DY + DZ * DZ
                                        If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                    Next
                                End If
                                If Ym1 >= 0 Then
                                    For J = 1& To CELL(Xp1, Ym1, Zp1).NP    ' X -Y Z
                                        iJ = CELL(Xp1, Ym1, Zp1).pIDX(J)
                                        DX = pX(iJ) - PXI
                                        DY = pY(iJ) - PYI
                                        DZ = pZ(iJ) - PZI
                                        D = DX * DX + DY * DY + DZ * DZ
                                        If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                    Next
                                End If
                            End If


                            '                                                              X + 0
                            For J = 1& To CELL(X, Y, Zp1).NP    ' . . Z
                                iJ = CELL(X, Y, Zp1).pIDX(J)
                                DX = pX(iJ) - PXI
                                DY = pY(iJ) - PYI
                                DZ = pZ(iJ) - PZI
                                D = DX * DX + DY * DY + DZ * DZ
                                If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                            Next
                            If Yp1 <= CellH Then
                                For J = 1& To CELL(X, Yp1, Zp1).NP    ' . Y Z
                                    iJ = CELL(X, Yp1, Zp1).pIDX(J)
                                    DX = pX(iJ) - PXI
                                    DY = pY(iJ) - PYI
                                    DZ = pZ(iJ) - PZI
                                    D = DX * DX + DY * DY + DZ * DZ
                                    If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                Next
                            End If
                            If Ym1 >= 0 Then
                                For J = 1& To CELL(X, Ym1, Zp1).NP    ' . -Y Z
                                    iJ = CELL(X, Ym1, Zp1).pIDX(J)
                                    DX = pX(iJ) - PXI
                                    DY = pY(iJ) - PYI
                                    DZ = pZ(iJ) - PZI
                                    D = DX * DX + DY * DY + DZ * DZ
                                    If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                Next
                            End If




                            If Xm1 >= CellW Then    '                  X - 1
                                For J = 1& To CELL(Xm1, Y, Zp1).NP    ' -X . Z
                                    iJ = CELL(Xm1, Y, Zp1).pIDX(J)
                                    DX = pX(iJ) - PXI
                                    DY = pY(iJ) - PYI
                                    DZ = pZ(iJ) - PZI
                                    D = DX * DX + DY * DY + DZ * DZ
                                    If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                Next
                                If Yp1 <= CellH Then
                                    For J = 1& To CELL(Xm1, Yp1, Zp1).NP    ' -X Y Z
                                        iJ = CELL(Xm1, Yp1, Zp1).pIDX(J)
                                        DX = pX(iJ) - PXI
                                        DY = pY(iJ) - PYI
                                        DZ = pZ(iJ) - PZI
                                        D = DX * DX + DY * DY + DZ * DZ
                                        If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                    Next
                                End If
                                If Ym1 >= 0 Then
                                    For J = 1& To CELL(Xm1, Ym1, Zp1).NP    ' -X -Y Z
                                        iJ = CELL(Xm1, Ym1, Zp1).pIDX(J)
                                        DX = pX(iJ) - PXI
                                        DY = pY(iJ) - PYI
                                        DZ = pZ(iJ) - PZI
                                        D = DX * DX + DY * DY + DZ * DZ
                                        If D < MaxDistance2 Then pvAddPairWDist iI, iJ, DX, DY, DZ, D
                                    Next
                                End If
                            End If

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


Private Sub pvAddPairWDist(ByVal P1 As Long, ByVal P2 As Long, _
                           ByVal DX As Double, ByVal DY As Double, ByVal DZ As Double, _
                           ByVal D As Double)
    Npairs = Npairs + 1

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
