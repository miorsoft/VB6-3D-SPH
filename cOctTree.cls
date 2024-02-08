VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOctTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public X0         As Single
Public X1         As Single
Public Y0         As Single
Public Y1         As Single
Public Z0         As Single
Attribute Z0.VB_VarUserMemId = 1073938434
Public Z1         As Single
Attribute Z1.VB_VarUserMemId = 1073938435

Private MX        As Single
Attribute MX.VB_VarUserMemId = 1073938435
Private MY        As Single
Attribute MY.VB_VarUserMemId = 1610809344
Private MZ        As Single
Attribute MZ.VB_VarUserMemId = 1073938432

'Private Q1        As cOctTree
'Private Q2        As cOctTree
'Private Q3        As cOctTree
'Private Q4        As cOctTree
'
'Private Q5        As cOctTree
'Private Q6        As cOctTree
'Private Q7        As cOctTree
'Private Q8        As cOctTree

Private QUA(1 To 8) As cOctTree

Private Subdivided As Long
Attribute Subdivided.VB_VarUserMemId = 1073938441

Public Capacity   As Long
Attribute Capacity.VB_VarUserMemId = 1073938443
Public NP         As Long
Attribute NP.VB_VarUserMemId = 1073938444
Private MaxNP     As Long
Attribute MaxNP.VB_VarUserMemId = 1610809347

Private px()      As Single
Attribute px.VB_VarUserMemId = 1073741837
Private py()      As Single
Attribute py.VB_VarUserMemId = 1073741838
Private pz()      As Single
Attribute pz.VB_VarUserMemId = 1610809348

Private pIDX()    As Single
Attribute pIDX.VB_VarUserMemId = 1073938448
'"Ghost Points" used in GetPairsWDist
Private GPX()     As Single
Attribute GPX.VB_VarUserMemId = 1610809349
Private GPY()     As Single
Attribute GPY.VB_VarUserMemId = 1073741841
Private GPZ()     As Single
Attribute GPZ.VB_VarUserMemId = 1073938451

Public GhostNP    As Long
Attribute GhostNP.VB_VarUserMemId = 1610809350
Private GhostmaxNP As Long
Attribute GhostmaxNP.VB_VarUserMemId = 1073938461

Public UnUsed     As Long
Attribute UnUsed.VB_VarUserMemId = 1073741846

Private H         As Single

Private Sub Class_Initialize()
    '    Subdivided = 0
End Sub


Friend Sub Setup(ByVal Xleft As Single, ByVal Ytop As Single, ByVal Zup As Single, _
                 ByVal Xright As Single, ByVal Ydown As Single, ByVal ZDown As Single, ByVal MyCapacity As Long)    ', DDD As single)


    X0 = Xleft
    X1 = Xright
    Y0 = Ytop
    Y1 = Ydown
    Z0 = Zup
    Z1 = ZDown


    Capacity = MyCapacity


    '    H = DDD

    Reset
End Sub
Friend Sub Reset()
    NP = 0
    GhostNP = 0
    If Subdivided Then
        '        If Q1.NP Then Q1.Reset
        '        If Q2.NP Then Q2.Reset
        '        If Q3.NP Then Q3.Reset
        '        If Q4.NP Then Q4.Reset
        '
        '        If Q5.NP Then Q5.Reset
        '        If Q6.NP Then Q6.Reset
        '        If Q7.NP Then Q7.Reset
        '        If Q8.NP Then Q8.Reset

        Dim I     As Long
        For I = 1 To 8
            If QUA(I).NP Then QUA(I).Reset
        Next

    End If
End Sub

Private Sub Subdivide()

    MX = (X0 + X1) * 0.5
    MY = (Y0 + Y1) * 0.5
    MZ = (Z0 + Z1) * 0.5

    Subdivided = -1

    Dim I         As Long

    '    If Q1 Is Nothing Then
    '        Set Q1 = New cOctTree
    '        Set Q2 = New cOctTree
    '        Set Q3 = New cOctTree
    '        Set Q4 = New cOctTree
    '
    '        Set Q5 = New cOctTree
    '        Set Q6 = New cOctTree
    '        Set Q7 = New cOctTree
    '        Set Q8 = New cOctTree
    '    End If
    If QUA(1) Is Nothing Then
        For I = 1 To 8
            Set QUA(I) = New cOctTree
        Next
    End If

    '    Q1.Setup X0, Y0, Z0, MX, MY, MZ, Capacity
    '    Q2.Setup MX, Y0, Z0, X1, MY, MZ, Capacity
    '    Q3.Setup X0, MY, Z0, MX, Y1, MZ, Capacity
    '    Q4.Setup MX, MY, Z0, X1, Y1, MZ, Capacity
    '
    '    Q5.Setup X0, Y0, MZ, MX, MY, Z1, Capacity
    '    Q6.Setup MX, Y0, MZ, X1, MY, Z1, Capacity
    '    Q7.Setup X0, MY, MZ, MX, Y1, Z1, Capacity
    '    Q8.Setup MX, MY, MZ, X1, Y1, Z1, Capacity

    QUA(1).Setup X0, Y0, Z0, MX, MY, MZ, Capacity    ', H
    QUA(2).Setup MX, Y0, Z0, X1, MY, MZ, Capacity    ', H
    QUA(3).Setup X0, MY, Z0, MX, Y1, MZ, Capacity    ', H
    QUA(4).Setup MX, MY, Z0, X1, Y1, MZ, Capacity    ', H

    QUA(5).Setup X0, Y0, MZ, MX, MY, Z1, Capacity    ', H
    QUA(6).Setup MX, Y0, MZ, X1, MY, Z1, Capacity    ', H
    QUA(7).Setup X0, MY, MZ, MX, Y1, Z1, Capacity    ', H
    QUA(8).Setup MX, MY, MZ, X1, Y1, Z1, Capacity    ', H

End Sub

Friend Sub InsertSinglePoint(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, ByVal Index As Long, Optional isRoot As Boolean = True)


    Dim Octant    As Long

    If isRoot Then                 ' FOR GetPairsWDist
        GhostNP = GhostNP + 1
        If GhostNP > GhostmaxNP Then
            GhostmaxNP = 1 + GhostNP * 1.25
            ReDim Preserve GPX(1 To GhostmaxNP)
            ReDim Preserve GPY(1 To GhostmaxNP)
            ReDim Preserve GPZ(1 To GhostmaxNP)

        End If
        GPX(GhostNP) = X
        GPY(GhostNP) = Y
        GPZ(GhostNP) = Z

    End If

    If NP < Capacity Then
        NP = NP + 1
        If NP > MaxNP Then
            MaxNP = Capacity
            ReDim px(1 To Capacity)
            ReDim py(1 To Capacity)
            ReDim pz(1 To Capacity)
            ReDim pIDX(1 To Capacity)
        End If
        px(NP) = X
        py(NP) = Y
        pz(NP) = Z
        pIDX(NP) = Index

    Else


        'If (X1 - X0) < H * 0.5 Or (Y1 - Y0) < H * 0.5 Then
        'Stop
        'End If


        If Not (Subdivided) Then Subdivide

        If X > MX Then Octant = 1&
        If Y > MY Then Octant = Octant Or 2&
        If Z > MZ Then Octant = Octant Or 4&

        QUA(Octant + 1).InsertSinglePoint X, Y, Z, Index, False

    End If

End Sub



Friend Sub InsertALLpoints(X() As Single, Y() As Single, Z() As Single, Optional PTBound As Long, Optional isRoot As Boolean = True)
    Dim I         As Long
    Reset
    If PTBound = 0 Then PTBound = UBound(X)
    For I = 1 To PTBound

        InsertSinglePoint X(I), Y(I), Z(I), I, isRoot

    Next
End Sub


Friend Sub QueryAroundPoint(ByVal CurrPointX As Single, ByVal CurrPointY As Single, ByVal CurrPointZ As Single, _
                            ByVal CurrPointIDX As Long, ByVal R As Single, _
                            RP1() As Long, _
                            RP2() As Long, _
                            Rdx() As Single, _
                            Rdy() As Single, _
                            Rdz() As Single, _
                            rDD() As Single, _
                            rPairsCount As Long, MAXNPairs As Long)

    Dim I         As Long
    Dim DX As Single, DY As Single, DZ As Single, R2 As Single, DD As Single

    If CurrPointX + R < X0 Then Exit Sub
    If CurrPointY + R < Y0 Then Exit Sub
    If CurrPointZ + R < Z0 Then Exit Sub
    If CurrPointX - R > X1 Then Exit Sub
    If CurrPointY - R > Y1 Then Exit Sub
    If CurrPointZ - R > Z1 Then Exit Sub

    R2 = R * R



    For I = 1 To NP
        If pIDX(I) < CurrPointIDX Then
            DX = px(I) - CurrPointX
            DY = py(I) - CurrPointY
            DZ = pz(I) - CurrPointZ
            DD = (DX * DX + DY * DY + DZ * DZ)
            If DD < R2 Then


                rPairsCount = rPairsCount + 1
                If rPairsCount > MAXNPairs Then
                    MAXNPairs = 1 + rPairsCount * 1.25
                    ReDim Preserve RP1(MAXNPairs)
                    ReDim Preserve RP2(MAXNPairs)
                    ReDim Preserve Rdx(MAXNPairs)
                    ReDim Preserve Rdy(MAXNPairs)
                    ReDim Preserve Rdz(MAXNPairs)
                    ReDim Preserve rDD(MAXNPairs)
                End If
                RP1(rPairsCount) = CurrPointIDX
                RP2(rPairsCount) = pIDX(I)
                Rdx(rPairsCount) = DX
                Rdy(rPairsCount) = DY
                Rdz(rPairsCount) = DZ
                rDD(rPairsCount) = DD
            End If
        End If
    Next

    '    Dim K         As Long

    If Subdivided Then
        '        If Q1.NP Then Q1.QueryAroundPoint CurrPointX, CurrPointY, CurrPointZ, CurrPointIDX, R, RP1, RP2, Rdx, Rdy, Rdz, rDD, rPairsCount, MAXNPairs 'Else: Q1.UnUsed = Q1.UnUsed + 1&: K = K + 1&
        '        If Q2.NP Then Q2.QueryAroundPoint CurrPointX, CurrPointY, CurrPointZ, CurrPointIDX, R, RP1, RP2, Rdx, Rdy, Rdz, rDD, rPairsCount, MAXNPairs 'Else: Q2.UnUsed = Q2.UnUsed + 1&: K = K + 1&
        '        If Q3.NP Then Q3.QueryAroundPoint CurrPointX, CurrPointY, CurrPointZ, CurrPointIDX, R, RP1, RP2, Rdx, Rdy, Rdz, rDD, rPairsCount, MAXNPairs 'Else: Q3.UnUsed = Q3.UnUsed + 1&: K = K + 1&
        '        If Q4.NP Then Q4.QueryAroundPoint CurrPointX, CurrPointY, CurrPointZ, CurrPointIDX, R, RP1, RP2, Rdx, Rdy, Rdz, rDD, rPairsCount, MAXNPairs 'Else: Q4.UnUsed = Q4.UnUsed + 1&: K = K + 1&
        '
        '        If Q5.NP Then Q5.QueryAroundPoint CurrPointX, CurrPointY, CurrPointZ, CurrPointIDX, R, RP1, RP2, Rdx, Rdy, Rdz, rDD, rPairsCount, MAXNPairs 'Else: Q5.UnUsed = Q5.UnUsed + 1&: K = K + 1&
        '        If Q6.NP Then Q6.QueryAroundPoint CurrPointX, CurrPointY, CurrPointZ, CurrPointIDX, R, RP1, RP2, Rdx, Rdy, Rdz, rDD, rPairsCount, MAXNPairs 'Else: Q6.UnUsed = Q6.UnUsed + 1&: K = K + 1&
        '        If Q7.NP Then Q7.QueryAroundPoint CurrPointX, CurrPointY, CurrPointZ, CurrPointIDX, R, RP1, RP2, Rdx, Rdy, Rdz, rDD, rPairsCount, MAXNPairs 'Else: Q7.UnUsed = Q7.UnUsed + 1&: K = K + 1&
        '        If Q8.NP Then Q8.QueryAroundPoint CurrPointX, CurrPointY, CurrPointZ, CurrPointIDX, R, RP1, RP2, Rdx, Rdy, Rdz, rDD, rPairsCount, MAXNPairs 'Else: Q8.UnUsed = Q8.UnUsed + 1&: K = K + 1&


        For I = 1 To 8
            If QUA(I).NP Then QUA(I).QueryAroundPoint CurrPointX, CurrPointY, CurrPointZ, CurrPointIDX, R, RP1, RP2, Rdx, Rdy, Rdz, rDD, rPairsCount, MAXNPairs    'Else: Q1.UnUsed = Q1.UnUsed + 1&: K = K + 1&
        Next


        '
        '        If K = 8& Then
        '            If Q1.UnUsed + Q2.UnUsed + Q3.UnUsed + Q4.UnUsed + _
                     '               Q5.UnUsed + Q6.UnUsed + Q7.UnUsed + Q8.UnUsed > 100000 Then
        '                Q1.UnUsed = 0
        '                Q2.UnUsed = 0
        '                Q3.UnUsed = 0
        '                Q4.UnUsed = 0
        '                Q5.UnUsed = 0
        '                Q6.UnUsed = 0
        '                Q7.UnUsed = 0
        '                Q8.UnUsed = 0
        '                Subdivided = 0
        '            End If
        '        End If

    End If

End Sub



Friend Sub QueryPointsInCircle(ByVal cX As Single, ByVal cY As Single, ByVal cZ As Single, ByVal R As Single, _
                               rdist() As Single, _
                               rIdx() As Long, _
                               rPntCount As Long, MaxrPntCount As Long, Optional isRoot As Boolean = True)
    Dim I         As Long
    Dim DX As Single, DY As Single, DZ As Single, R2 As Single, DD As Single
    If cX + R < X0 Then Exit Sub
    If cY + R < Y0 Then Exit Sub
    If cZ + R < Z0 Then Exit Sub
    If cX - R > X1 Then Exit Sub
    If cY - R > Y1 Then Exit Sub
    If cZ - R > Z1 Then Exit Sub


    R2 = R * R

    If isRoot Then rPntCount = 0   ': MaxrPntCount = 0

    For I = 1 To NP
        DX = px(I) - cX: DY = py(I) - cY: DZ = pz(I) - cZ
        DD = (DX * DX + DY * DY + DZ * DZ)
        If DD < R2 Then
            rPntCount = rPntCount + 1
            If rPntCount > MaxrPntCount Then
                MaxrPntCount = rPntCount * 2
                ReDim Preserve rIdx(MaxrPntCount)
                ReDim Preserve rdist(MaxrPntCount)
            End If
            rdist(rPntCount) = DD
            rIdx(rPntCount) = pIDX(I)

        End If
    Next


    If Subdivided Then
        '        If Q1.NP Then Q1.QueryPointsInCircle cX, cY, cZ, R, rdist, rIdx, rPntCount, MaxrPntCount, False
        '        If Q2.NP Then Q2.QueryPointsInCircle cX, cY, cZ, R, rdist, rIdx, rPntCount, MaxrPntCount, False
        '        If Q3.NP Then Q3.QueryPointsInCircle cX, cY, cZ, R, rdist, rIdx, rPntCount, MaxrPntCount, False
        '        If Q4.NP Then Q4.QueryPointsInCircle cX, cY, cZ, R, rdist, rIdx, rPntCount, MaxrPntCount, False
        '
        '        If Q5.NP Then Q5.QueryPointsInCircle cX, cY, cZ, R, rdist, rIdx, rPntCount, MaxrPntCount, False
        '        If Q6.NP Then Q6.QueryPointsInCircle cX, cY, cZ, R, rdist, rIdx, rPntCount, MaxrPntCount, False
        '        If Q7.NP Then Q7.QueryPointsInCircle cX, cY, cZ, R, rdist, rIdx, rPntCount, MaxrPntCount, False
        '        If Q8.NP Then Q8.QueryPointsInCircle cX, cY, cZ, R, rdist, rIdx, rPntCount, MaxrPntCount, False
        For I = 1 To 8
            If QUA(I).NP Then QUA(I).QueryPointsInCircle cX, cY, cZ, R, rdist, rIdx, rPntCount, MaxrPntCount, False

        Next
    End If

End Sub
''Friend Sub QuerySquare(cX As single, _
 ''                       cY As single, _
 ''                       halfSide As single, _
 ''                       rpX() As single, rpY() As single, rpIDX() As Long, Found As Long, MaxFound As Long, Optional isRoot As Boolean = True)
''    Dim I         As Long
''    Dim DX as single, DY as single
''    If cX + halfSide < X0 Then Exit Sub
''    If cY + halfSide < Y0 Then Exit Sub
''    If cX - halfSide > X1 Then Exit Sub
''    If cY - halfSide > Y1 Then Exit Sub
''
''
''    If isRoot Then Found = 0      ': MaxFound = 0
''
''    For I = 1 To NP
''        DX = px(I) - cX: DY = py(I) - cY
''        If Abs(DX) < halfSide And Abs(DY) < halfSide Then
''            Found = Found + 1
''            If Found > MaxFound Then
''                MaxFound = Found * 2
''                ReDim Preserve rpIDX(MaxFound)
''                ReDim Preserve rpX(MaxFound)
''                ReDim Preserve rpY(MaxFound)
''
''            End If
''            rpX(Found) = px(I)
''            rpY(Found) = py(I)
''            rpIDX(Found) = pIDX(I)
''
''        End If
''    Next
''
''    If Subdivided Then
''        If Q1.NP Then Q1.QuerySquare cX, cY, halfSide, rpX, rpY, rpIDX, Found, MaxFound, False
''        If Q2.NP Then Q2.QuerySquare cX, cY, halfSide, rpX, rpY, rpIDX, Found, MaxFound, False
''        If Q3.NP Then Q3.QuerySquare cX, cY, halfSide, rpX, rpY, rpIDX, Found, MaxFound, False
''        If Q4.NP Then Q4.QuerySquare cX, cY, halfSide, rpX, rpY, rpIDX, Found, MaxFound, False
''    End If
''End Sub

''Friend Sub QueryCircleOLD(ByVal cX As single, ByVal cY As single, ByVal R As single, _
 ''                          rpX() As single, rpY() As single, rpIDX() As Long, Found As Long, MaxFound As Long, Optional isRoot As Boolean = True)
''    Dim I         As Long
''    Dim DX as single, DY as single, R2 as single, DD as single
''    If cX + R < X0 Then Exit Sub
''    If cY + R < Y0 Then Exit Sub
''    If cX - R > X1 Then Exit Sub
''    If cY - R > Y1 Then Exit Sub
''
''    R2 = R * R
''
''    If isRoot Then Found = 0      ': maxfound = 0
''
''    For I = 1 To NP
''        DX = px(I) - cX: DY = py(I) - cY: DD = (DX * DX + DY * DY)
''        If DD < R2 Then
''            Found = Found + 1
''            If Found > MaxFound Then
''                MaxFound = Found * 2
''                ReDim Preserve rpIDX(MaxFound)
''                ReDim Preserve rpX(MaxFound)
''                ReDim Preserve rpY(MaxFound)
''
''            End If
''            rpX(Found) = px(I)
''            rpY(Found) = py(I)
''            rpIDX(Found) = pIDX(I)
''
''        End If
''    Next
''
''    If Subdivided Then
''        If Q1.NP Then Q1.QueryCircleOLD cX, cY, R, rpX, rpY, rpIDX, Found, MaxFound, False
''        If Q2.NP Then Q2.QueryCircleOLD cX, cY, R, rpX, rpY, rpIDX, Found, MaxFound, False
''        If Q3.NP Then Q3.QueryCircleOLD cX, cY, R, rpX, rpY, rpIDX, Found, MaxFound, False
''        If Q4.NP Then Q4.QueryCircleOLD cX, cY, R, rpX, rpY, rpIDX, Found, MaxFound, False
''    End If
''
''End Sub
Friend Sub GetPairsWDist(ByVal Radius As Single, RP1() As Long, RP2() As Long, Rdx() As Single, Rdy() As Single, Rdz() As Single, rDD() As Single, rPairsCount As Long, MAXNPairs As Long)
    Dim I         As Long
    '  USING GHOST
    rPairsCount = 0
    For I = 1 To GhostNP
        QueryAroundPoint GPX(I), GPY(I), GPZ(I), I, Radius, RP1, RP2, Rdx, Rdy, Rdz, rDD, rPairsCount, MAXNPairs
    Next

    '   Without GHOST
    ''    For I = 1 To NP
    ''        QueryAroundPoint PX(I), PY(I), Pidx(I), Radius, rP1, rP2, rDX, rDY, rDD, rPairsCount, MaxNpairs
    ''    Next
    ''    If Not (NotSubdivided) Then
    ''        If Q1.NP Then Q1.GetPairsWDist Radius, rP1, rP2, rDX, rDY, rDD, rPairsCount, MaxNpairs
    ''        If Q2.NP Then Q2.GetPairsWDist Radius, rP1, rP2, rDX, rDY, rDD, rPairsCount, MaxNpairs
    ''        If Q3.NP Then Q3.GetPairsWDist Radius, rP1, rP2, rDX, rDY, rDD, rPairsCount, MaxNpairs
    ''        If Q4.NP Then Q4.GetPairsWDist Radius, rP1, rP2, rDX, rDY, rDD, rPairsCount, MaxNpairs
    ''    End If

End Sub








Friend Sub DrawToCC(CC As cCairoContext, Optional Clean As Boolean = False)
    Dim I         As Long
    '
    '    If Clean Then CC.SetSourceColor 0: CC.Paint
    '
    '    CC.SetSourceColor vbGreen
    '
    '    CC.Rectangle X0, Y0, X1 - X0, Y1 - Y0: CC.Stroke
    '
    '    CC.SetSourceColor vbYellow
    '    For I = 1 To NP
    '        CC.Arc px(I), py(I), 1.5: CC.Stroke
    '    Next
    '
    '    If Subdivided Then
    '        If Q1.NP Then Q1.DrawToCC CC
    '        If Q2.NP Then Q2.DrawToCC CC
    '        If Q3.NP Then Q3.DrawToCC CC
    '        If Q4.NP Then Q4.DrawToCC CC
    '    End If

End Sub
