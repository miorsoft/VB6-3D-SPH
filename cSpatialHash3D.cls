VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSpatialHash3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'https://github.com/matthias-research/pages/blob/master/tenMinutePhysics/11-hashing.html
Private Declare Sub AssignZero Lib "kernel32" Alias "RtlZeroMemory" (pDst As Any, Optional ByVal cb& = 4)


Private Spacing   As Long
Attribute Spacing.VB_VarUserMemId = 1073938434
Private TableSize As Long
Attribute TableSize.VB_VarUserMemId = 1073938435
Private cellStart() As Long
Private cellEntries() As Long
Attribute cellEntries.VB_VarUserMemId = 1073938435
Private mQueryIds() As Long
Attribute mQueryIds.VB_VarUserMemId = 1610809344
Public querySize  As Long
Attribute querySize.VB_VarUserMemId = 1073938432


Friend Property Get QueryIds(I As Long) As Long
    QueryIds = mQueryIds(I)
End Property
Friend Sub constructor(mSpacing As Long, maxNumObjects As Long)
    Spacing = mSpacing
    TableSize = 2 * maxNumObjects
    ReDim cellStart(TableSize + 1)
    ReDim cellEntries(maxNumObjects)
    ReDim mQueryIds(maxNumObjects)
    querySize = 0
End Sub
'hashCoords(xi, yi, zi) {
'                    var h = (xi * 92837111) ^ (yi * 689287499) ^ (zi * 283923481);  // fantasy function
'                    return Math.abs(h) % this.tableSize;
'                }

Friend Function hashCoords(Xi As Single, Yi As Single, Zi As Single) As Long
    Dim H         As Single
    'h = (Xi * 92837111) ^ (Yi * 689287499) ^ (Zi * 283923481)    '  // fantasy function
    H = (Xi * 0.92837111) * (Yi * 6.89287499) * (Zi * 2.83923481)    '  // fantasy function

    hashCoords = Abs(H) Mod TableSize
End Function

'intCoord(coord) {
'                    return Math.floor(coord / this.spacing);
'                }

Friend Function intCoord(coord As Single) As Long
    intCoord = coord \ Spacing
End Function



'                hashPos(pos, nr) {
'                    return this.hashCoords(
'                        this.intCoord(pos[3 * nr]),
'                        this.intCoord(pos[3 * nr + 1]),
'                        this.intCoord(pos[3 * nr + 2]));
'                }

Friend Function hashPos(Pos As tVec3) As Long

    hashPos = hashCoords(intCoord(Pos.X), intCoord(Pos.Y), intCoord(Pos.Z))

End Function






'create(pos) {
'                    var numObjects = Math.min(pos.length / 3, this.cellEntries.length);
'
'                    // determine cell sizes
'
'                    this.cellStart.fill(0);
'                    this.cellEntries.fill(0);
'
'                    for (var i = 0; i < numObjects; i++) {
'                        var h = this.hashPos(pos, i);
'                        this.cellStart[h]++;
'                    }
'
'                    // determine cells starts
'
'                    var start = 0;
'                    for (var i = 0; i < this.tableSize; i++) {
'                        start += this.cellStart[i];
'                        this.cellStart[i] = start;
'                    }
'                    this.cellStart[this.tableSize] = start; // guard
'
'                    // fill in objects ids
'
'                    for (var i = 0; i < numObjects; i++) {
'                        var h = this.hashPos(pos, i);
'                        this.cellStart[h]--;
'                        this.cellEntries[this.cellStart[h]] = i;
'                    }
'                }


Friend Sub InsertPoints2(X() As Single, Y() As Single, Z() As Single)
    Dim numObjects&
    Dim I&, Start&, H&
    numObjects = UBound(X)
    If UBound(cellEntries) < numObjects Then numObjects = UBound(cellEntries)

    ' // determine cell sizes

    '                    this.cellStart.fill(0);
    '                    this.cellEntries.fill(0);

    '    ReDim cellStart(TableSize + 1)    ' Da cambiare in fill 0
    '    ReDim cellEntries(UBound(cellEntries))    ' Da cambiare in fill 0

    AssignZero cellStart(0), LenB(cellStart(0)) * (TableSize + 1)
    AssignZero cellEntries(0), LenB(cellEntries(0)) * (UBound(cellEntries) + 1)

    For I = 1 To numObjects
        H = hashPos(Vec3(X(I), Y(I), Z(I)))
        cellStart(H) = cellStart(H) + 1
    Next

    '// determine cells starts

    Start = 0
    For I = 1 To TableSize
        Start = Start + cellStart(I)
        cellStart(I) = Start
    Next
    cellStart(TableSize) = Start   '; // guard

    ' // fill in objects ids

    For I = 1 To numObjects
        H = hashPos(Vec3(X(I), Y(I), Z(I)))
        ' this.cellStart[h]--;
        ' this.cellEntries[this.cellStart[h]] = i;
        cellStart(H) = cellStart(H) - 1&
        cellEntries(cellStart(H)) = I
    Next
End Sub
Friend Sub InsertPoints(Pos() As tVec3)

    '                    var numObjects = Math.min(pos.length / 3, this.cellEntries.length);
    Dim numObjects&
    Dim I&, Start&, H&
    numObjects = UBound(Pos)
    If UBound(cellEntries) < numObjects Then numObjects = UBound(cellEntries)

    ' // determine cell sizes

    '                    this.cellStart.fill(0);
    '                    this.cellEntries.fill(0);

    '    ReDim cellStart(TableSize + 1)    ' Da cambiare in fill 0
    '    ReDim cellEntries(UBound(cellEntries))    ' Da cambiare in fill 0

    'Stop

    AssignZero cellStart(0), LenB(cellStart(0)) * (TableSize + 1)
    AssignZero cellEntries(0), LenB(cellEntries(0)) * (UBound(cellEntries))

    For I = 1 To numObjects
        H = hashPos(Pos(I))
        cellStart(H) = cellStart(H) + 1
    Next

    '// determine cells starts

    Start = 0
    For I = 1 To TableSize
        Start = Start + cellStart(I)
        cellStart(I) = Start
    Next
    cellStart(TableSize) = Start   '; // guard

    ' // fill in objects ids

    For I = 1 To numObjects
        H = hashPos(Pos(I))
        ' this.cellStart[h]--;
        ' this.cellEntries[this.cellStart[h]] = i;
        cellStart(H) = cellStart(H) - 1&
        cellEntries(cellStart(H)) = I
    Next
End Sub


'query(pos, nr, maxDist) {
'                    var x0 = this.intCoord(pos[3 * nr] - maxDist);
'                    var y0 = this.intCoord(pos[3 * nr + 1] - maxDist);
'                    var z0 = this.intCoord(pos[3 * nr + 2] - maxDist);
'
'                    var x1 = this.intCoord(pos[3 * nr] + maxDist);
'                    var y1 = this.intCoord(pos[3 * nr + 1] + maxDist);
'                    var z1 = this.intCoord(pos[3 * nr + 2] + maxDist);
'
'                    this.querySize = 0;
'
'                    for (var xi = x0; xi <= x1; xi++) {
'                        for (var yi = y0; yi <= y1; yi++) {
'                            for (var zi = z0; zi <= z1; zi++) {
'                                var h = this.hashCoords(xi, yi, zi);
'                                var iStart = this.cellStart[h];
'                                var end = this.cellStart[h + 1];
'
'                                for (var i = iStart; i < end; i++) {
'                                    this.mqueryIds[this.querySize] = this.cellEntries[i];
'                                    this.querySize++;
'                                }
'                            }
'                        }
'                    }
'                }

Friend Sub Query(Pos As tVec3, maxDist As Single)
    Dim X0&, Y0&, Z0&
    Dim X1&, Y1&, Z1&
    Dim Xi&, Yi&, Zi&
    Dim H&, iStart&, iEnd&
    Dim I&
    X0 = intCoord(Pos.X - maxDist)
    Y0 = intCoord(Pos.Y - maxDist)
    Z0 = intCoord(Pos.Z - maxDist)

    X1 = intCoord(Pos.X + maxDist)
    Y1 = intCoord(Pos.Y + maxDist)
    Z1 = intCoord(Pos.Z + maxDist)

    querySize = 0

    For Xi = X0 To X1
        For Yi = Y0 To Y1
            For Zi = Z0 To Z1
                H = hashCoords(Xi * 1&, Yi * 1&, Zi * 1&)
                iStart = cellStart(H)
                iEnd = cellStart(H + 1&)
                '
                '                                for (var i = iStart; i < end; i++) {
                '                                    this.mqueryIds[this.querySize] = this.cellEntries[i];
                '                                    this.querySize++;

                For I = iStart To iEnd
                    querySize = querySize + 1
                    mQueryIds(querySize) = cellEntries(I)

                Next

            Next
        Next
    Next

End Sub
