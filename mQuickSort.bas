Attribute VB_Name = "mQuickSort"
Option Explicit
'http://www.vbforums.com/showthread.php?231925-VB-Quick-Sort-algorithm-(very-fast-sorting-algorithm)&p=4739885&viewfull=1#post4739885

Public Sub QuicksortSingle(List() As Double, ByVal min As Long, ByVal max As Long)
' from Low to hi
    Dim med_value As Double
    Dim hi        As Long
    Dim lo        As Long
    Dim I         As Long
    If max <= min Then Exit Sub
    I = Int((max - min + 1) * Rnd + min)
    med_value = List(I)
    List(I) = List(min)
    lo = min
    hi = max
    Do
        Do While List(hi) >= med_value
            hi = hi - 1
            If hi <= lo Then Exit Do
        Loop
        If hi <= lo Then
            List(lo) = med_value
            Exit Do
        End If
        List(lo) = List(hi)
        lo = lo + 1
        Do While List(lo) < med_value
            lo = lo + 1
            If lo >= hi Then Exit Do
        Loop
        If lo >= hi Then
            lo = hi
            List(hi) = med_value
            Exit Do
        End If

        ' Swap the lo and hi values.
        List(hi) = List(lo)
    Loop
    QuicksortSingle List(), min, lo - 1
    QuicksortSingle List(), lo + 1, max
End Sub


Public Sub QuickSortSingle2(Dist() As Double, OtherInfo() As Long, ByVal min As Long, ByVal max As Long)

    Dim med_value As Double
    Dim med_OtherInfo As Long

    Dim hi        As Long
    Dim lo        As Long
    Dim I         As Long
    If max <= min Then Exit Sub
    I = Int((max - min + 1) * Rnd + min)
    med_value = Dist(I)
    med_OtherInfo = OtherInfo(I)

    Dist(I) = Dist(min)
    OtherInfo(I) = OtherInfo(min)

    lo = min
    hi = max
    Do
        Do While Dist(hi) >= med_value
            hi = hi - 1
            If hi <= lo Then Exit Do
        Loop
        If hi <= lo Then
            Dist(lo) = med_value
            OtherInfo(lo) = med_OtherInfo
            Exit Do
        End If
        Dist(lo) = Dist(hi)
        OtherInfo(lo) = OtherInfo(hi)
        lo = lo + 1
        Do While Dist(lo) < med_value
            lo = lo + 1
            If lo >= hi Then Exit Do
        Loop
        If lo >= hi Then
            lo = hi
            Dist(hi) = med_value
            OtherInfo(hi) = med_OtherInfo
            Exit Do
        End If

        ' Swap the lo and hi values.
        Dist(hi) = Dist(lo)
        OtherInfo(hi) = OtherInfo(lo)

    Loop
    QuickSortSingle2 Dist(), OtherInfo(), min, lo - 1
    QuickSortSingle2 Dist(), OtherInfo(), lo + 1, max
End Sub

