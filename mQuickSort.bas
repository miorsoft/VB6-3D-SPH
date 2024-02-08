Attribute VB_Name = "mQuickSort"
Option Explicit
'http://www.vbforums.com/showthread.php?231925-VB-Quick-Sort-algorithm-(very-fast-sorting-algorithm)&p=4739885&viewfull=1#post4739885

Public SORTSWAPS  As Long
Attribute SORTSWAPS.VB_VarUserMemId = 1073938433


Public Sub QuicksortSingle(List() As Single, ByVal Min As Long, ByVal Max As Long)
Attribute QuicksortSingle.VB_UserMemId = 1073938434
    ' from Low to hi
    Dim med_value As Single
    Dim hi        As Long
    Dim lo        As Long
    Dim I         As Long
    If Max <= Min Then Exit Sub
    'I = Int((max - min + 1) * Rnd + min)
    I = (Max + Min) \ 2
    med_value = List(I)
    List(I) = List(Min)
    lo = Min
    hi = Max
    Do
        Do While List(hi) >= med_value
            hi = hi - 1&
            If hi <= lo Then Exit Do
        Loop
        If hi <= lo Then
            List(lo) = med_value
            Exit Do
        End If
        List(lo) = List(hi)
        lo = lo + 1
        Do While List(lo) < med_value
            lo = lo + 1&
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
    QuicksortSingle List(), Min, lo - 1&
    QuicksortSingle List(), lo + 1&, Max
End Sub


Public Sub QuickSortSingle2(Dist() As Single, OtherInfo() As Long, ByVal Min As Long, ByVal Max As Long)
Attribute QuickSortSingle2.VB_UserMemId = 1073741850

    Dim med_value As Single
    Dim med_OtherInfo As Long

    Dim hi        As Long
    Dim lo        As Long
    Dim I         As Long
    If Max <= Min Then Exit Sub
    '  I = Int((max - min + 1) * Rnd + min)
    I = (Max + Min) \ 2

    med_value = Dist(I)
    med_OtherInfo = OtherInfo(I)

    Dist(I) = Dist(Min)
    OtherInfo(I) = OtherInfo(Min)

    lo = Min
    hi = Max
    Do
        Do While Dist(hi) >= med_value
            hi = hi - 1&
            If hi <= lo Then Exit Do
        Loop
        If hi <= lo Then
            Dist(lo) = med_value
            OtherInfo(lo) = med_OtherInfo
            Exit Do
        End If
        Dist(lo) = Dist(hi)
        OtherInfo(lo) = OtherInfo(hi)
        lo = lo + 1&

        Do While Dist(lo) < med_value
            lo = lo + 1&
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
        SORTSWAPS = SORTSWAPS + 1&
    Loop
    QuickSortSingle2 Dist(), OtherInfo(), Min, lo - 1&
    QuickSortSingle2 Dist(), OtherInfo(), lo + 1&, Max
End Sub



Public Sub QuickSortSingle3(Dist() As Single, OtherInfo() As Long, ByVal Min As Long, ByVal Max As Long)
Attribute QuickSortSingle3.VB_UserMemId = 1610612744
    ' FROM HI to LOW  'https://www.vbforums.com/showthread.php?11192-quicksort
    Dim Low As Long, high As Long, temp As Single, TestElement As Single, tmp&
    '     Debug.Print min, max
    Low = Min: high = Max
    '    TestElement = Dist((Min + Max) \ 2)
    TestElement = (Dist(Min) + Dist(Max)) * 0.5
    Do
        Do While Dist(Low) > TestElement: Low = Low + 1&: Loop
        Do While Dist(high) < TestElement: high = high - 1&: Loop
        If (Low <= high) Then
            temp = Dist(Low): Dist(Low) = Dist(high): Dist(high) = temp
            tmp = OtherInfo(Low): OtherInfo(Low) = OtherInfo(high): OtherInfo(high) = tmp
            Low = Low + 1&: high = high - 1&
            SORTSWAPS = SORTSWAPS + 1&
        End If
    Loop While (Low <= high)
    If (Min < high) Then QuickSortSingle3 Dist(), OtherInfo(), Min, high
    If (Low < Max) Then QuickSortSingle3 Dist(), OtherInfo(), Low, Max
End Sub



