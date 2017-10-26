Sub a()
    Count = 0
    For i = 7 To 536
        If InStr(Range("H" & i).Value, "RAT") or InStr(Range("H" & i).Value, "FAST") Then
            Range("G" & i).Value = ""
            Count = Count + 1
        End If
    Next
End Sub