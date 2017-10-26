Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function

Sub a()
    Dim run
    run = Array("0004", "0011", "0030", "0042", "0051", "0001", "0002", "0028", "0071")
    For i = 7 To 536
        If InStr(Range("H" & i).Value, "RAT") Or InStr(Range("H" & i).Value, "FAST") Then
            If IsInArray(Right(Range("B" & i).Value, 4), run) Or (InStr(LCase(Range("N" & i).Value), "merged") And IsInArray(Right(Range("N" & i).Value, 4), run)) Then
                Range("G" & i).Value = "pass"
            End If
        End If
    Next
End Sub
