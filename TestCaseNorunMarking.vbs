Sub a()
    Count = 0
    For i = 7 To 100
        If InStr(Range("H" & i).Value, "TOFT") Then
            Range("G" & i).Value = ""
            Count = Count + 1
        ElseIf Not Range("H" & i).Value = "" Then
            Range("G" & i).Value = "norun"
        End If
    Next
    MsgBox (Count)
End Sub
