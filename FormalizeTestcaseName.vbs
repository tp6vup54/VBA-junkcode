Sub a()
    Count = 1
    For i = 7 To 536
        If Range("B" & i).Value <> "" And Range("C" & i).Value <> "" Then
            LastUnderlineIndex = InStrRev(Range("B" & i).Value, "_")
            LastSecUnderlineIndex = InStrRev(Range("B" & i).Value, "_", LastUnderlineIndex - 1)
            TailCount = Format(Count, "0000")
            TestcaseName = Left(Range("B" & i).Value, LastSecUnderlineIndex - 1)
            Range("B" & i).Value = TestcaseName & "_" & Range("H" & i).Value & "_" & TailCount
            Count = Count + 1
        End If
    Next
End Sub
