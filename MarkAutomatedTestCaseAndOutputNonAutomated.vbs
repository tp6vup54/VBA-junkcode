Sub a()
    aa = Array("000", "001", "004", "017", "030", "043", "056", "067", "071", "075", "080", "092", "100", "112", "124", "134", "142", "151", "160", "172", "180", "190", "198", "203", "211", "216", "224", "229", "237", "242", "250", "255", "263", "269", "271", "310", "316", "324", "331", "338", "345", "352", "359", "366", "373", "379", "386", "395", "401", "409", "418", "426", "432", "440", "446", "491", "494", "074", "186", "268")
    case_type = "FAST"
    Open "D:\bbb.txt" For Output As #1
    
    For i = 7 To 534
        If IsInArray(Right(Range("B" & i).Value, 3), aa) Or InStr(LCase(Range("C" & i).Value), "but cancel") Then
            Range("J" & i).Value = "Yes"
        Else
            Range("J" & i).Value = ""
        End If
    Next
    
    For i = 7 To 534
        If Range("H" & i).Value = case_type And Not Range("J" & i) = "Yes" And Not Range("M" & i) = "Disable" Then
            append_str = ""
            If InStr(Range("C" & i), "trigger") Or InStr(Range("C" & i), "Trigger") Then
                append_str = "trigger alert"
            ElseIf InStr(Range("C" & i), "test") Or InStr(Range("C" & i), "Test") Then
                append_str = "test notification"
            End If
            Print #1, Range("B" & i).Value & Chr(9) & append_str
        End If
    Next
    Close #1
End Sub

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
  IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function
