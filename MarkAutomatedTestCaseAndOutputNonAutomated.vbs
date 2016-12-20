Sub a()
    testCaseFolderPath = "C:\STAF\testsuites\sCloud\CasePool\EventNotifications\"
    testCaseName = "EventNotifications"
    aa = getTestCaseNames(testCaseFolderPath, testCaseName)
    case_type = "FAST"
    Open "C:\Users\sean_c_chen\Desktop\bbb.txt" For Output As #1
    
    For i = 7 To 536
        If IsInArray(Right(Range("B" & i).Value, 3), aa) Or InStr(LCase(Range("C" & i).Value), "but cancel") Then
            Range("J" & i).Value = "Yes"
        Else
            Range("J" & i).Value = ""
        End If
    Next
    
    For i = 7 To 536
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

Function getTestCaseNames(folderPath, testCaseName)
    Dim ret(536)
    i = 0
    file = Dir(folderPath)
    While (file <> "")
        If InStr(file, testCaseName & "_") > 0 Then
            ret(i) = Left(Right(file, 6), 3)
            i = i + 1
        End If
        file = Dir
    Wend
    getTestCaseNames = ret
End Function

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function