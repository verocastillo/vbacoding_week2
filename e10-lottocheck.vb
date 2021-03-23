' Advanced Loops: Lotto Checker
'
' Instructions: using advanced loops and conditionals, determine
' the winners of the lotto
'
' Test only on e10.xlsm

Sub Winner()
    ' Create and dim necessary variables, give values
    Dim First As Long
    First = 3957481
    Dim Secondd As Long
    Secondd = 5865187
    Dim Third As Long
    Third = 2817729
    Dim RunnerUp1 As Long
    RunnerUp1 = 2275339
    Dim RunnerUp2 As Long
    RunnerUp2 = 5868182
    Dim RunnerUp3 As Long
    RunnerUp3 = 1841402
    Dim RunnerUp As Long

    ' For loop to check all tickets for winners:
    For i = 1 To 1001
        ' Check if it's first place:
        If Cells(i, 3).Value = First Then
            MsgBox " Congratulations " + Cells(i, 1).Value
            ' Put values into corresponding cells
            Cells(2, 6).Value = Cells(i, 1).Value
            Cells(2, 7).Value = Cells(i, 2).Value
            Cells(2, 8).Value = First
        ' Same for second place:
        ElseIf Cells(i, 3).Value = Secondd Then
            Cells(3, 6).Value = Cells(i, 1).Value
            Cells(3, 7).Value = Cells(i, 2).Value
            Cells(3, 8).Value = Second
        ' Same for third place:
        ElseIf Cells(i, 3).Value = Third Then
            Cells(4, 6).Value = Cells(i, 1).Value
            Cells(4, 7).Value = Cells(i, 2).Value
            Cells(4, 8).Value = Third
        End If
    Next i

    ' For loop to check runner ups:
    For i = 1 to 1001
        ' Conditionals using or:
        If Cells(i, 3).Value = RunnerUp1 Or Cells(i, 3).Value = RunnerUp2 Or Cells(i, 3).Value = RunnerUp3 Then
            RunnerUp = Cells(i, 3).Value
            Cells(5, 6).Value = Cells(i, 1).Value
            Cells(5, 7).Value = Cells(i, 2).Value
            Cells(5, 8).Value = RunnerUp
            Exit for
        End If
    Next i 

End Sub
