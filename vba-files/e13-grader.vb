' Advanced Conditionals: Grader
'
' Instructions: determine the status and letter grade of 
' a student, including formatting.
'
' Test only on e13.xlsm

Sub Grades()
    ' Conditionals for each letter grade.
    If Cells(2, 2).Value >= 90 Then
        Cells(2, 3).Value = "Pass"
        Cells(2, 4).Value = "A"
        Cells(2, 3).Interior.ColorIndex = 4
    ElseIf Cells(2, 2).Value >= 80 Then
        Cells(2, 3).Value = "Pass"
        Cells(2, 4).Value = "B"
        Cells(2, 3).Interior.ColorIndex = 4
    ElseIf Cells(2, 2).Value >= 70 Then
        Cells(2, 3).Value = "Warning"
        Cells(2, 4).Value = "C"
        Cells(2, 3).Interior.ColorIndex = 6
    Else
        Cells(2, 3).Value = "Failed"
        Cells(2, 4).Value = "F"
        Cells(2, 3).Interior.ColorIndex = 3
    End If
End Sub

'BONUS: Reset everything.
Sub Reset()
    ' Copy last values to defined cells
    Cells(12, 2) = Cells(2, 2).Value
    Cells(12, 3) = Cells(2, 3).Value
    Cells(12, 4) = Cells(2, 4).Value
    ' Clear grader
    Cells(2, 2).Value = " "
    Cells(2, 3).Value = " "
    Cells(2, 4).Value = " "
    Cells(12, 1).Value = "Last grade"
    Cells(2, 3).Interior.ColorIndex = 2
End Sub
