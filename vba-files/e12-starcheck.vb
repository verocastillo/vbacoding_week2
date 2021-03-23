' Loops and Counters: Star Checker
'
' Instructions: count the stars from each student, and display
' the result in other column.
'
' Test only on e12.xlsm

Sub StarCounter()
    ' Dim and give value to the variable
    Dim CountVar As Integer
    CountVar = 0
    ' For loop checks through rows and columns, counts full stars
    For i = 2 To 51
        For j = 4 To 8
            If Cells(i, j).Value = "Full-Star" Then
            CountVar = CountVar + 1
            End If
            Cells(i, 9).Value = CountVar
        Next j
        CountVar = 0
    Next i
    
End Sub

' BONUS: using VBA to determine the last row.
Sub StarCounterBonus1()
    ' Dim and give value to the variables
    Dim CountVar As Integer
    Dim LastRow As Long
    ' Find last row
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    ' For loop checks through rows and columns, counts full stars
    For i = 2 To LastRow
        For j = 4 To 8
            If Cells(i, j).Value = "Full-Star" Then
            CountVar = CountVar + 1
            End If
            Cells(i, 9).Value = CountVar
        Next j
        CountVar = 0
    Next i
End Sub


