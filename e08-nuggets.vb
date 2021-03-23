' For Loops: Chicken Nuggets
'
' Instructions: using a for loop, create a script that
' writes how many chicken nuggets you will eat.

Sub Nuggets()
    ' For loop: 10 rows
    For i = 1 To 10
        ' Column 1
        Cells(i, 1).Value = "I will eat "
        ' Column 2
        Cells(i, 2).Value = i + 10
        ' Column 3
        Cells(i, 3).Value = "Chicken Nuggets"
    Next i
End Sub
