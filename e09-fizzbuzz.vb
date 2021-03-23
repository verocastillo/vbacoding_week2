' Loops and Conditionals: FizzBuzz
'
' Instructions: using loops and conditionals, fill the
' second column accordingly.
'
' Test only on e09.xlsm

Sub FizzBuzz()
    ' Foor loop to check column.
    For i = 2 To 100
        Number = Cells(i, 1).Value
        
        ' Is it divisible by 3 and 5?
        If (Number Mod 3 = 0 And Number Mod 5 = 0) Then
                Cells(i, 2).Value = "Fizzbuzz"
    
        ' Is it divisible by 3?
        ElseIf (Number Mod 3 = 0) Then
                Cells(i, 2).Value = "Fizz"
        
        ' Is it divisible by 5?
        ElseIf (Number Mod 5 = 0) Then
            Cells(i, 2).Value = "Buzz"
        
        ' If it is neither:
        Else
            Cells(i, 2).Value = " "
        
        End If
        
    Next i
    
End Sub
