' Conditionals: Choose a Story
'
' Instructions: using conditionals, create a different story
' for each of the buttons you click in the document.
'
' Test only on e06.xlsm

Sub Conditions()
    ' Conditionals for different inputs, input number is valid from 1-4, 
    ' and it's taken from cell A2
    If (Range("A2").Value = 1) Then
        MsgBox ("You choose to enter the wooded forest of doom!")

    ElseIf (Range("A2").Value = 2) Then
        MsgBox ("You choose to enter the fiery volcano of doom!")

    ElseIf (Range("A2").Value = 3) Then
        MsgBox ("You choose to enter the terrifying jungle of doom!")

    ElseIf (Range("A2").Value = 4) Then
        MsgBox ("You choose to go back home and watch Netflix instead of impending doom")

    Else
        MsgBox ("Try following directions")

    End If
End Sub
