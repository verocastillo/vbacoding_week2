' Nested Loops: Hornet Infestation
'
' Instructions: using nested loops and different functions,
' fight the hornet infestation and complete the three parts
' specified on the assignment.
'
' Test only on e11.xlsm

' Part I: Count the number of Hornets Found

Sub Hornets1()
  ' Dim variables and set values
  Dim Count As Integer
  Count = 0
  ' Loop through all rows and columns
  For i = 1 To 6
    For j = 1 To 7
      ' Check if the value is a hornet and add to count
      If Cells(i, j).Value = "Hornets" Then
        Count = Count + 1
      End If
    Next j
  Next i
  ' Display the hornets count
  MsgBox (Count & " Hornets were found!")
End Sub

' Part II: Each time you find Hornets replace them with Bugs
Sub Hornets2()
  ' Dim variables and set values
  Dim Count As Integer
  Count = 0
  ' Loop through all rows and columns
  For i = 1 To 6
    For j = 1 To 7
      ' Check if the value is a hornet and add to count
      If Cells(i, j).Value = "Hornets" Then
        Count = Count + 1
        ' Replace the with bugs
        Cells(i, j).Value = "Bugs"
      End If
    Next j
  Next i
  ' Display the hornets count
  MsgBox (Count & " Hornets were found!")
End Sub

' Part III: You have a max amount of Bees and Hornets, utilize no more than what's provided.

Sub Hornets3()
  ' Dim variables and set values
  Dim Count As Integer
  Dim Bugs As Integer
  Dim Bees As Integer
  Count = 0
  Bugs = Range("L2").Value
  Bees = Range("R2").Value
  ' Loop through all rows and columns
  For i = 1 To 6
    For j = 1 To 7
      ' Check if the value is a hornet and add to count
      If Cells(i, j).Value = "Hornets" Then
        Count = Count + 1
        ' Checks if bugs and bees are available, and substract for each use
        If (Bugs > 0) Then
          Cells(i, j).Value = "Bugs"
          Bugs = Bugs - 1
        ElseIf (Bees > 0) Then
          Cells(i, j).Value = "Bees"
          Bees = Bees - 1
        End If
      End If
    Next j
  Next i
  ' Show the number of hornets found
  MsgBox (Count & " Hornets were found!")
  ' Create the final message if we still have hornets
  If (Range("L2").Value + Range("R2").Value < Count) Then
    MsgBox ("Oh no! We still have hornets... ")
  End If
End Sub

