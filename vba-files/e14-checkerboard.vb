' Coding Logic: Checkerboard
'
' Instructions: use a code in VBA to format a checkerboard.

Sub Checkerboard()
    ' For loop for formatting, through rows and columns.
    For i = 1 To 8
        For j = 1 To 8
            ' Checks if row is a non par number
            If i Mod 2 <> 0 Then
                ' Checks if column is a non par number.
                If j Mod 2 <> 0 Then
                    ' Cell is red if row is non par and column is non par
                    Cells(i, j).Interior.ColorIndex = 3
                Else
                    ' Cell is black if row is non par and column is par
                    Cells(i, j).Interior.ColorIndex = 1
                End If
            ' Checks if row is a par number
            ElseIf i Mod 2 = 0 Then
                If j Mod 2 <> 0 Then
                    ' Cell is black if row is par and column is non par
                    Cells(i, j).Interior.ColorIndex = 1
                Else
                    ' Cell is red if row is par and column is par
                    Cells(i, j).Interior.ColorIndex = 3
                End If
            End If
        Next j
    Next i
End Sub
