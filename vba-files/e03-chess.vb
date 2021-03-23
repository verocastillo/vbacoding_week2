' Cells and Ranges: Create a Chess Board.
'
' Instructions: create and fill the chess board with text-based
' pieces, using ranges and cells.

Sub ChessBoard()

' Using Ranges: Top Half
  
  ' Insert pieces:
  Range("A2:H2").Value = "Pawn"
  Range("A1, H1").Value = "Rook"
  Range("B1, G1").Value = "Knight"
  Range("C1, F1").Value = "Bishop"
  Range("D1").Value = "Queen"
  Range("E1").Value = "King"

' Using Cells: Bottom Half

  ' Insert Pieces
  Cells(8, 1).Value = "Rook"
  Cells(8, 8).Value = "Rook"
  Cells(8, 2).Value = "Knight"
  Cells(8, 7).Value = "Knight"
  Cells(8, 3).Value = "Bishop"
  Cells(8, 6).Value = "Bishop"
  Cells(8, 4).Value = "Queen"
  Cells(8, 5).Value = "King"

  ' BONUS: Use a for loop for efficiency!
  For i = 1 To 8
    Cells(7, i).Value = "Pawn"
  Next i

 ' Cell size:
  Range("A1:H8").RowHeight = 60
  Range("A1:H8").ColumnWidth = 20

  ' Color formatting:
    For j = 1 To 8
        For k = 1 To 8
            If i Mod 2 = 0 Then
                If j Mod 2 <> 0 Then
                Cells(j, k).Interior.ColorIndex = 1
                End If
            Else
                If j Mod 2 = 0 Then
                Cells(j, k).Interior.ColorIndex = 1
                End If
            End If
        Next k
    Next j

  ' Text formatting:
  Range("A1:H2").Font.ColorIndex = 3
  Range("A1:H2").Font.Bold = True

  Range("A7:H8").Font.ColorIndex = 5
  Range("A7:H8").Font.Bold = True

End Sub

