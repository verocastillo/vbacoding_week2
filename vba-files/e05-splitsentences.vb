' Arrays and Splitting: Split the Sentence
'
' Instructions: Split the sentence according to the values
' in the worksheet: display the word associated with the
' position number.
'
' Test only on e05.xlsm

Sub SentenceSplit()

    ' Read and store the sentence from the corresponding cell
    Dim Sentence As String
    Sentence = Cells(1, 2).Value

    ' Read and store the position numbers for the words in the sentence
    Dim Word1 As Integer
    Dim Word2 As Integer
    Dim Word3 As Integer
    Word1 = Cells(4, 1).Value
    Word2 = Cells(5, 1).Value
    Word3 = Cells(6, 1).Value

    ' Split the sentence using space as a separator
    Dim Separate() As String
    Separate = Split(Sentence, " ")

    ' The position numbers are used to display the specific word in the
    ' corresponding cell. It starts at 0 so it's important to consider
    ' the offset.
    Cells(4, 2).Value = Separate(Word1 - 1)
    Cells(5, 2).Value = Separate(Word2 - 1)
    Cells(6, 2).Value = Separate(Word3 - 1)

End Sub
