' First Recap: Budget Checker
'
' Instructions: using skills learned during past exercises, solve
' the budget checker in this three-part exercise.
'
' Test only on e07.xlsm

' Part 1: Calculate the total after fees

Sub BudgetCheck()

    ' Part 1: Calculate the total after fees.
    ' Define and get values for variables
    Dim Total As Double
    Dim Price As Double
    Price = Range("F3").Value
    Dim Tax As Double
    Tax = Range("H3").Value
    
    'Calcultate total and print to desired cell
    Total = Price * (1 + Tax)
    MsgBox (Total)
    Range("L3").Value = Total

    ' Part 2: Alert in case it's over the budget
    ' Define and get values for variables
    Dim Budget As Double
    Budget = Range("B3").Value

    ' Conditionals for message box
    If Budget > Total Then
        MsgBox ("It's under budget")
    ElseIf Budget = Total Then
        MsgBox ("Exact!")
    Else
        MsgBox ("It's over budget")
    ' Part 3: Correct the price if overbudget
        ' Dim and get values for variables, calculate new price
        Dim Correct As Double
        Correct = Budget / (1 + Tax)
        ' Round to nearest whole dollar
        Correct = Application.WorksheetFunction.RoundDown(Correct, 0)
        ' Change the cells to match new price and total
        Range("F3").Value = Correct
        Range("L3").Value = Correct * (1 + Tax)
    End If


End Sub

