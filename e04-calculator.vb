' Variables: Create a Calculator
'
' Instructions: create and dim variables appropriately, use cell
' ranges to input the values from the worksheet, calculate and display
' the result for the total in a message box.
'
' Test only on e04.xlsm

Sub ExVariables():

    ' Create and dim the variables 
    Dim Price As Double
    Dim Tax As Double
    Dim Quantity As Double
    Dim Total As Double

    ' Get and store the data values for each variable
    Price = Range("B2").Value
    Tax = Range("C2").Value
    Quantity = Range("D2").Value

    ' Calculate the total using the variables
    Total = Price * (1 + Tax) * Quantity

    ' Create a Message Box for the Total and insert into cell
    MsgBox ("Your total with tax is $" + Str(Total))
    Range("E2").Value = Total

End Sub