# stock-analysis
Public Sub AlphaStockAnalisys()

'Declare variables
Dim Open_Price, Clos_Price As Double
Dim Start, Total As Double

'Initiate values
Start = 2
Open_Price = Cells(2, "C")
Total = 0

'Columns titles'
Cells(1, "I") = "Ticker"
Cells(1, "J") = "Yearly Change"
Cells(1, "K") = "Percentage"
Cells(1, "L") = "Total Volume"

'Gow through all row and columns
For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row


'Setting conditions

    Ticker = Cells(i, "A")

    If Cells(i, "A") <> Cells(i + 1, "A") Then
        Cells(Start, "I") = Ticker
        Close_Price = Cells(i, "F")
        Total = Total + Cells(i, "G").Value
        Cells(Start, "J") = Close_Price - Open_Price
        If Cells(Start, "J") > 0 Then
            Cells(Start, "J").Interior.ColorIndex = 4
        Else
            Cells(Start, "J").Interior.ColorIndex = 3
        End If
        If Open_Price <> 0 Then
            Cells(Start, "K") = FormatPercent((Close_Price - Open_Price) / Open_Price, 2)
        Else
            Cells(Start, "K") = Null
        End If
        
        
        Cells(Start, "L") = Total
    
    
    
        Start = Start + 1
        Open_Price = Cells(i + 1, "C")
        Total = 0
    Else
        Total = Total + Cells(i, "G").Value
    End If
Next i
End Sub
