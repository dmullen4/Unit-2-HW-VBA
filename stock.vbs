Attribute VB_Name = "Module1"
Sub StockMarket()

    ' Sets an initial variable for holding the Ticker Symbol
    Dim Ticker_Symbol As String

    ' Sets an initial variable for holding the Total volume of stock per Ticker Symbol
    Dim Total_Volume As Double
    Total_Volume = 0

    ' Keeps track of the location for each Ticker Symbol in the Summary Table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    ' Sets variable for determining the Last Row
    Dim LastRow As Long
    
    ' Determines the Last Row of the worksheet
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

  
    ' Loops through all stock transactions
    For i = 2 To LastRow

        ' Checks ticker symbol to determine if it is the same as the previous or a new one
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        ' Sets the ticker symbol
        Ticker_Symbol = Cells(i, 1).Value

        ' Adds new values to Total_Volume of stock
        Total_Volume = Total_Volume + Cells(i, 7).Value

        ' Prints the Ticker Symbol to the Summary Table
        Range("I" & Summary_Table_Row).Value = Ticker_Symbol

        ' Prints the Total Volume of stock to the Summary Table
        Range("J" & Summary_Table_Row).Value = Total_Volume

        ' Adds one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
      
        ' Resets the Total_Volume of stock for new Ticker Symbol
        Total_Volume = 0

        ' Condition if the next row contains the same Ticker Symbol
        Else

        ' Add to the Total Volume of stock
        Total_Volume = Total_Volume + Cells(i, 7).Value

        End If

    Next i


End Sub


