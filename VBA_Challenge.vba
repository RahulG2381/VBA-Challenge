Attribute VB_Name = "Module1"
    'The ticker symbol
    
    'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
    
    'The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
    
    'The total stock volume of the stock.
    
    Sub Ticker()
    
    'set an intial variable for holding the Ticker name
        Dim Ticker_Name As String
        Dim Openingprice As Double
        Dim Closingprice As Double
        Dim YearlyChange As Double
        Dim PercentageChange As Double
    
    ' Variables for tracking greatest % increase, % decrease, and total stock volume
        Dim GreatestIncrease As Double
        Dim GreatestDecrease As Double
        Dim GreatestTotalStockVolume As LongLong
        Dim Ticker_GreatestIncrease As String
        Dim Ticker_GreatestDecrease As String
        Dim Ticker_GreatestTotalStockVolume As String
    
        ' Set initial values for tracking variables
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestTotalStockVolume = 0
        Ticker_GreatestIncrease = ""
        Ticker_GreatestDecrease = ""
        Ticker_GreatestTotalStockVolume = ""
    
    
    'Set an initial value for YearlyChange
        YearlyChange = 0
    'Set an initial value for Percentage Change
        PercentageChange = 0
    'Set an intial variable for  holding the total Stock valume
        Dim Ticker_Totalstock As LongLong
    
        Ticker_Totalstock = 0
    
    'Keep track of the location for each Ticker symbol in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
    
        Dim ws As Worksheet
        Set ws = ActiveSheet
    'Find the last row of the active sheet
    
        lastrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    'loop through all ticker names
        For i = 2 To lastrow
    
    'check if we are still within th esame Ticker , if it is not..
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    'set the Ticker Name
        Ticker_Name = Cells(i, 1).Value
    
    'add to the Ticker Total stock Volume
    
        Ticker_Totalstock = Ticker_Totalstock + Cells(i, 7).Value

    'Set the closing price
        Closingprice = Cells(i, 6).Value
 
     'Calculate Yearlychange
        YearlyChange = Closingprice - Openingprice
     
     'Calculate Percentage Change
        If Openingprice <> 0 Then
        PercentageChange = (YearlyChange / Openingprice) * 100
        Else
        PercentageChange = 0
        End If
    
    'print the Ticker Name in the summary Table
        Range("J" & Summary_Table_Row).Value = Ticker_Name
    
    'Print the Yearly Change to the summary Table
        Range("K" & Summary_Table_Row).Value = YearlyChange
    
    'Print the Percentage Change to the summary Table
        Range("L" & Summary_Table_Row).Value = Format(PercentageChange, "0.00") & "%"
    
    
    'Apply conditional formatting based on YearlyChange
        If YearlyChange > 0 Then
        Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
        ElseIf YearlyChange < 0 Then
        Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
        End If
    
    'Apply conditional formatting based on PercentageChange
        If PercentageChange > 0 Then
        Range("L" & Summary_Table_Row).Interior.ColorIndex = 4
        ElseIf PercentageChange < 0 Then
        Range("L" & Summary_Table_Row).Interior.ColorIndex = 3
        End If
    
    'print the Ticker Amount to the summary Table
    
        Range("M" & Summary_Table_Row).Value = Ticker_Totalstock
    
    'Check for greatest % increase, % decrease, and total stock volume
       If PercentageChange > GreatestIncrease Then
       GreatestIncrease = PercentageChange
       Ticker_GreatestIncrease = Ticker_Name
       End If
                
       If PercentageChange < GreatestDecrease Then
       GreatestDecrease = PercentageChange
       Ticker_GreatestDecrease = Ticker_Name
       End If
    
       If Ticker_Totalstock > GreatestTotalStockVolume Then
       GreatestTotalStockVolume = Ticker_Totalstock
       Ticker_GreatestTotalStockVolume = Ticker_Name
       End If
    
    
    'Add one to the summary table row
       Summary_Table_Row = Summary_Table_Row + 1
    
    'Reset the Ticker Total stock
    
       Ticker_Totalstock = 0
    
    'Reset the Opening price for the next ticker
       Openingprice = 0
    
    'if the cell immidiately following a row is the same ticker..
        Else
    'Add to the Tikcer Total stock
        Ticker_Totalstock = Ticker_Totalstock + Cells(i, 7).Value
    
    'If it's teh first instance of the ticker, set the Opening price
        If Openingprice = 0 Then
        Openingprice = Cells(i, 3).Value
    
    End If
    
    End If
    
    Next i
    
     'Print Greatest % Increase, Greatest % Decrease,and Greatest Total Stock Volume and their respective Ticker names in Table
     
        Range("R2").Value = Format(GreatestIncrease, "0.00") & "%"
        Range("Q2").Value = Ticker_GreatestIncrease
    
        Range("R3").Value = Format(GreatestDecrease, "0.00") & "%"
        Range("Q3").Value = Ticker_GreatestDecrease
    
        Range("R4").Value = GreatestTotalStockVolume
        Range("Q4").Value = Ticker_GreatestTotalStockVolume
    
    End Sub
    

