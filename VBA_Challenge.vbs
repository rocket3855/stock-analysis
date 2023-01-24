Sub DQAnalysis()
    Worksheets("DQ Analysis").Activate

    Range("A1").Value = "DAQO (Ticker: DQ)"

    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    Worksheets("2018").Activate

    'set initial volume to zero
    totalVolume = 0

    Dim startingPrice As Double
    Dim endingPrice As Double

    'Establish the number of rows to loop over
    rowStart = 2
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row

    'loop over all the rows
    For i = rowStart To rowEnd

        If Cells(i, 1).Value = "DQ" Then

            'increase totalVolume by the value in the current row
            totalVolume = totalVolume + Cells(i, 8).Value

        End If

        If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then

            startingPrice = Cells(i, 6).Value

        End If

        If Cells(i + 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then

            endingPrice = Cells(i, 6).Value

        End If

    Next i

    Worksheets("DQ Analysis").Activate
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume
    Cells(4, 3).Value = (endingPrice / startingPrice) - 1


End Sub

Sub AllStocksAnalysis()
    Dim startTime As Single
    Dim endTime As Single

    'Ask user for year to run analysis on
    'yearValueAnalysis <- this does not work here. Possibly yearValue is not in scope
    YearValue = InputBox("What year would you like to run the analysis on?")
    
    startTime = Timer
    
    '1) Format the output sheet on the "All Stocks Analysis" worksheet.
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + YearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    '2) Initialize an array of all tickers.
    Dim tickers(11) As String
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
 
    
    '3) Prepare for the analysis of tickers

        '3a) Initialize variables for the starting price and ending price.
        Dim startingPrice As Double
        Dim endingPrice As Double
        
        '3b) Activate the data worksheet.
        Worksheets(YearValue).Activate
        
        '3c)Find the number of rows to loop over.
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
        
    '4) Loop through the tickers.
    For i = 0 To 11
    
        ticker = tickers(i)
        totalVolume = 0
    
        '5) Loop through rows in the data.
        Worksheets(YearValue).Activate
        
        For j = 2 To RowCount
        
            '5a) Find the total volume for the current ticker.
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
            
            '5b) Find the starting price for the current ticker.
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
            End If
            
            '5c) Find the ending price for the current ticker.
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endingPrice = Cells(j, 6).Value
            End If
            
        Next j
    
        '6) Output the data for the current ticker.
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

    Next i
    
    'Call formatting routine so button shows formatting
    formatAllStocksAnalysisTable
    
    endTime = Timer
   
End Sub

Sub formatAllStocksAnalysisTable()

    'Formatting
    Worksheets("All Stocks Analysis").Activate
    
    'Format Header
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Font.TintAndShade = 1
    Range("A3:C3").Font.Color = vbBlue
    Range("A3:C3").Font.Size = 14
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    'Set RowCount to the number of rows to be formatted
    dataRowStart = 4
    dataRowEnd = Cells(Rows.Count, "A").End(xlUp).Row
    
    'Format Volume
    Range("B4:B" + Trim(Str(dataRowEnd))).NumberFormat = "#,##0"
    
    'Format Return
    Range("C4:C" + Trim(Str(dataRowEnd))).NumberFormat = "0.0%"
    
    'Format Column Width
    Columns("B").AutoFit
    
    'Conditionally format the color of Returns depending on their value
    For i = dataRowStart To dataRowEnd
    
        If Cells(i, 3) > 0 Then
        
            'Color the cell green
            Cells(i, 3).Interior.Color = vbGreen
            
        ElseIf Cells(4, 3) < 0 Then
        
            'Color the cell red
            Cells(i, 3).Interior.Color = vbRed
        
        Else
        
            'Clear the cell color
            Cells(i, 3).Interior.Color = xlNone
    
        End If
    
    Next i

End Sub

Sub ClearWorkSheet()

    Cells.Clear
    
End Sub

Sub yearValueAnalysis()

    YearValue = InputBox("What year would you like to run the analysis on?")

End Sub



