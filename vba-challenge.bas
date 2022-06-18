Attribute VB_Name = "Module1"
' VBA Homework: The VBA of Wall Street
Sub stockAnalysis()
    For Each ws In Worksheets
        ' Declare Variables
        Dim worksheetName As String
        Dim tickerSymbol As String
        Dim yearlyChange As Double
        Dim percentChange As Double
        Dim totalStockVolumn As Double
        
        ' Initialize Variables
        worksheetName = ws.Name
        tickerSymbol = ws.Cells(2, 1)
        yearlyChange = 0
        percentChange = 0
        totalStockVolumn = 0
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Column = 1
        analysisRow = 2
        openPrice = ws.Cells(2, 3)
        closePrice = 0
        
        ' Fills in first row of analysis
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volumn"
        
        ' Bonus: Static Rows and Columns
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volumn"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ' Loop to analyze and calculate data based off of ticker symbol
        For Row = 2 To lastRow
            ' Checks for when ticker symbol is different (IE. Not equal to current ticker symbol)
            If ws.Cells(Row + 1, Column).Value <> ws.Cells(Row, Column).Value Then
            ' If ticker is different
                ' Sets initial cell values
                closePrice = ws.Cells(Row, 6).Value
                tickerSymbol = ws.Cells(Row, Column).Value
                totalStockVolumn = totalStockVolumn + ws.Cells(Row, 7).Value
                
                ' Calculates values for analysis
                yearlyChange = closePrice - openPrice
                percentChange = yearlyChange / openPrice
                
                ' Fills in cells related to analysis
                ws.Range("I" & analysisRow).Value = tickerSymbol
                ws.Range("J" & analysisRow).Value = yearlyChange
                ws.Range("K" & analysisRow).Value = percentChange
                ws.Cells(analysisRow, 11).NumberFormat = "0.00%" ' Formats Percent Change column
                ws.Range("L" & analysisRow).Value = totalStockVolumn
                
                ' Adds Conditional Formatting
                ' Changes color of cells under Yearly Change column to Green if value is 0 or positive
                ws.Range("J" & analysisRow).FormatConditions.Add Type:=xlCellValue, Operator:=xlLessEqual, Formula1:="=0"
                ws.Range("J" & analysisRow).FormatConditions(1).Interior.Color = RGB(255, 0, 0)
                ' Changes color of cells under Yearly Change column to Green if value is negative
                ws.Range("J" & analysisRow).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
                ws.Range("J" & analysisRow).FormatConditions(2).Interior.Color = RGB(0, 255, 0)
                
                ' Resets values to next value as it relates to the next row that passes condition (IE. next ticker symbol)
                analysisRow = analysisRow + 1
                openPrice = ws.Cells(Row + 1, 3)
                totalStockVolumn = ws.Cells(Row + 1, 7).Value
            Else ' If ticker is the same
                totalStockVolumn = totalStockVolumn + ws.Cells(Row, 7).Value
            End If
        Next Row
    
        ' Bonus
        ' Initialize Variables
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolumn = 0
        greatestIncreaseTicker = ""
        greatestDecreaseTicker = ""
        greatestVolumnTicker = ""
        lastRowBonus = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        For Row = 2 To lastRowBonus
            currentPercentChange = ws.Cells(Row, 11).Value
            currentVolumn = ws.Cells(Row, 12).Value
            ' Looks for greatest % increase among values in "Percent Change" column
            If currentPercentChange > greatestIncrease Then
                greatestIncrease = currentPercentChange
                greatestIncreaseTicker = ws.Cells(Row, 9)
            End If
            
            ' Looks for greatest % decrease among values in "Percent Change" column
            If currentPercentChange < greatestDecrease Then
                greatestDecrease = currentPercentChange
                greatestDecreaseTicker = ws.Cells(Row, 9)
            End If
            
            ' Looks for greatest volumn among values in "Total Stock Volumn" column
            If currentVolumn > greatestVolumn Then
                greatestVolumn = currentVolumn
                greatestVolumnTicker = ws.Cells(Row, 9)
            End If
        Next Row
        
        ' Prints final values to worksheet
        ws.Range("P2").Value = greatestIncreaseTicker
        ws.Range("Q2").Value = greatestIncrease
        ws.Range("Q2").NumberFormat = "0.00%" ' Formats Greatest % Increase Cell
        ws.Range("P3").Value = greatestDecreaseTicker
        ws.Range("Q3").Value = greatestDecrease
        ws.Range("Q3").NumberFormat = "0.00%" ' Formats Greatest % Decrease Cell
        ws.Range("P4").Value = greatestVolumnTicker
        ws.Range("Q4").Value = greatestVolumn
        ws.Range("Q4").NumberFormat = "#"
        
        ' Additional formatting
        ws.Range("L:L").NumberFormat = "#"
        ws.Range("I:Q").Columns.AutoFit ' Automatically resizes column
    Next ws
End Sub
' Sub to clear data from worksheets
Sub clearData()
    For Each ws In Worksheets
        ws.Range("I:Q") = ""
        ws.Range("J:J").FormatConditions.Delete
    Next ws
End Sub

