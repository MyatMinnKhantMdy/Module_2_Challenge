Attribute VB_Name = "Module1"
Sub CalculateQuarterlyChanges()
    Dim ws As Worksheet
    Dim startRow As Long, endRow As Long
    Dim ticker As String
    Dim openPrice As Double, closePrice As Double, quarterlyChange As Double, percentChange As Double
    Dim totalVolume As Double
    Dim currentQuarter As Integer, lastQuarter As Integer
    Dim previousTicker As String
    Dim outputRow As Long
    
    ' Variables to track the greatest values
    Dim maxPercentIncrease As Double, maxPercentDecrease As Double, maxTotalVolume As Double
    Dim tickerMaxIncrease As String, tickerMaxDecrease As String, tickerMaxVolume As String
    Dim sheetMaxIncrease As String, sheetMaxDecrease As String, sheetMaxVolume As String
    
    ' Initialize tracking values
    maxPercentIncrease = -999999
    maxPercentDecrease = 999999
    maxTotalVolume = 0
    
    
   ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row of data in the current worksheet
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Add titles for the new columns
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Quarterly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

    outputRow = 2 ' Start output in row 2

    lastQuarter = -1 ' Initialize to an invalid quarter value to start
    previousTicker = "" ' Initialize previous ticker
    
    ' Loop through each row of data
    For i = 2 To lastRow ' Assuming header is in row 1
        ticker = ws.Cells(i, 1).Value
        currentQuarter = DatePart("q", ws.Cells(i, 2).Value)
        
        If currentQuarter <> lastQuarter Or ticker <> previousTicker Then
            ' New quarter or new ticker detected
            If lastQuarter <> -1 Then
                ' Calculate and store previous quarter's data
                quarterlyChange = closePrice - openPrice
                If openPrice <> 0 Then
                    percentChange = quarterlyChange / openPrice
                Else
                    percentChange = 0
                End If
                
                ' Write results
                ws.Cells(outputRow, 9).Value = previousTicker ' Write the correct ticker
                ws.Cells(outputRow, 10).Value = Format(quarterlyChange, "0.00") ' Quarterly Change
                ws.Cells(outputRow, 11).Value = Format(percentChange, "00.0%") ' Percent Change
                ws.Cells(outputRow, 12).Value = totalVolume ' Total Stock Volume
                outputRow = outputRow + 1
            End If
           
           ' Update maximums for increase, decrease, and volume
                    If percentChange > maxPercentIncrease Then
                        maxPercentIncrease = percentChange
                        tickerMaxIncrease = previousTicker
                       
                    End If
                    If percentChange < maxPercentDecrease Then
                        maxPercentDecrease = percentChange
                        tickerMaxDecrease = previousTicker
                       
                    End If
                    If totalVolume > maxTotalVolume Then
                        maxTotalVolume = totalVolume
                        tickerMaxVolume = previousTicker
                        
                    End If
            
  
            ' Initialize new quarter's data
            openPrice = ws.Cells(i, 3).Value ' Open price
            totalVolume = 0 ' Reset total volume
        End If

        ' Update data for the current row
        closePrice = ws.Cells(i, 6).Value ' Close price
        totalVolume = totalVolume + ws.Cells(i, 7).Value ' Volume
        
        ' Update last quarter and previous ticker
        lastQuarter = currentQuarter
        previousTicker = ticker
    Next i

    ' Write the last quarter's data after exiting the loop
    quarterlyChange = closePrice - openPrice
    If openPrice <> 0 Then
        percentChange = quarterlyChange / openPrice * 100
    Else
        percentChange = 0
    End If
    
    ' Write final results
    ws.Cells(outputRow, 9).Value = previousTicker
    ws.Cells(outputRow, 10).Value = Format(quarterlyChange, "0.00")
    ws.Cells(outputRow, 11).Value = Format(percentChange, "00.0%")
    ws.Cells(outputRow, 12).Value = totalVolume
    
    ' Check for final maximums
        If percentChange > maxPercentIncrease Then
            maxPercentIncrease = percentChange
            tickerMaxIncrease = previousTicker
            sheetMaxIncrease = ws.Name
        End If
        If percentChange < maxPercentDecrease Then
            maxPercentDecrease = percentChange
            tickerMaxDecrease = previousTicker
            sheetMaxDecrease = ws.Name
        End If
        If totalVolume > maxTotalVolume Then
            maxTotalVolume = totalVolume
            tickerMaxVolume = previousTicker
            sheetMaxVolume = ws.Name
        End If

       ' Apply conditional formatting to "Quarterly Change" column only
        With Range(ws.Cells(2, 10), ws.Cells(outputRow - 1, 10))
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
            .FormatConditions(1).Interior.Color = RGB(144, 238, 144) ' Light green for positive values
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
            .FormatConditions(2).Interior.Color = RGB(255, 182, 193) ' Light red for negative values
        End With
        
        
     ' Output summary of greatest values at specific positions
    Dim summarySheet As Worksheet
    Set summarySheet = ThisWorkbook.Sheets(1) ' Assuming output goes to the first sheet

    ' Output summary of greatest values at specific positions in each sheet
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ws.Cells(2, 16).Value = tickerMaxIncrease
        ws.Cells(3, 16).Value = tickerMaxDecrease
        ws.Cells(4, 16).Value = tickerMaxVolume
        
        ws.Cells(2, 17).Value = Format(maxPercentIncrease, "0.00%")
        ws.Cells(3, 17).Value = Format(maxPercentDecrease, "0.00%")
        ws.Cells(4, 17).Value = maxTotalVolume
    
    Next ws
End Sub


