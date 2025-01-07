Sub AnalyzeStocksTicker()
    Dim ws As Worksheet
    Dim Ticker As String
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim QuarterlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim SummaryRow As Long
    Dim LastRow As Long

    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    Dim GreatestIncreaseTicker As String
    Dim GreatestDecreaseTicker As String
    Dim GreatestVolumeTicker As String

    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate

        ' Set the variables
        SummaryRow = 2
        TotalVolume = 0
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestVolume = 0

        ' Find the last row of data
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' Add headers for the summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "% Change"
        ws.Cells(1, 12).Value = "Total Volume"

        ' Initialize the opening price for the first ticker
        OpeningPrice = ws.Cells(2, 3).Value

        ' Loop through all rows in the data
        For i = 2 To LastRow
            ' Get the volume
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value

            ' Check if its the last row of the ticker, and set ticker symbol
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Or i = LastRow Then
                Ticker = ws.Cells(i, 1).Value
                
                ' Get closing price
                ClosingPrice = ws.Cells(i, 6).Value

                ' Quarterly change and percent change
                QuarterlyChange = ClosingPrice - OpeningPrice
                If OpeningPrice <> 0 Then
                    PercentChange = QuarterlyChange / OpeningPrice
                Else
                    PercentChange = 0
                End If

                ' Place results in the summary table
                ws.Cells(SummaryRow, 9).Value = Ticker
                ws.Cells(SummaryRow, 10).Value = QuarterlyChange
                ws.Cells(SummaryRow, 11).Value = PercentChange
                ws.Cells(SummaryRow, 12).Value = TotalVolume

                ' Highlight the positive & negative changes
                If QuarterlyChange > 0 Then
                    ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4 
                ElseIf QuarterlyChange < 0 Then
                    ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3 
                Else
                    ws.Cells(SummaryRow, 10).Interior.ColorIndex = 2 
                End If


                ' Update the greatest increase, decrease, and volume
                If PercentChange > GreatestIncrease Then
                    GreatestIncrease = PercentChange
                    GreatestIncreaseTicker = Ticker
                End If

                If PercentChange < GreatestDecrease Then
                    GreatestDecrease = PercentChange
                    GreatestDecreaseTicker = Ticker
                End If

                If TotalVolume > GreatestVolume Then
                    GreatestVolume = TotalVolume
                    GreatestVolumeTicker = Ticker
                End If

                ' Start again for next ticker
                SummaryRow = SummaryRow + 1
                TotalVolume = 0
                If i <> LastRow Then
                    OpeningPrice = ws.Cells(i + 1, 3).Value
                End If
            End If
        Next i

        ' Update the greatest values
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"

        ws.Cells(2, 16).Value = GreatestIncreaseTicker
        ws.Cells(2, 17).Value = GreatestIncrease
        ws.Cells(3, 16).Value = GreatestDecreaseTicker
        ws.Cells(3, 17).Value = GreatestDecrease
        ws.Cells(4, 16).Value = GreatestVolumeTicker
        ws.Cells(4, 17).Value = GreatestVolume
    Next ws

    MsgBox "Analysis complete!"
End Sub




