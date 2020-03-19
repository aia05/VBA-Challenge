Sub Stocks()
    
    Dim ws As Worksheet

    For Each ws In Worksheets

        'Create column labels for the summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        Dim ticker As String
        Dim totalVolume As Double
        Dim rowCount As Long
        Dim startPrice As Double
        Dim endPrice As Double
        Dim yearlyChange As Double
        Dim percentChange As Double
        Dim lastRow As Long
 
        totalVolume = 0
        rowCount = 2
        startPrice = 0
        endPrice = 0
        yearlyChange = 0
        percentChange = 0
        
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To lastRow
            
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                startPrice = ws.Cells(i, 3).Value
            End If

            totalVolume = totalVolume + ws.Cells(i, 7)

            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

                ws.Cells(rowCount, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(rowCount, 12).Value = totalVolume

                endPrice = ws.Cells(i, 6).Value

                yearlyChange = endPrice - startPrice
                ws.Cells(rowCount, 10).Value = yearlyChange

                If yearlyChange >= 0 Then
                    With ws.Cells(rowCount, 10)  
                        .Interior.ColorIndex = 4
                        .Font.Color = vbGreen
                    End With
                    With ws.Cells(rowCount, 11) 
                        .Interior.ColorIndex = 4
                        .Font.Color = vbGreen
                    End With
                Else
                    With ws.Cells(rowCount, 10)  
                        .Interior.ColorIndex = 3
                        .Font.Color = vbRed
                    End With
                    With ws.Cells(rowCount, 11) 
                        .Interior.ColorIndex = 3
                        .Font.Color = vbRed
                    End With
                End If

                If startPrice = 0 Or endPrice = 0 Then
                    ws.Cells(rowCount, 11).Value = 0
                Else
                    percentChange = yearlyChange / startPrice
                    With ws.Cells(rowCount, 11)
                        .Value = percentChange
                        .NumberFormat = "0.00%"
                    End With
                End If

                rowCount = rowCount + 1

                totalVolume = 0
                startPrice = 0
                endPrice = 0
                yearlyChange = 0
                percentChange = 0
                
            End If
        Next i

        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
      
        last_row = ws.Cells(Rows.Count, 9).End(xlUp).Row

        Dim bestStock As String
        Dim worstStock As String
        Dim greatestChange As Double
        Dim worstChange As Double
        Dim stockVolume As String
        Dim mostVolume As Double

        greatestChange = ws.Range("K2").Value
        worstChange = ws.Range("K2").Value
        mostVolume = ws.Range("L2").Value

        For j = 2 To last_row
            
            If ws.Cells(j, 11).Value > best_percent_change Then
                bestChange = ws.Cells(j, 11).Value
                bestStock = ws.Cells(j, 9).Value
            End If

            If ws.Cells(j, 11).Value < worstChange Then
                worstChange = ws.Cells(j, 11).Value
                worstStock = ws.Cells(j, 9).Value
            End If

            If ws.Cells(j, 12).Value > most_volume_value Then
                mostVolume = ws.Cells(j, 12).Value
                stockVolume = ws.Cells(j, 9).Value
            End If

        Next j

        ws.Range("P2").Value = bestStock
        With ws.Range("Q2")
            .Value = bestChange
            .NumberFormat = "0.00%"
            .Font.Color = vbGreen
        End With
        
        ws.Range("P3").Value = worstStock
        With ws.Range("Q3")
            .Value = worstChange
            .NumberFormat = "0.00%"
            .Font.Color = vbRed
        End With
        
        ws.Range("P4").Value = stockVolume
        With ws.Range("Q4")
            .Value = mostVolume
            .NumberFormat = "#,###,##0"
        End With

        'autofit
        ws.Columns("I:Q").EntireColumn.AutoFit

    Next ws

End Sub
