Attribute VB_Name = "Module1"
Sub stock_summary()

    'Setting the variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryRow As Long
    Dim j As Integer
    
    j = 0
    yearlyChange = 0
    percentChange = 0
    totalVolume = 0
    
    openingPrice = Range("C2")
    
    'column headers for stock analysis
    Cells(1, "I").Value = "Ticker"
    Cells(1, "J").Value = "Yearly Change"
    Cells(1, "K").Value = "Percent Change"
    Cells(1, "L").Value = "Total Stock Volume"

    
    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
    
        ' Find the last row in column A
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Initialize summary table starting row
        summaryRow = 2
        
        ' Loop through all rows in column A
        For i = 2 To lastRow
        
            ' Check if the ticker symbol has changed
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
                'total stock volume
                totalVolume = totalVolume + Cells(i, 7).Value
                
                    ' Get the ticker symbol
                    ticker = ws.Cells(i, 1).Value
                    
                    ' Get the closing price
                    closingPrice = ws.Cells(i, 6).Value
                    
                    ' Calculate the yearly change
                    yearlyChange = closingPrice - openingPrice
                    

                    ' Calculate the percentage change
                    percentChange = ((closingPrice - openingPrice) / openingPrice)
                    
                   
                    ' Output the information to the summary table
                    ws.Cells(j + 2, 9).Value = ticker
                    ws.Cells(j + 2, 10).Value = yearlyChange
                    ws.Cells(j + 2, 11).Value = percentChange
                    ws.Cells(j + 2, 12).Value = totalVolume
                    
                    ' Increment the summary table row
                    summaryRow = summaryRow + 1
                    'End If
                    
                'resetting back to zero
                yearlyChange = 0
                'percentChange = 0
                totalVolume = 0
                j = j + 1
                
                ElseIf ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                    
                    openingPrice = ws.Cells(i, 3).Value
                    
                    
            Else
               'If the ticker stays the same then add the results and print
                totalVolume = totalVolume + Cells(i, 7).Value
                
                
            End If
            
        Next i
    Next ws
End Sub

Sub ColorCodeYearlyChange()

    'Color added to column J to highlight a positive change with green and negative with red.
    Dim lastRow As Long
    Dim i As Long
    
    lastRow = Cells(Rows.Count, "J").End(xlUp).Row
    
    For i = 2 To lastRow
        If Cells(i, "J").Value > 0 Then
            Cells(i, "J").Interior.Color = RGB(0, 255, 0) ' Green
        ElseIf Cells(i, "J").Value < 0 Then
            Cells(i, "J").Interior.Color = RGB(255, 0, 0) ' Red
        End If
    Next i
End Sub


Sub CalculateStockStats()


   'Declaring variables
    Dim lastRow As Long
    Dim ticker As String
    Dim percentIncrease As Double
    Dim percentDecrease As Double
    Dim totalVolume As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    
    ' Find the last row in column A
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Loop through each row
    For i = 2 To lastRow
        ' Get the ticker symbol
        ticker = Cells(i, 1).Value
        
        ' Get the percent increase and decrease
        percentIncrease = Cells(i, 11).Value
        percentDecrease = Cells(i, 11).Value
        
        ' Get the total volume
        totalVolume = Cells(i, 12).Value
        
        ' Check if the current percent increase is greater than the greatest increase
        If percentIncrease > greatestIncrease Then
            greatestIncrease = percentIncrease
            greatestIncreaseTicker = ticker
        End If
        
        ' Check if the current percent decrease is less than the greatest decrease
        If percentDecrease < greatestDecrease Then
            greatestDecrease = percentDecrease
            greatestDecreaseTicker = ticker
        End If
        
        ' Check if the current total volume is greater than the greatest volume
        If totalVolume > greatestVolume Then
            greatestVolume = totalVolume
            greatestVolumeTicker = ticker
        End If
    Next i
    
    ' Output the results in columns O, P, and Q
    Range("O1").Value = "Greatest % Increase"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("P2").Value = greatestIncreaseTicker
    Range("Q2").Value = greatestIncrease
    
    Range("O3").Value = "Greatest % Decrease"
    Range("P3").Value = "Ticker"
    Range("Q3").Value = "Value"
    Range("P4").Value = greatestDecreaseTicker
    Range("Q4").Value = greatestDecrease
    
    Range("O5").Value = "Greatest Total Volume"
    Range("P5").Value = "Ticker"
    Range("Q5").Value = "Value"
    Range("P6").Value = greatestVolumeTicker
    Range("Q6").Value = greatestVolume
End Sub



