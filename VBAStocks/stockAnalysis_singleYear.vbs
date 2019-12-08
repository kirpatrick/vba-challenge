Sub stockAnalysis_singleYear()

    MsgBox ("Analyzing data from this sheet...")

    ' Declare variables to time execution
    Dim beginTimeSec As Single
    Dim endTimeSec As Single
    
    'Start timer
    beginTimeSec = Timer()

    '***** Declare & initialize the variables *****

    'Stock ticker name
    Dim ticker As String

    'Hold year begin and end price for a stock
    Dim yearBeginPrice As Double
    Dim yearEndPrice As Double
        
    'Hold total annual volume for a stock
    Dim totalVolume As Variant

    'Hold tickerIndex to check for new ticker symbol during data read/write
    Dim tickerIndex As Integer

    'Will be updated to the record count for the sheet
    Dim recordCount As Long
    
    'Hold greatestPercentIncrease variables
    Dim greatestPercentIncrease As Double
    Dim greatestPercentIncreaseTicker As String
    
    'Hold greatestPercentDecrease variables
    Dim greatestPercentDecrease As Double
    Dim greatestPercentDecreaseTicker As String
    
    'Hold greatestTotalVolume variables
    Dim greatestTotalVolume As Variant
    Dim greatestTotalVolumeTicker As String

    ' ***** initialize all variables *****
    ticker = Range("A2").Value
    yearBeginPrice = Range("C2").Value
    yearEndPrice = 0
    totalVolume = Range("G2").Value
    tickerIndex = 2
    greatestPercentIncrease = 0
    greatestPercentDecrease = 0
    greatestTotalVolume = 0
    recordCount = ActiveSheet.UsedRange.Rows.Count

    ' ***** Print column headers for output *****
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"

    ' ***** Begin analysis *****

    ' Process all records starting with 3rd row...
    ' Accounting for headers and initalization
    For i = 3 To recordCount

        ' Check for new ticker
        ' If yes
        If (ticker <> Cells(i, 1).Value) Or (i = recordCount) Then

            ' Capture data for previous ticker and format as needed
            
                ' ***** Ticker *****
                Cells(tickerIndex, 9).Value = ticker
                
                ' ***** Yearly Change ****
                Cells(tickerIndex, 10).Value = yearEndPrice - yearBeginPrice
                
                ' Format Yearly Change output to 2 decimal places
                Cells(tickerIndex, 10).NumberFormat = "0.00"
                
                'Highlight negative Yearly Change red
                If Cells(tickerIndex, 10).Value < 0 Then
                    Cells(tickerIndex, 10).Interior.ColorIndex = 3
                    
                'Highlight positive Yearly Change green
                ElseIf Cells(tickerIndex, 10).Value >= 0 Then
                    Cells(tickerIndex, 10).Interior.ColorIndex = 10
                End If
                
                ' ***** Percent Change *****
                'Account for possible divide by 0
                If yearBeginPrice <> 0 Then
                    Cells(tickerIndex, 11).Value = (yearEndPrice - yearBeginPrice) / yearBeginPrice
                Else
                    Cells(tickerIndex, 11).Value = 0
                End If
                
                ' Format Yearly Percent Change output to percentage with 2 decimal places
                Cells(tickerIndex, 11).NumberFormat = "0.00%"
                
                'Check & set greatest % increase
                If Cells(tickerIndex, 11).Value > greatestPercentIncrease Then
                    greatestPercentIncrease = Cells(tickerIndex, 11).Value
                    greatestPercentIncreaseTicker = ticker
                End If
                
                'Check & set greatest % decrease
                If Cells(tickerIndex, 11).Value < greatestPercentDecrease Then
                    greatestPercentDecrease = Cells(tickerIndex, 11).Value
                    greatestPercentDecreaseTicker = ticker
                End If
                
                ' ***** Total Stock Volume *****
                Cells(tickerIndex, 12).Value = totalVolume
                
                'Check & set greatest TotalVolume
                If Cells(tickerIndex, 12).Value > greatestTotalVolume Then
                    greatestTotalVolume = Cells(tickerIndex, 12).Value
                    greatestTotalVolumeTicker = ticker
                End If
            
            ' Start data capture for new ticker
            
                ' Retain next output row index for new ticker data
                tickerIndex = tickerIndex + 1
                
                ' Store the next ticker variable
                ticker = Cells(i, 1).Value
                
                ' Store the beginning year stock price
                yearBeginPrice = Cells(i, 3).Value
                
                ' Start tracking the year end price
                yearEndPrice = Cells(i, 6).Value
                
                ' Start tracking the totalVolume
                totalVolume = Cells(i, 7).Value

        ' If no change to ticker
        Else
            ' Update the year end price to the most current value checked
            yearEndPrice = Cells(i, 6).Value
            
            ' Add the current total stock volume to the running total
            totalVolume = totalVolume + Cells(i, 7).Value

        End If

    Next i
    
    ' If all records have been viewed, update the year end analysis table

    'Print the greatest percent increase data for the sheet
    Range("P2").Value = greatestPercentIncreaseTicker
    Range("Q2").Value = greatestPercentIncrease
    Range("Q2").NumberFormat = "0.00%"
    
    'Print the greatest percent decrease data for the sheet
    Range("P3").Value = greatestPercentDecreaseTicker
    Range("Q3").Value = greatestPercentDecrease
    Range("Q3").NumberFormat = "0.00%"
    
    'Print the greatest total volume data for the sheet
    Range("P4").Value = greatestTotalVolumeTicker
    Range("Q4").Value = greatestTotalVolume

    'Stop timer
    endTimeSec = Timer()
    
    ' Confirm script completion
    MsgBox ("Analysis Complete!")

    'Display runtime in seconds
    MsgBox ("Runtime:  " & endTimeSec - beginTimeSec & " seconds")
    
End Sub