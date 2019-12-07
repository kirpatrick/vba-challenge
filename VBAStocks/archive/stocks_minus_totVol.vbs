Sub stocks()
    
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

    'Iterate through all sheets in workbook
    For j = 1 To Sheets.Count
    
        ' ***** initialize all variables *****
        ticker = Sheets(j).Range("A2").Value
        yearBeginPrice = Sheets(j).Range("C2").Value
        yearEndPrice = 0
        totalVolume = Sheets(j).Range("G2").Value
        tickerIndex = 2
        greatestPercentIncrease = 0
        greatestPercentDecrease = 0
        recordCount = ActiveSheet.UsedRange.Rows.Count
    
        ' ***** Print column headers for output *****
        Sheets(j).Range("I1").Value = "Ticker"
        Sheets(j).Range("J1").Value = "Yearly Change"
        Sheets(j).Range("K1").Value = "Percent Change"
        Sheets(j).Range("L1").Value = "Total Stock Volume"
    
        Sheets(j).Range("O2").Value = "Greatest % increase"
        Sheets(j).Range("O3").Value = "Greatest % Decrease"
        Sheets(j).Range("P1").Value = "Ticker"
        Sheets(j).Range("Q1").Value = "Value"
    
        ' ***** Begin analysis *****
    
        ' Process all records starting with 3rd row...
        ' Accounting for headers and initialization
        For i = 3 To recordCount
    
            ' Check for new ticker
            ' If yes or at the end of records for sheet
            If (ticker <> Sheets(j).Cells(i, 1).Value) Or (i = recordCount) Then
    
                ' Capture data for previous ticker and format as needed
                
                    ' ***** Ticker *****
                    Sheets(j).Cells(tickerIndex, 9).Value = ticker
                    
                    ' ***** Yearly Change ****
                    Sheets(j).Cells(tickerIndex, 10).Value = yearEndPrice - yearBeginPrice
                    
                    ' Format Yearly Change output to 2 decimal places
                    Sheets(j).Cells(tickerIndex, 10).NumberFormat = "0.00"
                    
                    'Highlight negative Yearly Change red
                    If Sheets(j).Cells(tickerIndex, 10).Value < 0 Then
                        Sheets(j).Cells(tickerIndex, 10).Interior.ColorIndex = 3
                        
                    'Highlight positive Yearly Change green
                    ElseIf Sheets(j).Cells(tickerIndex, 10).Value >= 0 Then
                        Sheets(j).Cells(tickerIndex, 10).Interior.ColorIndex = 10
                    End If
                    
                    ' ***** Percent Change *****
					'Account for possible divide by 0
                    If yearBeginPrice <> 0 Then
                        Sheets(j).Cells(tickerIndex, 11).Value = (yearEndPrice - yearBeginPrice) / yearBeginPrice
                    Else
                        Sheets(j).Cells(tickerIndex, 11).Value = 0
                    End If
                    
                    ' Format Yearly Percent Change output to percentage with 2 decimal places
                    Sheets(j).Cells(tickerIndex, 11).NumberFormat = "0.00%"
                    
                    'Check & set greatest % increase variables
                    If Sheets(j).Cells(tickerIndex, 11).Value > greatestPercentIncrease Then
                        greatestPercentIncrease = Sheets(j).Cells(tickerIndex, 11).Value
                        greatestPercentIncreaseTicker = ticker
                    End If
                    
                    'Check & set greatest % decrease variables
                    If Sheets(j).Cells(tickerIndex, 11).Value < greatestPercentDecrease Then
                        greatestPercentDecrease = Sheets(j).Cells(tickerIndex, 11).Value
                        greatestPercentDecreaseTicker = ticker
                    End If
                    
                    ' ***** Total Stock Volume *****
                    Sheets(j).Cells(tickerIndex, 12).Value = totalVolume
                
                ' Start data capture for new ticker
                
                    ' Retain next output row index for new ticker data
                    tickerIndex = tickerIndex + 1
                    
                    ' Store the next ticker variable
                    ticker = Sheets(j).Cells(i, 1).Value
                    
                    ' Store the beginning year stock price
                    yearBeginPrice = Sheets(j).Cells(i, 3).Value
                    
                    ' Start tracking the year end price
                    yearEndPrice = Sheets(j).Cells(i, 6).Value
                    
                    ' Start tracking the totalVolume
                    totalVolume = Sheets(j).Cells(i, 7).Value
    
            ' If no change to ticker
            Else
                ' Update the year end price to the most current value checked
                yearEndPrice = Sheets(j).Cells(i, 6).Value
                
                ' Add the current total stock volume to the running total
                totalVolume = totalVolume + Sheets(j).Cells(i, 7).Value
    
            End If
    
        Next i
        
		'Print the greatest percent increase data for the sheet
        Sheets(j).Range("P2").Value = greatestPercentIncreaseTicker
        Sheets(j).Range("Q2").Value = greatestPercentIncrease
        Sheets(j).Range("Q2").NumberFormat = "0.00%"
        
		'Print the greatest percent decrease data for the sheet
        Sheets(j).Range("P3").Value = greatestPercentDecreaseTicker
        Sheets(j).Range("Q3").Value = greatestPercentDecrease
        Sheets(j).Range("Q3").NumberFormat = "0.00%"
    
    'Process next sheet
    Next j
    
	'Stop timer
    endTimeSec = Timer()
    
	'Display runtime in seconds
    MsgBox ("Runtime:  " & endTimeSec - beginTimeSec & " seconds")
    
End Sub