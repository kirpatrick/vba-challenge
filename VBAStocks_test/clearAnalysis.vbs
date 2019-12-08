Sub clearAnalysis()

    ' Confirm script execution
    MsgBox ("Clearing analysis...")

    ' Iterate through all data sheets in this workbook
    For i = 2 To Sheets.Count
    
        ' For each sheet
        For j = 1 To 9
            ' Delete the data in columns 9-18
            Sheets(i).Columns(9).Delete
        Next j
        
    Next i
    
    ' Confirm script completion
    MsgBox ("Clear Complete!")

End Sub