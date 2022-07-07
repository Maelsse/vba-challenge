Attribute VB_Name = "Module1"

Sub MultiYearData()

    ' Begin Looping Through Worksheets
    For Each ws In Worksheets
    
    ' Declaring Variables
    Dim WorksheetName As String
    Dim j As Long
    Dim Percentage As Double
    Dim Ticker As Long
    Dim GPerDecr As Double
    Dim GPerIncr As Double
    Dim GTotalVol As Double
    
    
    'Creating Cells
    WorksheetName = ws.Name
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    ' Setting the Counts
    
    Ticker = 2
    j = 2
    
    LastrowA = Cells(Rows.Count, 1).End(xlUp).Row

    
    For i = 2 To Lastrow
    
    ' Begin our first check
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        ws.Cells(Ticker, 9).Value = ws.Cells(i, 1).Value
        ws.Cells(Ticker, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
        ' Conditional Formatting
        
        If ws.Cells(Ticker, 10).Value < 0 Then
        
        ws.Cells(Ticker, 10).Interior.ColorIndex = 3
        
        Else
        
        ws.Cells(Ticker, 10).Interior.ColorIndex = 4
        
        End If
        
        ' Percentage Change
        
        If ws.Cells(j, 3).Value <> 0 Then
        
        Percentage = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
        
        ws.Cells(Ticker, 11).Value = Format(Percentage, "Percent")
        
        Else
        
        ws.Cells(Ticker, 11).Value = Format(0, "Percent")
        
        End If
        
        ws.Cells(Ticker, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
        
        Ticker = Ticker + 1
        j = i + 1
        End If
    Next i
    
    ' Preparing for Summary Bonus
        
    LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    GPerIncr = ws.Cells(2, 11).Value
    GPerDecr = ws.Cells(2, 11).Value
    GTotalVol = ws.Cells(2, 12).Value
    
    For i = 2 To LastRowI
    
    ' Check if next Per value is larger
    
    If ws.Cells(i, 11).Value > GPerIncr Then
    GPerIncr = ws.Cells(i, 11).Value
    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
    
    Else
    
    GPerIncr = GPerIncr
    
    End If
    ' Checking for greatest decrease
    
    If ws.Cells(i, 11).Value < GPerDecr Then
    
    GPerDecr = ws.Cells(i, 11).Value
    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
    
    Else
    
    GPerDecr = GPerDecr
    
    End If
    
    'Checking for largest total
    
    If ws.Cells(i, 12).Value > GTotalVol Then
    GTotalVol = ws.Cells(i, 12).Value
    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
    
    Else
    GTotalVol = GTotalVol
    
    End If

    ws.Cells(2, 17).Value = Format(GPerIncr, "Percent")
    ws.Cells(3, 17).Value = Format(GPerDecr, "Percent")
    ws.Cells(4, 17).Value = Format(GTotalVol, "Scientific")
    
    Next i
    
        Worksheets(WorksheetName).Columns("A:Z").AutoFit
    Next ws
        


End Sub
