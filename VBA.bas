Attribute VB_Name = "Module1"
Sub StockChecker()

    For Each ws In Worksheets

    'Declare variables
    Dim lastrow As Long
    Dim currentstock As String
    Dim startvalue As Double
    Dim tickercount As Integer
    
    Dim increase As Double
    Dim decrease As Double
    Dim volume As Double
    Dim sheetname As String
    
    'ws code get name of worksheet
    sheetname = ws.Name
    
    
    'Find last row and skip header row and reset count of the ticker
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    startvalue = 2
    tickercount = 0
    
    'Print headers for output
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'Sort the data
    Worksheets(sheetname).Sort.SortFields.Clear
    With Worksheets(sheetname).Sort
        .SortFields.Add Key:=ws.Range("A1"), Order:=xlAscending
        .SortFields.Add Key:=ws.Range("B1"), Order:=xlAscending
        .SetRange Range(ws.Cells(1, 1), ws.Cells(lastrow, 7))
        .Header = xlYes
        .Apply
    End With
    
    'Start the loop
    For i = 2 To lastrow
    
        'Define the current stock ticker
        currentstock = ws.Cells(i, 1).Value
        
        If ws.Cells(i + 1, 1) <> currentstock Then
        
            tickercount = tickercount + 1
            
            'Store text as next ticker
            ws.Cells(tickercount + 1, 9).Value = currentstock
            
            'Calculate yearly change
            ws.Cells(tickercount + 1, 10).Value = ws.Cells(i, 6).Value - ws.Cells(startvalue, 3).Value
                
                If ws.Cells(tickercount + 1, 10).Value > 0 Then
                    ws.Cells(tickercount + 1, 10).Style = "40% - Accent3"
                    ws.Cells(tickercount + 1, 11).Style = "40% - Accent3"
                    
                ElseIf ws.Cells(tickercount + 1, 10).Value < 0 Then
                    ws.Cells(tickercount + 1, 10).Style = "40% - Accent2"
                    ws.Cells(tickercount + 1, 11).Style = "40% - Accent2"
                End If
        
            'Determine percentage change
            If ws.Cells(startvalue, 3).Value = 0 Then
                ws.Cells(tickercount + 1, 11).Value = FormatPercent(0)
                
            Else
                ws.Cells(tickercount + 1, 11).Value = FormatPercent((ws.Cells(i, 6).Value - ws.Cells(startvalue, 3).Value) / ws.Cells(startvalue, 3).Value)
            
            End If
            
            'stock volume
            ws.Cells(tickercount + 1, 12).Value = Format(WorksheetFunction.Sum(ws.Range(ws.Cells(startvalue, 7), ws.Cells(i, 7))), "#,###")
            
            'startvalue reset
            startvalue = i + 1
            
        End If
        
    Next i
    
    'Format columns J to L
    
    ws.Columns("J:L").EntireColumn.AutoFit
    With ws.Range("J:K")
        .VerticalAlignment = Excel.Constants.xlCenter
        .HorizontalAlignment = Excel.Constants.xlCenter
    End With
    
    'Print headers for analysis
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Yearly Change"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    'Sort for greatest percentage
    Worksheets(sheetname).Sort.SortFields.Clear
    With Worksheets(sheetname).Sort
        .SortFields.Add Key:=ws.Range("K1"), Order:=xlDescending
        .SetRange ws.Range(ws.Cells(1, 9), ws.Cells(lastrow, 12))
        .Header = xlYes
        .Apply
    End With
    
    'Print greatest Increase/decrease
    lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    ws.Range("P2").Value = ws.Cells(2, 9).Value
    ws.Range("Q2").Value = FormatPercent(ws.Cells(2, 11).Value)
    ws.Range("P3").Value = ws.Cells(lastrow, 9).Value
    ws.Range("Q3").Value = FormatPercent(ws.Cells(lastrow, 11).Value)
    
    'Sort to find Volume
    Worksheets(sheetname).Sort.SortFields.Clear
    With Worksheets(sheetname).Sort
        .SortFields.Add Key:=ws.Range("L1"), Order:=xlDescending
        .SetRange Range(ws.Cells(1, 9), ws.Cells(lastrow, 12))
        .Header = xlYes
        .Apply
    End With
    
    'Print Greatest Volume
    ws.Range("P4").Value = ws.Cells(2, 9).Value
    ws.Range("Q4").Value = Format(ws.Cells(2, 12).Value, "#, ###")
    
    'Format columns O to Q
    ws.Columns("O:Q").EntireColumn.AutoFit
    With ws.Range("P:Q")
        .VerticalAlignment = Excel.Constants.xlCenter
        .HorizontalAlignment = Excel.Constants.xlCenter
    End With
    
    'Sort to alphabetical
    Worksheets(sheetname).Sort.SortFields.Clear
    With Worksheets(sheetname).Sort
        .SortFields.Add Key:=ws.Range("I1"), Order:=xlAscending
        .SetRange Range(ws.Cells(1, 9), ws.Cells(lastrow, 12))
        .Header = xlYes
        .Apply
    End With
    
    'Change to the next worksheet
    Next ws
    
End Sub
