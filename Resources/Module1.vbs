Attribute VB_Name = "Module1"
Sub GenerateReport()
    
    Dim ticker As String
    Dim yearOpen As Double
    Dim yearClose As Double
    Dim totalVolume As Double
    Dim lastRow As Double
    Dim writeRow As Double
    Dim greatestInc As Double
    Dim greatestDec As Double
    Dim greatestVol As Double
    Dim ws As Worksheet
    
    'Loop through worksheets
    For Each ws In Worksheets
    
        'Initialize variables
        ticker = ""
        yearOpen = 0
        yearClose = 0
        totalVolume = 0
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        writeRow = 2
        greatestInc = 0
        greatestDec = 0
        greatestVol = 0
    
        'Add summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
    
        'Add greatest changes table
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
    
        'Format cells
        ws.Columns("J").NumberFormat = "0.00"
        ws.Columns("K").NumberFormat = "0.00%"
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
    
        'Loop thru table
        For i = 2 To lastRow
        
            'Summarize results when ticker changes value on next row
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
                'Initialize ticker
                ticker = ws.Cells(i, 1).Value
            
                'Initialize yearOpen if stock has been zero for the entire year
                If ws.Cells(i, 3).Value = 0 Then
                    yearOpen = ws.Cells(i, 3).Value
                End If
            
                'Initialize yearClose
                yearClose = ws.Cells(i, 6).Value
                
                'Increase totalVolume for last value
                totalVolume = ws.Cells(i, 7).Value + totalVolume
            
                'Add ticker
                ws.Cells(writeRow, 9).Value = ws.Cells(i, 1).Value
        
                'Add yearly change
                ws.Cells(writeRow, 10).Value = yearClose - yearOpen
            
                'Format cell with yearly change
                If ws.Cells(writeRow, 10).Value >= 0 Then
            
                    ws.Cells(writeRow, 10).Interior.Color = RGB(0, 255, 0)
            
                Else
            
                    ws.Cells(writeRow, 10).Interior.Color = RGB(255, 0, 0)
                
                End If
            
                'Add yearly percent change, check for zero values
                If yearOpen = 0 And yearClose = 0 Then
            
                    ws.Cells(writeRow, 11).Value = 0
                
                Else
                
                    ws.Cells(writeRow, 11).Value = ws.Cells(writeRow, 10) / yearOpen
            
                End If
            
                'Add total stock volume
                ws.Cells(writeRow, 12).Value = totalVolume
            
                'Update greatest increase
                If greatestInc < ws.Cells(writeRow, 11).Value Then
            
                    greatestInc = ws.Cells(writeRow, 11).Value
                
                    ws.Range("P2").Value = ws.Cells(writeRow, 9).Value
                
                    ws.Range("Q2").Value = greatestInc
                
                End If
            
                'Update greatest decrease
                If greatestDec > ws.Cells(writeRow, 11).Value Then
            
                    greatestDec = ws.Cells(writeRow, 11).Value
                
                    ws.Range("P3").Value = ws.Cells(writeRow, 9).Value
                
                    ws.Range("Q3").Value = greatestDec
                
                End If
        
                'Update greatest total volume
                If greatestVol < ws.Cells(writeRow, 12).Value Then
            
                    greatestVol = ws.Cells(writeRow, 12).Value
            
                    ws.Range("P4").Value = ws.Cells(writeRow, 9).Value
                
                    ws.Range("Q4").Value = greatestVol
                
                End If
                                                 
            'Increase currentRow
            writeRow = writeRow + 1
            
            'Set totalVolume to zero
            totalVolume = 0
        
            Else
            
                'Check for start of new ticker and initialize yearOpen
                If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value And ws.Cells(i, 3) <> 0 Then
                
                    'Initialize yearOpen
                    yearOpen = ws.Cells(i, 3).Value
                
                End If
            
                'Increase totalVolume
                totalVolume = ws.Cells(i, 7).Value + totalVolume
            
            End If
        
        Next i
        
    Next
    
End Sub

