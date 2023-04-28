Sub Stock_Analysis()

Application.ScreenUpdating = False
      
Dim ws As Worksheet

'Loop through all worksheets
For Each ws In Worksheets

'Define Variables:

    Dim lastRow As Long
    
    Dim i As Long
    Dim Ticker As String
    
    Dim OpenPrice As Double
    OpenPrice = ws.Cells(2, 3).Value
    
    Dim ClosePrice As Double
    
    Dim YearlyChange As Double
    
    Dim PercentChange As Double
    
    Dim TSV As Double
    
    Dim Table As Integer        'To account for headers
    Table = 2
    

'Create Column Headers for Summary Table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'Find last row
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    'Assign Values
    TSV = 0 ' Reset value of total stock volume for each ws
    
    For i = 2 To lastRow
        
        'If Ticker Symbol is Different
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value       'Find Ticker
            
            'Determine Closing Price
            ClosePrice = ws.Cells(i, 6).Value
            
            'Calculate Yearly Change
            YearlyChange = ClosePrice - OpenPrice
            
            'Calculate Percentage Change
            PercentChange = YearlyChange / OpenPrice
            
           'If statement to prevent Yearly Change and Percentage Change divisible by 0
            If OpenPrice = 0 Then
                PercentChange = 0
            End If
            
            'Determine Opening Price for next Ticker Symbol
            OpenPrice = ws.Cells(i + 1, 3).Value
            
            
            'Calculate Total Stock Value for each Ticker Symbol
            TSV = TSV + ws.Cells(i, 7).Value        'Accumulate Stock Volume
            
            'Display Outcomes in Summary Table:
            ws.Range("I" & Table).Value = Ticker
            ws.Range("J" & Table).Value = YearlyChange
            ws.Range("K" & Table).Value = PercentChange
            ws.Range("L" & Table).Value = TSV
            
            'Format Conditionals:
            ws.Range("K" & Table).NumberFormat = "0.00%"    'Format as percentage
            
            If ws.Range("J" & Table).Value > 0 Then
                ws.Range("J" & Table).Interior.ColorIndex = 4   'For positive numbers = green
                
            Else:
                ws.Range("J" & Table).Interior.ColorIndex = 3   'For negative numbers = red
            
            End If
            
            'Reset Stock Volume for next Ticker Symbol
            TSV = 0
            
            'Autopopulate results into next following row
            Table = Table + 1
            
       
        Else:
            TSV = TSV + ws.Cells(i, 7).Value        'Add to Stock for values of same Ticker Symbol
            
        End If
        
    Next i
    
'Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
    
'Define Variables for Table 2
    Dim r As Range
    Set r = ws.Range("K2:K" & lastRow)
    
    Dim v As Range
    Set v = ws.Range("L2:L" & lastRow)
    
    Dim Max As Double
    Max = Application.WorksheetFunction.Max(r)  'Find max value of percentage
    
    Dim Min As Double
    Min = Application.WorksheetFunction.Min(r)  'Find min value of percentage
    
    Dim GI As String
    Dim GD As String
    Dim GTV As String
    
    Dim Volume As Double
    Volume = Application.WorksheetFunction.Max(v)   'Find greatest total volume
    
    Dim j As Long
    
'Create Column Headers for Table 2
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
      
'Loop to find Ticker Value (
    For j = 2 To lastRow
        
        'Greatest Total Increase %
        If ws.Cells(j, 11).Value = Max Then
        GI = ws.Cells(j, 9).Value
        End If
        
        'Greatest Total Decrease %
        If ws.Cells(j, 11).Value = Min Then
        GD = ws.Cells(j, 9).Value
        End If
        
        'Greatest Total Volume %
        If ws.Cells(j, 12).Value = Volume Then
        GTV = ws.Cells(j, 9).Value
        End If
    
    Next j

'Print Outcomes
   'Ticker
    ws.Range("O2").Value = GI
    ws.Range("O3").Value = GD
    ws.Range("O4").Value = GTV
    
    'Value
    ws.Range("P2").Value = Max
    ws.Range("P3").Value = Min
    ws.Range("P4").Value = Volume
    ws.Range("P2").NumberFormat = "0.00%"
    ws.Range("P3").NumberFormat = "0.00%"
    
'Autofit columns
    ws.Columns("I:Q").AutoFit
    
Next ws

End Sub

        
        

        
        
