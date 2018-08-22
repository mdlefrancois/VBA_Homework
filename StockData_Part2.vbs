Sub StockMrkt_PII()
'Create a script that will loop through all the stocks and take the following info.

'Yearly change from what the stock opened the year at to what the closing price was.
'The percent change from the what it opened the year at to what it closed.
'The total Volume of the stock Ticker Symbol

    Dim ws As Worksheet 'for each worksheet in a workbook
    Dim ws_num As Integer
    Dim Ticker As String
    Dim StkVol As Double
    Dim lRow As Long
    Dim TickerRow As Integer
    Dim StockOpen As Double
    Dim StockClose As Double
    Dim YearChange As Double
    Dim PercentChange As Double
               
    'count sheets in a workbook
    ws_num = ThisWorkbook.Worksheets.Count

    'Loop through each sheet and tabulate stock data
    For k = 1 To ws_num
    Worksheets(k).Activate
    StkVol = 0      'initialize for each sheet -- tabulate stock volume summation
    TickerRow = 2   'ititialize for each sheet -- tabulate unique stock name
    
    'Pull initial opening price
    'Update loc below to keep tabs on all initial stock opening prices
    StockOpen = Cells(2, 3).Value

    
    'within each worksheet place headers for stock tabulation
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Worksheets(k).Columns.AutoFit 'adjust the column spacing for contents
        
    'Find the last non-blank cell in column A(1) for looping purposes
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
                                          
        'loop through all the stock information on a worksheet
        For i = 2 To lRow
        
            ' Sum the information if stock ticker symbol remains the same
            If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
                                         
                'Sum stock volume
                StkVol = StkVol + Cells(i, 7).Value
            
                'Ticker name
                Ticker = Cells(i, 1).Value
                                                                                          
            ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                        
                'Tabulate the ticker symbol
                Cells(TickerRow, 9).Value = Ticker
            
                'Add last stock volume before tabulating
                StkVol = StkVol + Cells(i, 7).Value
                Cells(TickerRow, 12).Value = StkVol
                                           
                'Pull the stock close price and update opening stock counter for next stock
                StockClose = Cells(i, 6).Value
                                         
                'Calculate yearly change
                YearChange = StockClose - StockOpen
                Cells(TickerRow, 10).Value = YearChange
                
                
                If YearChange > 0 Then
                    'Set the Font Color to Green
                    Cells(TickerRow, 10).Interior.ColorIndex = 4
                Else
                    'Set the cell to red
                    Cells(TickerRow, 10).Interior.ColorIndex = 3
                End If
                    
                
                'Calculate yearly change
                If StockOpen > 0 Then
                    PercentChange = ((StockClose - StockOpen) / StockOpen)
                Else
                    PercentChange = 0#
                End If
                    
                Cells(TickerRow, 11).Value = PercentChange
                Cells(TickerRow, 11).NumberFormat = "0.00%"
                
                'Update Open Stock Value
                StockOpen = Cells(i + 1, 3).Value
                
                'update tabulation row
                TickerRow = TickerRow + 1
                                
            End If
            
        Next i
        
    Next k
    
End Sub

