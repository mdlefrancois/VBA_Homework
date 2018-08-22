Sub StockMrkt()
'Create a script that will loop through each year of stock data and grab the total amount of volume each stock had over the year.
'You will also need to display the ticker symbol to coincide with the total volume.

'<ticker>    <date>      <open>  <high>  <low>   <close> <vol>
'A           20150101    40.94   40.94   40.94   40.94   0

    Dim ws As Worksheet 'for each worksheet in a workbook
    Dim ws_num As Integer
    Dim Ticker As String
    Dim StkVol As Double
    Dim lRow As Long
    Dim TickerRow As Integer
       
    'count sheets in a workbook
    ws_num = ThisWorkbook.Worksheets.Count

    'Loop through each sheet and tabulate stock data
    For k = 1 To ws_num
    Worksheets(k).Activate
    StkVol = 0      'initialize for each sheet -- tabulate stock volume summation
    TickerRow = 2   'ititialize for each sheet -- tabulate unique stock name
    
    'within each worksheet place headers for stock tabulation
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Stock Total Volume"
        
    'Find the last non-blank cell in column A(1) for looping purposes
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
        
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
                        
                'Tabulate the stock volume summation
                Cells(TickerRow, 10).Value = StkVol
                        
                'update tabulation row
                TickerRow = TickerRow + 1
            
                'reset stock volume
                StkVol = 0#
                                
            End If
            
        Next i
        
    Next k
    
End Sub
