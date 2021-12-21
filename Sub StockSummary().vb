Sub StockSummary()
'Declare some variables
    Dim ticker As String
    Dim ticker_row As Integer
    Dim close1 As Double
    Dim open1 As Double
    Dim volume As Double
    Dim yearly_change As Double
    Dim percent_change As Double

For Each ws In Worksheets
    'Initial ticker symbol
    ticker = ws.Cells(2, 1).Value
    ticker_row = 2
    ws.Cells(ticker_row, 9).Value = ticker
    counter = 0
    
    'Initial opening price
    open1 = ws.Cells(2, 3).Value
    
    'Last row
    LastRow = ws.Cells(Rows.count, 1).End(xlUp).Row
    
    For i = 2 To LastRow + 1

    'Check if the next row is a new symbol
        If ws.Cells(i, 1).Value <> ticker Then
        
        ' Enter into the next line of summary chart
            ticker_row = ticker_row + 1
            
            ' Enter new symbol
            ticker = ws.Cells(i, 1).Value
            ws.Cells(ticker_row, 9).Value = ticker
            
            ' Collect closing price of previous ticker to find yearly change and percent change
            close1 = ws.Cells(i - 1, 6).Value
            yearly_change = close1 - open1
            
            'Avoid divide by zero
            If open1 = 1 Then
                percent_change = 0
                yearly_change = 0
            Else
                percent_change = yearly_change / open1
            End If
            
            ' Enter into summary
            ws.Cells(ticker_row - 1, 10).Value = yearly_change
            ws.Cells(ticker_row - 1, 11).Value = percent_change
            ws.Cells(ticker_row - 1, 11).NumberFormat = "0.00%"
            
            ' Format cells
            If ws.Cells(ticker_row - 1, 10).Value < 0 Then
            ws.Cells(ticker_row - 1, 10).Interior.ColorIndex = 3
            Else
            ws.Cells(ticker_row - 1, 10).Interior.ColorIndex = 4
            End If
                 
        
          ' Get new opening price
            Dim x As Integer
            x = 0
            If i <= LastRow Then
                Do While ws.Cells(i + x, 1).Value = ticker
                    If ws.Cells(i + x, 3).Value <> 0 Then
                        open1 = ws.Cells(i + x, 3).Value
                        Exit Do
                    Else
                    'This will catch if the opening price and closing price are both 0
                        open1 = 1
                        x = x + 1
                    End If
                Loop
            
           End If
           
        End If
        x = 0
    Next i

'Get the total stock volume

    'Reset ticker and summary row to ticker A
    ticker = ws.Cells(2, 1).Value
    ticker_row = 2
    
    
    For j = 2 To LastRow + 1
    
        'Check if we enter a new ticker
        If ws.Cells(j, 1) <> ticker Then
        
            'Input total volume into summary chart
            ws.Cells(ticker_row, 12).Value = volume
            
            'Get the new ticker
            ticker = ws.Cells(j, 1).Value
            
            'Set up the next row of summary chart
            ticker_row = ticker_row + 1
            
            'Reset the volume for the next ticker
            volume = ws.Cells(j, 7).Value
            
        'If it's the same ticker, keep adding up the volumes
        Else
            volume = volume + ws.Cells(j, 7).Value
        End If
        
    Next j
    
    'Bonus
    Dim increase As Double
    Dim decrease As Double
    Dim ticker_i As String
    Dim ticker_d As String
    Dim greatest_volume As Double
    Dim ticker_v As String
    
    'Set first comparison values
    increase = ws.Cells(2, 11).Value
    decrease = ws.Cells(2, 11).Value
   
   'Find greatest increase and decrease
    For k = 3 To 3169
        If ws.Cells(k, 11) > increase Then
            increase = ws.Cells(k, 11).Value
            ticker_i = ws.Cells(k, 9).Value
        ElseIf Cells(k, 11) < decrease Then
            decrease = ws.Cells(k, 11).Value
            ticker_d = ws.Cells(k, 9).Value
        End If
    Next k
    
    'FIll Cells
    ws.Range("P2").Value = ticker_i
    ws.Range("P3").Value = ticker_d
    ws.Range("Q2").Value = increase
    ws.Range("Q3").Value = decrease
      
    'Format Cells
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    
    'Find Greatest Volume
    greatest_volume = ws.Cells(2, 12).Value
    ticker_v = ws.Cells(2, 9).Value
    
    For l = 3 To 3169
        If ws.Cells(l, 12) > greatest_volume Then
            greatest_volume = ws.Cells(l, 12).Value
            ticker_v = ws.Cells(l, 9).Value
        End If
    Next l
    
   'Fill Cells
    ws.Range("P4").Value = ticker_v
    ws.Range("Q4").Value = greatest_volume
 
Next ws

End Sub
