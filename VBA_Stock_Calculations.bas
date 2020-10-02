Attribute VB_Name = "Module2"
Sub VBA_Stock_Calculations()

' time to define variables

' ticker symbol
Dim ticker As String

' number of tickers per sheet
Dim number_tickers As String

' keeps track of last row in each worksheet
Dim lastrow As Long

'what we are finding out from the worksheet
Dim opening_price As Double
Dim closing_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock_volume As Double
Dim greatest_percent_increase As Double
Dim greatest_percent_decrease As Double
Dim greatest_stock_volume As Double

'variables keep track of ticker that has greatest percent/volume increase or decreease for bonus
Dim greatest_percent_increase_ticker As String
Dim greatest_percent_decrease_ticker As String
Dim greatest_stock_volume_ticker As String

'loop into various sheets within workbook
For Each ws In Worksheets
    'activates worksheet
    ws.Activate
    
    'find last row
    lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    'add column headers to rows I through L
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'initialize variables for each worksheet
    ticker = ""
    number_tickers = 0
    opening_price = 0
    yearly_change = 0
    percent_change = 0
    total_stock_volume = 0
    
    'Nested for loop, loop through list of tickers, i is 1 to skip header -> For loop
    For i = 2 To lastrow
        'return our ticker value
        ticker = Cells(i, 3).Value
        
        'returns the opening price for that ticker
        If opening_price = 0 Then
            opening_price = Cells(i, 3).Value
        End If
        
        'Adds total stock volumes for ticker
        total_stock_volume = total_stock_volume + Cells(i, 7).Value
        
        'Checks if we get a different ticker (<> is same as =! in other languages) -> Conditional
        If Cells(i + 1, 1).Value <> ticker Then
            'When we get to different ticker, change the number of total tickers
            number_tickers = number_tickers + 1
            Cells(number_tickers + 1, 9) = ticker
            
            'Get end of year closing_price, yearly change for ticker
            closing_price = Cells(i, 6)
            yearly_change = closing_price - opening_price
            
            'Add yearly change values to correct column in spreadsheet
            Cells(number_tickers + 1, 10).Value = yearly_change
            
            'Set up conditional formatting for yearly_change columnb(red for negative, green for positive)
            If yearly_change > 0 Then
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 4
            ElseIf yearly_change < 0 Then
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 3
            'If yearly change value is 0, make it yellow
            Else
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 6
            End If
            
            
            
            'Calculate percent change value for ticker
            If opening_price = 0 Then
                percent_change = 0
            Else
                percent_change = (yearly_change / opening_price)
            End If
            
            
            
            'Format as percent and place in col 11 -> hw does not mention that this needs to have conditional formatting
            Cells(number_tickers + 1, 11).Value = Format(percent_change, "Percent")
            
            'Set opening price back to 0 when we get to different ticker
            opening_price = 0
            
            'Add total volume to column 12
            Cells(number_tickers + 1, 12).Value = total_stock_volume
            
            'set it back to 0 for the next ticker in the list
            total_stock_volume = 0
        End If
        
    Next i
    
        'BONUS SECTION -- displays greatest % increase and decrease, greatest total volume for each year, will go in column O, also adding ticker and value
        Range("O3").Value = "Greatest Percent Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
         ' Get the last row
        lastRowState = ws.Cells(Rows.Count, "I").End(xlUp).Row
        
        ' have to initialize variables, set them to first row
        greatest_percent_increase = Cells(2, 11).Value
        greatest_percent_increase_ticker = Cells(2, 9).Value
        greatest_percent_decrease = Cells(2, 11).Value
        greatest_percent_decrease_ticker = Cells(2, 9).Value
        greatest_stock_volume = Cells(2, 12).Value
        greatest_stock_volume_ticker = Cells(2, 9).Value
        
        
        ' loops through list of tickers, starts after header
        For i = 2 To lastRowState
        
            ' Find the greatest percent increase
            If Cells(i, 11).Value > greatest_percent_increase Then
                greatest_percent_increase = Cells(i, 11).Value
                greatest_percent_increase_ticker = Cells(i, 9).Value
            End If
            
            ' Find the greatest percent decrease
            If Cells(i, 11).Value < greatest_percent_decrease Then
                greatest_percent_decrease = Cells(i, 11).Value
                greatest_percent_decrease_ticker = Cells(i, 9).Value
            End If
            
            ' Find the greatest stock volume
            If Cells(i, 12).Value > greatest_stock_volume Then
                greatest_stock_volume = Cells(i, 12).Value
                greatest_stock_volume_ticker = Cells(i, 9).Value
            End If
            
        Next i
        
        ' Add the values we elucidated into the worksheet
        Range("P2").Value = Format(greatest_percent_increase_ticker, "Percent")
        Range("Q2").Value = Format(greatest_percent_increase, "Percent")
        Range("P3").Value = Format(greatest_percent_decrease_ticker, "Percent")
        Range("Q3").Value = Format(greatest_percent_decrease, "Percent")
        Range("P4").Value = greatest_stock_volume_ticker
        Range("Q4").Value = greatest_stock_volume
        
Next ws
    
    
End Sub


