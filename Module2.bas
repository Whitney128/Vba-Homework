Attribute VB_Name = "Module2"
Sub Alphabetical_testing()
'Declare variables for old and new
    Dim open_price As Double
    open_price = 0
    Dim close_price As Double
    close_price = 0
    Dim yearly_price As Double
    yearly_price = 0
    Dim yearly_percent_change As Double
    yearly_percent_change = 0
    Dim max_percent As Double
    max_percent = 0
    Dim min_percent As Double
    min_percent = 0
    Dim Summary_Table As Integer
    Summary_Table = 2
    Dim ticker As String
    Dim total_volume As Double
    Dim last_row As Long
    Dim max_ticker As String
    max_ticker = 0
    Dim min_ticker As String
    min_ticker = 0
    Dim max_volume As Double
    max_volume = 0
    Dim max_volume_ticker As String
    max_volume_ticker = " "
    
    
    
'Loop from the beginning of the worksheet(Row2) till its last row
    open_price = Cells(2, 3).Value
    ticker = Cells(2, 1).Value
    total_volume = 0
    max_volume = 0
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    For I = 2 To last_row
       
 'Set initial value of open price for ws
 'The rest tickers open price will be initialized within the for loop below
        'open_price = Cells(I, 3).Value
        
    
'Check if we are still within the same Ticker
        If Cells(I + 1, 1).Value = Cells(I, 1).Value Then
            total_volume = total_volume + Cells(I, 7).Value
'Calculate yearly_price and yearly_percent_change
            close_price = Cells(I, 6).Value
            yearly_price = close_price - open_price
'Check Division by 0 condition
'If open_price <> 0 Then
            If (ticker = "PLNT" And (yearly_price = 0 Or open_price = 0)) Then
                'MsgBox "PLNT"
               yearly_percent_change = 0
            Else
                yearly_percent_change = Round((yearly_price / open_price * 100), 2)
            End If
        Else
            
            
             If (yearly_percent_change > max_percent) Then
                max_percent = yearly_percent_change
                max_percent_ticker_name = ticker
            End If
            'If (ticker < max_ticker) Then
            '    max_ticker = ticker
            
            'End If
            If (yearly_percent_change < min_percent) Then
                min_percent = yearly_percent_change
                min_percent_ticker_name = ticker
            End If
         
            If (total_volume > max_volume) Then
                max_volume = total_volume
                max_volume_ticker_name = ticker
            End If
        
            
            ' MsgBox ("Next symbol")
                Range("I" & Summary_Table).Value = ticker
                Range("J" & Summary_Table).Value = yearly_price
            If yearly_price >= 0 Then
                Range("J" & Summary_Table).Interior.ColorIndex = 4
            Else
                Range("J" & Summary_Table).Interior.ColorIndex = 3
            End If
                Range("K" & Summary_Table).Value = yearly_percent_change
                Range("L" & Summary_Table).Value = total_volume
            open_price = Cells(I + 1, 3).Value
            ticker = Cells(I + 1, 1).Value
            Summary_Table = Summary_Table + 1
            total_volume = 0
        End If
        
        
    
    Next I
            Range("I1").Value = "Ticker"
            Range("J1").Value = "Yearly Change"
            Range("K1").Value = "Percent Change"
            Range("L1").Value = "Total Stock Volume"
            Range("O2").Value = "Greatest % Increase"
            Range("O3").Value = "Greatest % Decrease"
            Range("O4").Value = "Greatest Total Volume"
            Range("P1").Value = "Ticker"
            Range("Q1").Value = "Value"
            
            'Next ticker
           
    
        'This is the first, resulting worksheet, reset flag for the rest of worksheets
           
            'resetting counters
            yearly_percent_change = 0
            total_volume = 0
            
            'Else - If the cell immediately following a row is still the same ticker name,
            'just add to Totl Ticker Volume
            
            ' Increase the Total Ticker Volume
               ' Total_Ticker_Volume = Total_Ticker_Volume + Cells(2, 7).Value
         
                ' For debugging MsgBox (ws.Rows(i).Cells(2, 1))
            
            ' Check if it is not the first spreadsheet
            ' Record all new counts to the new summary table on the right of the current spreadsheet
            If Not Command_Spreadsheet Then
          
            Range("Q2").Value = ((max_percent) & "%")
            Range("Q3").Value = ((min_percent) & "%")
            Range("Q4").Value = max_volume
            Range("P2").Value = max_percent_ticker_name
            Range("P3").Value = min_percent_ticker_name
            Range("P4").Value = max_volume_ticker_name
            
            Else
                Command_Spreadsheet = False
            End If
            
            
            total_volume = 0
            yearly_price = 0
            I = 2
            
        'Next ws

End Sub



