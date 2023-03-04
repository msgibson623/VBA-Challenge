Attribute VB_Name = "AlphabeticalTesting"
Sub AlphabeticalTesting()

'Loops through all the worksheets
    Dim ws As Worksheet
    For Each ws In Worksheets
    
      'Delcare Variables
        Dim ticker As String
        Dim yearly_change As Double
        Dim close_price As Double
        Dim open_price As Double
        Dim percentage As Double
        Dim ticker_row As Double
        Dim ticker_count As Double
        Dim ticker_percent As Double
        ticker_row = 2
        ticker_colume = 0
        ticker_percent = 0
    
      'Create/add column header
        Cells(1, "I").Value = "Ticker"
        
        Cells(1, "J").Value = "Yearly Change"
        
        Cells(1, "K").Value = "Percent Change"
        
        Cells(1, "L").Value = "Total Stock Volume"
        
       'Determine the Last Row
        lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Sets open price
        open_price = Cells(2, 3).Value
        
        'loops through all the rows start from row 2 to the last
        For i = 2 To lrow
        
        'calculates the total volume of tickers
        ticker_count = ticker_count + Cells(i, 7).Value
        
        'Verifies if the ticker matches
        If Cells(i + 1, 1) <> Cells(i, 1).Value Then
        
            'Reference cells
            ticker = Cells(i, 1).Value
         
            close_price = Cells(i, 6).Value
             
             'Calculate yearly change
             yearly_change = close_price - open_price
             
        'Calculates and formats the column to a percentage
        If open_price <> 0 And close_price <> 0 Then
        
            percentage = yearly_change / open_price
            Cells(ticker_row, "K").NumberFormat = "0.00%"
            
        Else
        
           'Reset
            ticker_percent = 0
            
        End If
        
             'Insert data via cells
            Cells(ticker_row, "I").Value = ticker
            
            Cells(ticker_row, "J").Value = yearly_change
            
            Cells(ticker_row, "L").Value = ticker_count
            
            Cells(ticker_row, "K").Value = percentage
            
        'Check if the yearly change is greater than 0
        If Cells(ticker_row, "J").Value > 0 Then
        
            Cells(ticker_row, "J").Interior.Color = RGB(124, 252, 0)
            
        'Check if the yearly change is less than or equal to 0
        ElseIf Cells(ticker_row, "J").Value <= 0 Then
     
            Cells(ticker_row, "J").Interior.Color = RGB(255, 0, 0)
           
        Else
            Cells(ticker_row, "J").Interior.Color = RGB(0, 0, 255)
            
    End If
            
            'Create the alternation
            ticker_row = ticker_row + 1
            
            'Reset
            ticker_count = 0
            
            open_price = Cells(i + 1, 3).Value
         
        End If

        Next i
      
        Next ws

End Sub
