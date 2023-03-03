Attribute VB_Name = "AlphabeticalTesting"
Sub AlphabeticalTesting()

      'Loops through all the worksheets
    Dim ws As Worksheet
    
        For Each ws In Worksheets
    
        ws.Activate
      
      'Delcare Variables
        Dim Ticker As String
        Dim Yearly_Change As Double
        Dim Close_Price As Double
        Dim Open_Price As Double
        Dim Percentage As Double
        Dim Ticker_Row As Double
        Dim Ticker_Volume As Double
        Dim Ticker_Percent As Double
        Ticker_Row = 2
        Ticker_Volume = 0
        Ticker_Percent = 0
    
      'Create/add column header
        Cells(1, "I").Value = "Ticker"
        
        Cells(1, "J").Value = "Yearly Change"
        
        Cells(1, "K").Value = "Percent Change"
        
        Cells(1, "L").Value = "Total Stock Volume"
        
       'Determine the Last Row
        lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Sets open price
        Open_Price = Cells(2, 3).Value
        
        'loops through all the rows start from row 2 to the last
        For i = 2 To lrow
        
            'calculates the total volume of tickers
            Ticker_Volume = Ticker_Volume + Cells(i, 7).Value
        
        'Verifies if the ticker matches
        If Cells(i + 1, 1) <> Cells(i, 1).Value Then
        
            'Reference cells
            Ticker = Cells(i, 1).Value
         
            Close_Price = Cells(i, 6).Value
             
             'Calculate yearly change
             Yearly_Change = Close_Price - Open_Price
             
        'Calculates and formats the column to a percentage
        If Open_Price <> 0 And Close_Price <> 0 Then
        
            Percentage = Yearly_Change / Open_Price
            Cells(Ticker_Row, "K").NumberFormat = "0.00%"
            
        Else
        
           'Reset
            Ticker_Percent = 0
            
        End If
        
             'Insert data via cells
            Cells(Ticker_Row, "I").Value = Ticker
            
            Cells(Ticker_Row, "J").Value = Yearly_Change
            
            Cells(Ticker_Row, "L").Value = Ticker_Volume
            
            Cells(Ticker_Row, "K").Value = Percentage
            
        'Check if the yearly change is greater than 0
        If Cells(Ticker_Row, "J").Value > 0 Then
        
            Cells(Ticker_Row, "J").Interior.Color = RGB(124, 252, 0)
            
        'Check if the yearly change is less than or equal to 0
        ElseIf Cells(Ticker_Row, "J").Value <= 0 Then
     
           Cells(Ticker_Row, "J").Interior.Color = RGB(255, 0, 0)
           
        Else
            Cells(Ticker_Row, "J").Interior.Color = RGB(0, 0, 255)
            
    End If
            
            'Create the alternation
            Ticker_Row = Ticker_Row + 1
            
            'Reset
            Ticker_Volume = 0
            
            Open_Price = Cells(i + 1, 3).Value
         
        End If

        Next i
      
    Next ws

End Sub
