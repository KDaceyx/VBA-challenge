Attribute VB_Name = "Module1"
Sub Test_Ticker()

'Loop through all sheets

For Each ws In Worksheets

    'Set an initial variable for holding the Ticker name
    Dim Ticker As String
    
    'set an initial variable for holding the total volume
    Dim Total_Volume As Double
    Total_Volume = 0
    
    'Set an initial variable for holding the price change
    Dim Price_Change As Double
    Price_Change = 0
    
    'Set a variable for holding the percentage change in stock value
    Dim Percent_Change As Double
    Percent_Change = 0
    
    'Set a variable for holding initial price
    Dim Opening_Price As Double
    Opening_Price = 0
       
    'Set a variable for ending price
    Dim Ending_Price As Double
    Ending_Price = 0
    
    'Keep track of the location for each Ticker symbol in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    'Set a variable for the row number and end row
    Dim i As Long
    Dim EndRow As Long
    EndRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Create the headers for the new columns
    ws.Cells(1, 8).Value = "Ticker"
    ws.Cells(1, 9).Value = "Change in Value"
    ws.Cells(1, 10).Value = "Percent Change"
    ws.Cells(1, 11).Value = "Total Stock Volume"
    ws.Range("H1:K1").Columns.AutoFit
    
    'Set initial location of opening price of each stock
    Opening_Price = ws.Cells(2, 3)
    
        'loop through all stock data
        For i = 2 To EndRow
        
            'Check if we are still within the same ticker symbol. if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Set the ticker symbol
                Ticker = ws.Cells(i, 1).Value
                
                'Add to the stock volume total
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                
                'Find ending stock price
                Ending_Price = ws.Cells(i, 6).Value
                           
                'Subtract initial stock value from ending stock value
                Price_Change = Ending_Price - Opening_Price
                
                'Calculate percentage change from Price_Change and ending price
                    If Opening_Price <> 0 Then
                        Percent_Change = (Price_Change / Opening_Price) * 100
                    End If
                        
                'Print the ticker in the summary table
                ws.Range("H" & Summary_Table_Row).Value = Ticker
                
                'Print the total stock volume to the summary table
                ws.Range("K" & Summary_Table_Row).Value = Total_Volume
                
                'Print the change in stock price to the summary table
                ws.Range("I" & Summary_Table_Row).Value = Price_Change
                
                    'Add color formatting to summary table for price change column
                    If Price_Change > 0 Then
                        ws.Range("I" & Summary_Table_Row).Interior.ColorIndex = 4
                    Else
                        ws.Range("I" & Summary_Table_Row).Interior.ColorIndex = 3
                    End If
                
                'Print the percentage change to the summary table
                ws.Range("J" & Summary_Table_Row).Value = (CStr(Percent_Change) & "%")
                
                'Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                'Reset price change
                Price_Change = 0
                
                'Reset percent change
                Percent_Change = 0
                
                'Reset ending price
                Ending_Price = 0
                
                'Reset the total stock volume
                Total_Volume = 0
                
                'Find new stock's opening price
                Opening_Price = ws.Cells(i + 1, 3).Value
                
            'If the cell immediately following a row is the same ticker symbol...
            Else
                
                
                'Add to the total stock volume
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                
            End If
            
        Next i
        
     Next ws
                
End Sub
