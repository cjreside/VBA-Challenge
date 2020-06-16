Attribute VB_Name = "Module11"
Sub Stock_Data()

    'Loop through all worksheets
    For Each ws In Worksheets
    
        'Set Initial Variable for Holding the "Ticker" Value
        Dim Ticker_Symbol As String
        
        'Set initial variable for the Open Price
        Dim Open_Price As Double
        
        'Set initial variable for Close Price
        Dim Close_Price As Double
    
        'Set an initial variable for holding the yearly change from opening price at the beginning of a given year to the closing price at the end of that year
        Dim Ticker_Yearly_Change As Double
    
        'Set an initial variable for holding the percent change from opening price at the beginning of a given year to the closing price at the end of that year.
        Dim Ticker_Percent_Change As Double
    
        'Set an variable for holding the total stock volume
        Dim Total_Stock_Volume As Double
        Total_Stock_Volume = 0
    
        'Keep Track of the location for Ticker Name
        Dim Ticker_Summary_Row As Integer
        Ticker_Summary_Row = 2
        
        'Set initial Open Price
        Open_Price = ws.Cells(2, 3).Value
        
        'Determine the last row
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
         'Add the word "Ticker" to the column
        ws.Cells(1, 9).Value = "Ticker"
    
        'Add the words "Yearly Change" to the column
        ws.Cells(1, 10).Value = "Yearly Change"
    
        'Add the words "Percent Change" to the column
        ws.Cells(1, 11).Value = "Percent Change"
    
        'Add the words "Total Stock Volume" to the column
        ws.Cells(1, 12).Value = "Total Stock Volume"
     
            'Loop through the rows
            For i = 2 To lastrow
        
                'Check if we are still within the same Ticker Symbol
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                    'Set the Ticker Symbol
                    Ticker_Symbol = ws.Cells(i, 1).Value
                
                    'Set the Stock Volume
                    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                
                    'Apply Ticker Symbol and Ticker Stock Volume to ranges I and L with the Ticker Summary Table Row
                    ws.Range("I" & Ticker_Summary_Row).Value = Ticker_Symbol
                    ws.Range("L" & Ticker_Summary_Row).Value = Total_Stock_Volume
                
                    'Set Close Price to column 6
                    Close_Price = ws.Cells(i, 6).Value
            
                    'Add the Yearly Change
                    Ticker_Yearly_Change = (Close_Price - Open_Price)
                    ws.Range("j" & Ticker_Summary_Row).Value = Ticker_Yearly_Change
            
                        'Add the Percent Change
                        If Open_Price = 0 Then
                        Ticker_Percent_Change = 0
             
                        Else
                            Ticker_Percent_Change = Ticker_Yearly_Change / Open_Price
                        
                        End If
                
                'Apply Ticker Percent Change to range K and format as a percentage (%)
                ws.Range("K" & Ticker_Summary_Row).Value = Ticker_Percent_Change
                ws.Range("K" & Ticker_Summary_Row).NumberFormat = "0.00%"
                
                'Reset the Ticker Stock Volume
                Total_Stock_Volume = 0
                
                'Reset the Ticker Summary Row and add one
                Ticker_Summary_Row = Ticker_Summary_Row + 1
                
                'Reset the Open Price
                Open_Price = ws.Cells(i + 1, 3)
           
            Else
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
            
            'Add a conditional for color coding the yearly change
            If ws.Cells(i, 10).Value < 0 Then
                    ws.Cells(i, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(i, 10).Interior.ColorIndex = 4
            End If

        End If
            
        Next i
        
        'Set Greatest % Increase, % Decrease, and Total Volume
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        
        'Loop Through Summary Table to find the greatest % increase, freatest % decrease, and greatest total volume
        New_Summary_Table = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
            'Loop through the Summary Table tickers
            For i = 2 To New_Summary_Table
                If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & New_Summary_Table)) Then
                    ws.Range("O2").Value = ws.Cells(i, 9).Value
                    ws.Range("P2").Value = ws.Cells(i, 11).Value
                    ws.Range("P2").NumberFormat = "0.00%"
                
                ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & New_Summary_Table)) Then
                    ws.Range("O3").Value = ws.Cells(i, 9).Value
                    ws.Range("P3").Value = ws.Cells(i, 11).Value
                    ws.Range("P3").NumberFormat = "0.00%"
                    
                ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & New_Summary_Table)) Then
                    ws.Range("O4").Value = ws.Cells(i, 9).Value
                    ws.Range("P4").Value = ws.Cells(i, 12).Value
                    
                End If

        Next i
    
 Next ws

End Sub
