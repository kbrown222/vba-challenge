Attribute VB_Name = "Module1"
 'Create a script that loops through all the stocks for one year and outputs the following information:
            'The ticker symbol
            'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
            'The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
            'The total stock volume of the stock.
            
 
Sub stock_data():
 
        'Loop through each worksheet
        For Each ws In Worksheets
 
                'Set initial variables for Ticker Symbol, Opening Value, Closing Value, Yearly Change, Percent Change
                'Total Stock Volume, and Summary Table
                
                Dim Ticker As String
                
                Dim Opening_Value As Double
                Opening_Value = 0
                
                Dim Closing_Value As Double
                Closing_Value = 0
                
                Dim Yearly_Change As Double
                
                Dim Percent_Change As Double
                Percent_Change = 0
                
                Dim Total_Stock_Volume As LongLong
                Total_Stock_Volume = 0
                
                Dim Summary_Table_Row As Integer
                Summary_Table_Row = 2
                
                'Label and format column headers of summary table
                ws.Cells(1, 9).Value = "Ticker"
                ws.Cells(1, 10).Value = "Yearly Change"
                ws.Cells(1, 11).Value = "Percent Change"
                ws.Cells(1, 12).Value = "Total Stock Volume"
                ws.Columns("I:L").AutoFit
                
                
                'Determine the Last Row on each worksheet
                LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

          'Loop through all stock data rows
          For i = 2 To LastRow
          
                'Start Calculating Total Stock Volume
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

                'When the new ticker starts
                If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ' Set the Ticker Symbol name
                Ticker = ws.Cells(i, 1).Value
                
                ' Print the Ticker Symbol in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                
                'Set the opening value
                Opening_Value = ws.Cells(i, 3).Value
                
                End If
'
                'When the current ticker ends...
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Gather the closing value
                Closing_Value = ws.Cells(i, 6).Value
                
                'Populate the ticker stock volume
                ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
                
                'Calculate the Yearly Change
                Yearly_Change = Closing_Value - Opening_Value
                
                               
                'Calculate Percentage Change of Ticker Symbol (this needs to be formated to a decimal place.
                Percent_Change = ((Closing_Value - Opening_Value) / Opening_Value)
                
                'Print the Yearly Change in the Summary Table
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                
                'Based on the value of the Yearly Change column, color the value red if the value is negative and green if the value is positive.
                        If ws.Range("J" & Summary_Table_Row).Value > 0 Then
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                
                        ElseIf ws.Range("J" & Summary_Table_Row).Value < 0 Then
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                            
                        Else: ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 0
                        
                        End If
                
                'Print the Percent Change in the Summary Table
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                
                'Go to the next row in the Summary table
                Summary_Table_Row = Summary_Table_Row + 1
                
                'Reset the Total Stock Volume to 0
                Total_Stock_Volume = 0
                
                End If
                
    
                              
        'Restart the for loop for the next ticker symbol
        Next i

        
     'Restart the for ws loop for the next worksheet
    Next ws

End Sub



