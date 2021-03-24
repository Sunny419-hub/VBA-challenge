Sub Multiple_year_stock()

    'FINAL CODE

    Worksheets(1).Activate
           
    'Set initial variables for looping tickers
    
    Dim Last_Ticker As Integer
    Last_Ticker = 2
    Dim Ticker_Temp As String
    Ticker_Temp = ""
    Dim Ticker_Temp2 As String
    Ticker_Temp2 = ""
            
    'Set initial variables for calculating Yearly Change
    Dim Open_p As Double
    Open_p = 0
    Dim Close_p As Double
    Close_p = 0
    Dim Yearly_Change As Double
    Yearly_Change = 0
    
    'Set initial variables for calculating Total Stock Volume
    Dim Temp_Volume As LongLong
    Temp_Volume = 0
        
    'Bonus
    Dim max As Variant
    Dim fin As Integer
    Dim min As Variant
    Dim max2 As Variant
    Dim fin2 As Integer
    Dim max3 As Variant
    Dim fin3 As LongLong
    
    'Loop through all tickers
    
    For Each ws In Worksheets
    
        'Set initial Name colums and rows
    
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
    
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
    
        Dim WorksheetName As String
        
        'Obtain the last row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        WorksheetName = ws.Name
        
        'Loop through all tickers
        For i = 2 To LastRow
            
            'Set test variable for debuging
            Ticker_Temp2 = ws.Cells(i, 1).Value
            
            'Set the sum for calculating Total Stock Volume
            Temp_Volume = ws.Cells(i, 7).Value + Temp_Volume
                                   
                                   
            'Comparing if the strings are equal(do nothing), if is not...
            If StrComp(Ticker_Temp2, Ticker_Temp, 0) = 0 Then
                          
                          
            Else
                'Asigning the ticker value
                ws.Cells(Last_Ticker, 9).Value = ws.Cells(i, 1).Value
                                        
                ' start in 2 to not count the headers
                If i = 2 Then
                    
                Else
                    
                    close_ptemp = ws.Cells(i - 1, 6).Value
                    Close_p = close_ptemp
                    Yearly_Change = Close_p - Open_p
            
                    'Asigning Yearly Change(next column)
                    ws.Cells(Last_Ticker - 1, 10).Value = Yearly_Change
                                        
                    ' Formatting positive change in green and negative change in red
                    
                    If Yearly_Change > 0 Then
                        ws.Cells(Last_Ticker - 1, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(Last_Ticker - 1, 10).Interior.ColorIndex = 3
                    End If
                    
                    'Calculating the Percent Change
                    If Open_p = 0 Then
                        ws.Cells(Last_Ticker - 1, 11).Value = 0
                    Else
                        Percent_Change = (Yearly_Change / Open_p)
                        'Asigning Percent Change(next column)
                        ws.Cells(Last_Ticker - 1, 11).Value = Format(Percent_Change, "0.00%")
                    End If
                    
                    'Asigning Total Stock Volume(next column)
                    
                    ws.Cells(Last_Ticker - 1, 12).Value = Temp_Volume
                    
                    Temp_Volume = 0
                    
                    
                End If
                                
                       
                Open_p = ws.Cells(i, 3).Value
                
                'add 1 to print the ticker in the next row
                Last_Ticker = Last_Ticker + 1
                
                'when there is a change in the ticker then update
                Ticker_Temp = ws.Cells(i, 1).Value
                
            'End Comparing if the strings are equal
            End If
            
            If i = LastRow Then
                
                close_ptemp = ws.Cells(i, 6).Value
                Close_p = close_ptemp
                Yearly_Change = Close_p - Open_p
            
                'Asigning Yearly Change(next column)
                ws.Cells(Last_Ticker - 1, 10).Value = Yearly_Change
                
                ' Formatting positive change in green and negative change in red
                If Yearly_Change > 0 Then
                    ws.Cells(Last_Ticker - 1, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(Last_Ticker - 1, 10).Interior.ColorIndex = 3
                End If
                    
                'Calculating the Percent Change
                If Open_p = 0 Then
                    ws.Cells(Last_Ticker - 1, 11).Value = 0
                Else
                    Percent_Change = (Yearly_Change / Open_p)
                    'Asigning Percent Change(next column)
                    ws.Cells(Last_Ticker - 1, 11).Value = Format(Percent_Change, "0.00%")
                    'Asigning Total Stock Volume(next column)
                    ws.Cells(Last_Ticker - 1, 12).Value = Temp_Volume
                    Temp_Volume = 0
                End If
            End If
        Next i
        
        'Greatest % Increase
         fin = ws.Cells(Rows.Count, 11).End(xlUp).Row
         max = Application.WorksheetFunction.max(Range("K2:K" & fin))
    
        'Asigning value
        ws.Cells(2, 17).Value = Format(max, "0.00%")
                
        'Greatest % Decrease
        fin2 = ws.Cells(Rows.Count, 11).End(xlUp).Row
        max2 = Application.WorksheetFunction.min(Range("K2:K" & fin))
        
        'Asigning value
    
        ws.Cells(3, 17).Value = Format(max2, "0.00%")
    
    
        'Greatest Greatest Total Volume
        fin3 = ws.Cells(Rows.Count, 12).End(xlUp).Row
        max3 = Application.WorksheetFunction.max(Range("l2:l" & fin))
    
        'Asigning value
        ws.Cells(4, 17).Value = (max3)
        Last_Ticker = 2
    Next ws

End Sub