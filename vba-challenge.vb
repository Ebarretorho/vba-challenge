VBA HOMERWORK - code

Sub StockPrices()

'----- Initiate the loop for worksheets -----

Dim ws As Worksheet

For Each ws In Worksheets
ws.Activate

'------- Define the variables -------

'Ticker Name
Dim ticker_name As String

'Stock Price
Dim stock_price As Double

'Variance %
Dim variance_percent As Double

Dim start_price As Double

Dim close_price As Double

'Volume
Dim Volume As Double
Volume = 0

'Summary Sheet Row
Dim summary_sheet_row As Integer
summary_sheet_row = 2

'Finde the last row of data to set up the for loop
last_row = Cells(Rows.Count, 1).End(xlUp).Row


' ----- Formatting the summary table for each worksheet --------

    Cells(summary_sheet_row - 1, 10).Value = "Ticker"
    Cells(summary_sheet_row - 1, 11).Value = "YoY Prince Change"
    Cells(summary_sheet_row - 1, 12).Value = "Percent Change"
    Cells(summary_sheet_row - 1, 13).Value = "Total Stock Volume"
    Cells(summary_sheet_row - 1, 10).Interior.ColorIndex = 15
    Cells(summary_sheet_row - 1, 11).Interior.ColorIndex = 15
    Cells(summary_sheet_row - 1, 12).Interior.ColorIndex = 15
    Cells(summary_sheet_row - 1, 13).Interior.ColorIndex = 15
    Cells(summary_sheet_row - 1, 10).Font.Bold = True
    Cells(summary_sheet_row - 1, 11).Font.Bold = True
    Cells(summary_sheet_row - 1, 12).Font.Bold = True
    Cells(summary_sheet_row - 1, 13).Font.Bold = True

    start_price = Cells(2, 6).Value

'-------------- START THE FOR LOOP FOR EACH WORKSHEET DATA -------------------
    
    For i = 2 To last_row
    
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
            ' Ticker Name
            ticker_name = Cells(i, 1).Value
            'print Ticker Name
            Cells(summary_sheet_row, 10).Value = ticker_name
        
            'Volume running metric
             Volume = Volume + Cells(i, 7).Value
            'Print Volume
            Cells(summary_sheet_row, 13).Value = Volume
            Cells(summary_sheet_row, 13).Style = "Normal"
        
            'Set Up Close Price
            close_price = Cells(i, 6)
    
            'Print Price
        
            Dim change_price As Double
            change_price = close_price - start_price
        
            Cells(summary_sheet_row, 11).Value = change_price
            Cells(summary_sheet_row, 11).Style = "Currency"
         
            'Format the prince change field
            
                If change_price < 0 Then
         
                Cells(summary_sheet_row, 11).Interior.ColorIndex = 38
                Cells(summary_sheet_row, 11).Font.ColorIndex = 30
         
                Else
                Cells(summary_sheet_row, 11).Interior.ColorIndex = 50
                Cells(summary_sheet_row, 11).Font.ColorIndex = 4
         
                End If
                
            ' Calculate the percentage change and make sure no errors come from dividing by 0
            
            Dim index_change As Double
                
                If start_price = 0 Then
                index_change = 0
                Cells(summary_sheet_row, 12).Value = index_change
                Cells(summary_sheet_row, 12).NumberFormat = "0.00%"
                
                Else
                
                index_change = (close_price / start_price) - 1
                Cells(summary_sheet_row, 12).Value = index_change
                Cells(summary_sheet_row, 12).NumberFormat = "0.00%"
                
                End If

                If index_change < 0 Then
         
                Cells(summary_sheet_row, 12).Interior.ColorIndex = 38
                Cells(summary_sheet_row, 12).Font.ColorIndex = 30
         
                Else
                Cells(summary_sheet_row, 12).Interior.ColorIndex = 50
                Cells(summary_sheet_row, 12).Font.ColorIndex = 4
                
                End If
        
        
            'Reset Prices
            
            'Start price set to the first price for new ticker
            start_price = Cells(i + 1, 6)
            
            'Close price set to zero until next change is found
            close_price = 0
        
            summary_sheet_row = summary_sheet_row + 1
        
            Volume = 0
        
            Else
        
            Volume = Volume + Cells(i, 7).Value
        
            End If
        
        Next i
        
        Columns("J:M").AutoFit
        
                ' ------------------ Find the greatest increase/decrease/Volume ----------------------------

                ' Find the last row for the summary tables for the for loop
                
                last_row_2 = Cells(Rows.Count, 10).End(xlUp).Row
                
                Dim lead_increase As Double
                Dim lead_deacrease As Double
                Dim lead_volume As Double
                Dim lead_increase_tick As String
                Dim lead_deacrease_tick As String
                Dim lead_volume_ticker As String
                
                lead_increase = Cells(2, 12).Value
                lead_decrease = Cells(2, 12).Value
                lead_volume = Cells(2, 13).Value
                
                For i = 2 To last_row_2
                
                    'Find greatest increase
                    If Cells(i, 12).Value > lead_increase Then
                    
                    'Change lead increase value to hold higher number
                    lead_increase = Cells(i, 12).Value
                    'Change tick name to match the value
                    lead_increase_tick = Cells(i, 10).Value
                    
                    Else
                    
                    lead_increase = lead_increase
                    
                    End If
                    
                    
                    'Find greatest Decrease
                    If Cells(i, 12).Value < lead_decrease Then
                    
                    'Change lead increase value to hold higher number
                    lead_decrease = Cells(i, 12).Value
                    'Change tick name to match the value
                    lead_decrease_tick = Cells(i, 10).Value
                    
                    Else
                    
                    lead_decrease = lead_decrease
                    
                    End If
                    
                    'Find greatest Volume
                    If Cells(i, 13).Value > lead_volume Then
                    
                    'Change lead increase value to hold higher number
                    lead_volume = Cells(i, 13).Value
                    'Change tick name to match the value
                    lead_volume_ticker = Cells(i, 10).Value
                    
                    Else
                    
                    lead_volume = lead_volume
                    
                    End If
                    
                    
                Next i
                
                
                '------------------------------ FORMAT/BUIL THE SUMMARY DATA FOR THE INDICATORS CALCULATED ABOVE -----------------------------
                
                Cells(1, 15).Value = "Category"
                Cells(1, 16).Value = "Ticker"
                Cells(1, 17).Value = "Value"
                Range("O1:Q1").Interior.ColorIndex = 15
                Range("O1:Q1").Font.Bold = True
                Cells(2, 16).Value = lead_increase_tick
                Cells(3, 16).Value = lead_decrease_tick
                Cells(4, 16).Value = lead_volume_ticker
                Cells(2, 17).Value = lead_increase
                Cells(3, 17).Value = lead_decrease
                Cells(4, 17).Value = lead_volume
                Cells(2, 15).Value = "Greatest % Increase"
                Cells(3, 15).Value = "Greatest % Decrease"
                Cells(4, 15).Value = "Greatest Total Volume"
                Range("Q2:Q3").NumberFormat = "0.00%"
                Range("P4").Style = "Normal"
                Columns("O:Q").AutoFit
                                
Next ws

    
End Sub