# VBA-Challange
Sub Multible_year_stock_data()
    
'create a loop and set a variable for the worksheets

    Dim i As Long
    Dim ws As Worksheet

'Start loop

    For Each ws In Worksheets

 'Create  summary table toinput ticker,yearly change,% and total volume
 
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
    
  'Set variable for different data input to summary table and loop through all worksheet
  
        Dim ticker_symbol As String
 
        Dim total_vol As Double
        total_vol = 0
 
        Dim rowcount As Long
        rowcount = 2

        Dim year_open As Double
        year_open = 0

        Dim year_close As Double
        year_close = 0
        
        Dim year_change As Double
        year_change = 0

        Dim percent_change As Double
        percent_change = 0

        Dim lastrow As Long
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To lastrow
            
'Define and set conditions for data in summary table

  
    
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then

                year_open = ws.Cells(i, 3).Value

     End If

            total_vol = total_vol + ws.Cells(i, 7)

            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

'Put current ticker over in our summary table based on summaryrow
'Put totalvolume over in our summary table based on summaryrow
         
                ws.Cells(rowcount, 9).Value = ws.Cells(i, 1).Value

                ws.Cells(rowcount, 12).Value = total_vol

                year_close = ws.Cells(i, 6).Value

                year_change = year_close - year_open
                
                ws.Cells(rowcount, 10).Value = year_change
'Format yearly change to highlight positive change in green and negative change in red

                If year_change >= 0 Then
                
                    ws.Cells(rowcount, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(rowcount, 10).Interior.ColorIndex = 3
                    
   End If
                
                'Calculate Percent change
                
                If year_open = 0 And year_close = 0 Then
                    
                    percent_change = 0
                    
                    ws.Cells(rowcount, 11).Value = percent_change
                    
                    ws.Cells(rowcount, 11).NumberFormat = "0.00%"
                    
                ElseIf year_open = 0 Then
                
                    Dim percent_change_NA As String
                    
                    percent_change_NA = "NA"
                    
                    ws.Cells(rowcount, 11).Value = percent_change
                Else
                    percent_change = year_change / year_open
                    
                    ws.Cells(rowcount, 11).Value = percent_change
                    
                    ws.Cells(rowcount, 11).NumberFormat = "0.00%"
                End If

                rowcount = rowcount + 1
                
'Finally reset all data in summary table to avoid overlap or double counts
                
                total_vol = 0
                year_open = 0
                year_close = 0
                year_change = 0
                percent_change = 0
                
            End If
        
        Next i

'Create table to add functionality to the stock to show greatest % increase, decrease and total volume
       
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

        'Assign lastrow to count the number of rows in the summary table
        lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
'For greatest stock with the greatest increase assign best performance, worst for greatest degrease and most for greatest volume total

        Dim best_stock As String
        Dim best_value As Double
        Dim worst_stock As String
        Dim worst_value As Double
        Dim most_vol_stock As String
        Dim most_vol_value As Double
        
        best_value = ws.Cells(2, 11).Value

        worst_value = ws.Cells(2, 11).Value

        most_vol_value = ws.Cells(2, 12).Value

        'Loop to search through summary table
        For j = 2 To lastrow

 'Conditional to determine best performer
            If ws.Cells(j, 11).Value > best_value Then
                best_value = ws.Cells(j, 11).Value
                best_stock = ws.Cells(j, 9).Value
            End If

            'Conditional to determine worst performer
            If ws.Cells(j, 11).Value < worst_value Then
                worst_value = ws.Cells(j, 11).Value
                worst_stock = ws.Cells(j, 9).Value
            End If

            'Conditional to determine stock with the greatest volume traded
            If ws.Cells(j, 12).Value > most_vol_value Then
                most_vol_value = ws.Cells(j, 12).Value
                most_vol_stock = ws.Cells(j, 9).Value
            End If

        Next j

        'Move best performer, worst performer, and stock with the most volume items to the performance table
        ws.Cells(2, 16).Value = best_stock
        ws.Cells(2, 17).Value = best_value
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 16).Value = worst_stock
        ws.Cells(3, 17).Value = worst_value
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 16).Value = most_vol_stock
        ws.Cells(4, 17).Value = most_vol_value

        'Autofit table columns
        ws.Columns("I:L").AutoFit
        ws.Columns("O:Q").AutoFit
        

    Next ws
End Sub
