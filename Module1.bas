Attribute VB_Name = "Module1"
Sub stocklooper()
   
    For Each ws In Worksheets
    
        ' fill column names
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Stock Volume"
        ws.Cells(1, 16) = "Ticker"
        ws.Cells(1, 17) = "Value"
        ws.Cells(2, 15) = "Greatest % Increase"
        ws.Cells(3, 15) = "Greatest % Decrease"
        ws.Cells(4, 15) = "Greatest Total Volume"
        
        ' find the last row
        LastRowData = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Initialize a counter for unique ticker symbols
        unique_ticker_count = 0
        
        ' Initialize an aggregator to track total trade volumes
        Total_Volume = 0
        
        ' Get the opening value of the first ticker-- used in loop below
        ticker_value_start = ws.Cells(2, 3)
        
        ' initialize values for max % increase, max % decrease, and max volume
        max_increase = 0
        max_decrease = 0
        max_volume = 0
        
        For r = 2 To LastRowData
        
            ' Add volume to the total volume aggregator
            Total_Volume = Total_Volume + ws.Cells(r, 7)
    
            ' Searches for when the value of the next cell is different than that of the current cell
            If ws.Cells(r, 1) <> ws.Cells(r + 1, 1) Then
              
              ' Write the ticker symbol
              ws.Cells(unique_ticker_count + 2, 9) = ws.Cells(r, 1)
              
              ' Capture the ticker's end value
              ticker_value_end = ws.Cells(r, 6)
              
              ' Capture the numerical and percent differences of the ticker throughout the year
              ws.Cells(unique_ticker_count + 2, 10) = ticker_value_end - ticker_value_start
              
              ws.Cells(unique_ticker_count + 2, 11) = ticker_value_end / ticker_value_start - 1
              
              ' Write the total volume
              ws.Cells(unique_ticker_count + 2, 12) = Total_Volume
              
              ' Check if % change is greater than the current max increase
              If (ticker_value_end / ticker_value_start - 1) > max_increase Then
                    
                    max_increase = ticker_value_end / ticker_value_start - 1
                    max_increase_ticker = ws.Cells(r, 1)
                    
              End If
              
              ' Check if % change is less than current max decrease
              If (ticker_value_end / ticker_value_start - 1) < max_decrease Then
              
                    max_decrease = ticker_value_end / ticker_value_start - 1
                    max_decrease_ticker = ws.Cells(r, 1)
                    
              End If
              
              ' Check if volume is more than the current max
              If Total_Volume > max_volume Then
              
                    max_volume = Total_Volume
                    max_volume_ticker = ws.Cells(r, 1)
                    
              End If
              
              ' Reset the total volume for the new stock
              Total_Volume = 0
              
              ' Reset the ticker value for the new stock
              ticker_value_start = ws.Cells(r + 1, 3)
              
              ' Increment the unique ticker counter
              unique_ticker_count = unique_ticker_count + 1
        
            End If
    
        Next r
        
        ' write max values to summary table
        ws.Cells(2, 16) = max_increase_ticker
        ws.Cells(2, 17) = max_increase
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Cells(3, 16) = max_decrease_ticker
        ws.Cells(3, 17) = max_decrease
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Cells(4, 16) = max_volume_ticker
        ws.Cells(4, 17) = max_volume
        
        ' Create range object for formatting
        Dim MyRange As Range
        
        ' Apply color formatting to value changes
        Set MyRange = ws.Range("J2:J" & (unique_ticker_count + 1))

        ' Add first rule
        With MyRange.FormatConditions.Add(xlCellValue, xlGreater, "=0")
                .Interior.Color = vbGreen
        End With
                
        ' add second rule
        With MyRange.FormatConditions.Add(xlCellValue, xlLess, "=0")
                .Interior.Color = vbRed
        End With
        
        ' Apply color formatting to % changes
        Set MyRange = ws.Range("K2:K" & (unique_ticker_count + 1))
        
        MyRange.NumberFormat = "0.00%"

        ' Add first rule
        With MyRange.FormatConditions.Add(xlCellValue, xlGreater, "=0")
                .Interior.Color = vbGreen
        End With
                
        ' add second rule
        With MyRange.FormatConditions.Add(xlCellValue, xlLess, "=0")
                .Interior.Color = vbRed
        End With
                
    Next ws
    
End Sub

