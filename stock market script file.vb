Sub Main()
    'loop through all sheets
    For j = 1 To ActiveWorkbook.Worksheets.Count
       Call stock_market(Worksheets(j))
    Next j
    
    'return back to the first sheet
    Worksheets(1).Select
End Sub

Sub stock_market(myws As Worksheet)

    ' Set an initial variable for holding the ticker_symbol
    Dim ticker_symbol As String

    ' Set an initial variable for holding the total_stock_volume
    Dim total_stock_volume As LongLong
    total_stock_volume = 0
  
    ' Set an initial variable for holding the opening_price
    Dim opening_price As Double
  
    ' Set an initial variable for holding the closing_price
    Dim closing_price As Double
    
    ' Set an initial variable for holding the yearly_change
    Dim yearly_change As Double
    
    ' Keep track of the location for each ticker_symbol in the summary table
    Dim Summary_Table_Row As Integer

    Summary_Table_Row = 2
    
    ' activate the current worksheet
    myws.Select
    
    ' Add column headers
    Range("I1").FormulaR1C1 = "ticker symbol"
    Range("J1").FormulaR1C1 = "yearly change"
    Range("K1").FormulaR1C1 = "percent change"
    Range("L1").FormulaR1C1 = "total stock volume"
    Range("I1:L1").Font.Bold = True
    Columns("A:L").EntireColumn.EntireColumn.AutoFit
    
    ' set opening_price
    opening_price = Cells(2, 3).Value

    ' Loop through all stock data
    For i = 2 To ActiveSheet.UsedRange.Rows.Count
        
        ' Check if we are still within the same ticker_symbol, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
          
            ' Set the ticker_symbol
            ticker_symbol = Cells(i, 1).Value

            ' Set closing_price
            closing_price = Cells(i, 6).Value
              
            ' Set yearly_change
            yearly_change = (closing_price - opening_price)
            
            ' Add to the total_stock_volume total
            total_stock_volume = total_stock_volume + Cells(i, 7).Value
        
            ' Print the ticker_symbol in the Summary Table
            Range("I" & Summary_Table_Row).Value = ticker_symbol
              
            ' Print the total_stock_volume total for each stock to the Summary Table
            Range("L" & Summary_Table_Row).Value = total_stock_volume
              
            ' Print the yearly_change in the Summary Table
            Range("J" & Summary_Table_Row).Value = yearly_change
       
            ' Print the percent_change in the Summary Table
            Range("K" & Summary_Table_Row).Value = yearly_change / opening_price
            
            ' Print the % change in the Summary Table
            Range("K" & Summary_Table_Row).NumberFormat = "0.00%"

            'conditional formatting of yearly_change column
            If Range("J" & Summary_Table_Row).Value < 0 Then
                Range("J" & Summary_Table_Row).Interior.Color = RGB(255, 0, 0)
            End If
            If Range("J" & Summary_Table_Row).Value > 0 Then
                Range("J" & Summary_Table_Row).Interior.Color = RGB(0, 255, 0)
            End If
            
            'conditional formatting of percent_change column
            If Range("K" & Summary_Table_Row).Value < 0 Then
                Range("K" & Summary_Table_Row).Interior.Color = RGB(255, 0, 0)
            End If
            If Range("K" & Summary_Table_Row).Value > 0 Then
                Range("K" & Summary_Table_Row).Interior.Color = RGB(0, 255, 0)
            End If
            
            
            ' initialise for next iteration of summary result
            Summary_Table_Row = Summary_Table_Row + 1

            ' Reset the total_stock_volume
            total_stock_volume = 0
            
            ' Reset opening_price
            opening_price = Cells(i + 1, 3).Value
        
        ' If the cell immediately following a row is the same ticker_symbol...
        Else

            ' Add to the total_stock_volume
            total_stock_volume = total_stock_volume + Cells(i, 7).Value

        End If

    Next i

End Sub

