Attribute VB_Name = "Module2"
' Create a script that loops through all the stocks for one year and outputs the following information:

' The ticker symbol.
' Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
' The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
' The total stock volume of the stock.
' Note: Make sure to use conditional formatting that will highlight positive change in green and negative change in red.


' yearlyChange = lastClosingPrice - firstOpeningPrice
' percentChange = yearlyChange - openingPrice/opening price
' lastRow = Cells(Rows.Count, 1).End(xlUp).Row


Sub tickerSymbol():
  
  For Each ws In Worksheets
    ws.Select
   
   Range("I1").Value = "Ticker Name"
   Range("J1").Value = "Yearly Change"
   Range("K1").Value = "Percent Change"
   Range("L1").Value = "Total Stock Volume"
   
   ticker = ""
   ticker_symbol = 2
   lastRow = Cells(Rows.Count, 1).End(xlUp).Row
   total_stock = 0
   
   For Row = 2 To lastRow
   
     If Cells(Row + 1, 1).Value <> Cells(Row, 1).Value Then
         ticker = Cells(Row, 1).Value
         Cells(ticker_symbol, 9).Value = ticker
         
         total_stock = total_stock + Cells(Row, 7)
         Cells(ticker_symbol, 12).Value = total_stock
         
         ticker_symbol = ticker_symbol + 1
         total_stock = 0
         
     Else
         total_stock = total_stock + Cells(Row, 7).Value
    
     End If

   Next Row
   
   ' yearlyChange = lastClosingPrice - firstOpeningPrice
   
   
   Dim yearlyChange As Double
   yearlyChange = 0
   yearly_table = 2
   firstOpeningPrice = Cells(2, 3).Value
   lastClosingPrice = Cells(2, 6).Value
   ' startRow = 2
   
   For Row2 = 2 To lastRow
   
     If Cells(Row2 + 1, 1).Value <> Cells(Row2, 1).Value Then
         lastClosingPrice = Cells(Row2, 6).Value
         yearlyChange = lastClosingPrice - firstOpeningPrice
         Cells(yearly_table, 10).Style = "Currency"
         
         Cells(yearly_table, 10).Value = yearlyChange
         yearly_table = yearly_table + 1
         
         firstOpeningPrice = Cells(Row2 + 1, 3).Value
     
     ' Else
        ' lastClosingPrice = Cells(Row2, 6).Value
    
     End If

   Next Row2
   
 ' percentChange = yearlyChange/opening price
   
   Dim percentChange As Double
   percentChange = 0
   percent_table = 2
   firstOpeningPrice = Cells(2, 3).Value
   lastClosingPrice = Cells(2, 6).Value

   For Row3 = 2 To lastRow
   
       If Cells(Row3 + 1, 1).Value <> Cells(Row3, 1).Value Then
         lastClosingPrice = Cells(Row3, 6).Value
         yearlyChange = lastClosingPrice - firstOpeningPrice
         
         percentChange = yearlyChange / firstOpeningPrice
         Cells(percent_table, 11).Value = percentChange
         Cells(percent_table, 11).NumberFormat = "0.00%"
         
         percent_table = percent_table + 1
         
         firstOpeningPrice = Cells(Row3 + 1, 3).Value
         
       End If
        
            If Cells(percent_table, 11).Value > 0 Then
               Cells(percent_table, 11).Interior.ColorIndex = 4
             
            Else
               Cells(percent_table, 11).Interior.ColorIndex = 3
            
            End If
        
    Next Row3
    
    ' Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". The solution should match the following image:
       
            Dim maxPercent, minPercent, maxVolume As LongLong
            Range("N2").Value = "Greatest % increase"
            Range("N3").Value = "Greatest % decrease"
            Range("N4").Value = "Greatest Total Volume"
            Range("O1").Value = "Ticker"
            Range("P1").Value = "Value"
            
            lastRowp = Cells(Rows.Count, 9).End(xlUp).Row

            maxPercent = WorksheetFunction.Max(Range("K2:K" & lastRow))
            maxTickerNameIndex = WorksheetFunction.Match(maxPercent, Range("K2:K" & lastRowp), 0)
            Range("O2").Value = Range("I" & maxTickerNameIndex + 1).Value
            Range("P2").Value = maxPercent
            Cells(2, 16).NumberFormat = "0.00%"
            
            minPercent = WorksheetFunction.Min(Range("K2:K" & lastRowp))
            minTickerNameIndex = WorksheetFunction.Match(minPercent, Range("K2:K" & lastRowp), 0)
            Range("O3").Value = Range("I" & minTickerNameIndex + 1).Value
            Range("P3").Value = minPercent
            Cells(3, 16).NumberFormat = "0.00%"
            
            maxVolume = WorksheetFunction.Max(Range("L2:L" & lastRowp))
            maxTickerNameIndex = WorksheetFunction.Match(maxVolume, Range("L2:L" & lastRowp), 0)
            Range("O4").Value = Range("I" & maxTickerNameIndex + 1).Value
            Range("P4").Value = maxVolume
            Cells(4, 16).NumberFormat = "0.0 E+0"
            
            Next ws
        
End Sub


 












