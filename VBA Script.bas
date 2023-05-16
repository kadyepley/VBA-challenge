Attribute VB_Name = "Module1"
Sub Challenge_2()

' set all the variables that we are going to use in the equations later in the code
' the name of the ticker
Dim ticker As String
' the # of tickers in each worksheet
Dim ticker_count As Double
' last row in the worksheet
Dim thelastrow As Long
' the opening price for a ticker on a row
Dim open_price As Double
' the closing price for a ticker on a row
Dim close_price As Double
' the change in the price by the year
Dim yearly_change As Double
' the percent of the change in comparison to the open price
Dim percent_change As Double
' the total stock volume for a ticker
Dim total_stock_vol As Double
' the greatest increase in price as a percent
Dim greatest_percent_inc As Double
' the name of the ticker with the greatest percent increase
Dim greatest_percent_inc_name As String
' the greatest decrease in price as a percent
Dim greatest_percent_dec As Double
' the name of the ticker with the greatest percent decrease
Dim greatest_percent_dec_name As String
' the highest volume for a ticker
Dim greatest_stock_vol As Double
' the name of the ticker with the highest volume
Dim greatest_stock_vol_name As String

' loop in through each worksheet in the workbook
For Each ws In Worksheets

    ws.Activate
    
    thelastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row

    ' labels for chart 1
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ' clear out variables for each worksheet
    ticker_count = 0
    ticker = " "
    yearly_change = 0
    open_price = 0
    percent_change = 0
    total_stock_vol = 0
    
    ' set the loop to run through the first line of data to the end
    For i = 2 To thelastrow

       ' identify the name of the ticker we are looking at
        ticker = Cells(i, 1).Value
        
        ' get the open price for the year
        If open_price = 0 Then
            open_price = Cells(i, 3).Value
        End If
        
        ' equation for the total stock volume for the ticker
        total_stock_vol = total_stock_vol + Cells(i, 7).Value
        
        ' if the ticker changes
        If Cells(i + 1, 1).Value <> ticker Then
            ' add these together when the name of the ticker changes
            ticker_count = ticker_count + 1
            Cells(ticker_count + 1, 9) = ticker
            
            ' identify where the close price is
            close_price = Cells(i, 6)
            
            ' equation for yearly change and where to put the output
            yearly_change = close_price - open_price
            Cells(ticker_count + 1, 10).Value = yearly_change
            
            ' format green if there is an increase
            If yearly_change > 0 Then
                Cells(ticker_count + 1, 10).Interior.ColorIndex = 4
            ' format red if there is no increase
            ElseIf yearly_change <= 0 Then
                Cells(ticker_count + 1, 10).Interior.ColorIndex = 3
            End If
            
            ' calculate the percent change for the ticker
            If open_price = 0 Then
                percent_change = 0
            Else
                percent_change = (yearly_change / open_price)
            End If
            
            ' format the change values to be percents
            Cells(ticker_count + 1, 11).Value = Format(percent_change, "Percent")
            
            ' reset open price for the next ticker
            open_price = 0
            
            ' add in the total stock volume
            Cells(ticker_count + 1, 12).Value = total_stock_vol
            
            ' reset the total stock volume
            total_stock_vol = 0
        End If
        
        
    Next i
    
    ' set up the labels for the second chart
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    ' set the last row again
    thelastrow = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    ' identify where values are in the ws
    greatest_percent_inc = Cells(2, 11).Value
    greatest_percent_inc_name = Cells(2, 9).Value
    greatest_percent_dec = Cells(2, 11).Value
    greatest_percent_dec_name = Cells(2, 9).Value
    greatest_stock_vol = Cells(2, 12).Value
    greatest_stock_vol_name = Cells(2, 9).Value
    
    
    ' set the loop to run through the first line of data to the end
    For i = 2 To thelastrow
    
        ' identify the greatest percent increase
        If Cells(i, 11).Value > greatest_percent_inc Then
            greatest_percent_inc = Cells(i, 11).Value
            greatest_percent_inc_name = Cells(i, 9).Value
        End If
        
        ' identify the greatest percent decrease
        If Cells(i, 11).Value < greatest_percent_dec Then
            greatest_percent_dec = Cells(i, 11).Value
            greatest_percent_dec_name = Cells(i, 9).Value
        End If
        
        ' indentify the ticker with the greatest stock volume
        If Cells(i, 12).Value > greatest_stock_vol Then
            greatest_stock_vol = Cells(i, 12).Value
            greatest_stock_vol_name = Cells(i, 9).Value
        End If
        
    Next i
    
    ' place the values in the chart
    Range("P2").Value = Format(greatest_percent_inc_name, "Percent")
    Range("P3").Value = Format(greatest_percent_dec_name, "Percent")
    Range("P4").Value = greatest_stock_vol_name
    Range("Q2").Value = Format(greatest_percent_inc, "Percent")
    Range("Q3").Value = Format(greatest_percent_dec, "Percent")
    Range("Q4").Value = greatest_stock_vol
    
    'formatting all columns added to be the right width
    ws.Range("I:L").Columns.AutoFit
    ws.Range("O:Q").Columns.AutoFit
    
Next ws


End Sub

