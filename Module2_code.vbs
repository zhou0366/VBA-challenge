Attribute VB_Name = "Module1"

Sub Module2()

For Each ws In Worksheets
    'inital variables to hold yearly info per stock
    
    'var to hold current stock ticker
    Dim Ticker As String
    'var to count volume of a stock
    Dim VolumeTotal As Double
    'var to hold first year value of a stock
    Dim YearStartVal As Double
    'var to hold final year value of a stock
    Dim YearEndVal As Double
    'var to calculate yearly change
    Dim YearlyChange As Double
    'tracks percentage
    Dim Percentage As Double
    'counts how many rows in sheet
    Dim LastRow As Integer
    
    'count how many stocks we've viewed
    Dim StockNum As Integer
    StockNum = 0
    
    'vars to hold name of greatest value stocks
    Dim GainStock As String
    Dim LossStock As String
    Dim MostStock As String
    
    'rolling vars to actively track greatest values
    Dim GreatestGain As Double
    Dim GreatestLoss As Double
    Dim GreatestVolume As Double
    
    'determine number of rows in current sheet
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'write headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ws.Cells(1, 14).Value = "Greatest % increase"
    ws.Cells(1, 15).Value = "Greatest % decrease"
    ws.Cells(1, 16).Value = "Greatest Total Volume"
    
    
    'loop through each row
    For i = 2 To LastRow
       'scenario where we are counting for a new stock
       If ws.Cells(i, 1).Value <> Ticker Then
            YearlyChange = YearEndVal - YearStartVal
    
            If StockNum > 0 Then
            'output final vals to summary section
            ws.Cells(StockNum + 1, 9).Value = Ticker
            
            ws.Cells(StockNum + 1, 10).Value = YearlyChange
            
            'set color formatting based on whether or not the stock is negative
            If YearlyChange < 0 Then
                ws.Cells(StockNum + 1, 10).Interior.Color = RGB(250, 0, 0)
            Else
                ws.Cells(StockNum + 1, 10).Interior.Color = RGB(0, 250, 0)
            End If
            
            Percentage = YearlyChange / YearStartVal
            ws.Cells(StockNum + 1, 11).Value = Percentage
    
            If Percentage > GreatestGain Then
                GreatestGain = Percentage
                GainStock = Ticker
            End If
            If Percentage < GreatestLoss Then
                GreatestLoss = Percentage
                LossStock = Ticker
            End If
    
            'set percentage cell formatting
            ws.Cells(StockNum + 1, 11).NumberFormat = "0.00%"
            ws.Cells(StockNum + 1, 12).Value = VolumeTotal
            
            If VolumeTotal > GreatestVolume Then
                GreatestVolume = VolumeTotal
                MostStock = Ticker
            End If
            
        End If
            
       'set new ticker
       Ticker = ws.Cells(i, 1).Value
       
       'set year start value as the first value listed for the stock
       YearStartVal = ws.Cells(i, 3).Value
       
       'resets volume
       VolumeTotal = 0
       
       'increment stock count
       StockNum = StockNum + 1
       'scenario where we are counting for the same stock
       Else
            'update current value and total volume of stock traded
            YearEndVal = ws.Cells(i, 6).Value
            VolumeTotal = VolumeTotal + ws.Cells(i, 7).Value
       
       End If
       
    Next i
    
    'Give greatest increase, decrease, and volume
    
    'write final output
    ws.Cells(2, 14).Value = GainStock
    ws.Cells(2, 15).Value = LossStock
    ws.Cells(2, 16).Value = MostStock
    ws.Cells(3, 14).Value = GreatestGain
    ws.Cells(3, 14).NumberFormat = "0.00%"
    ws.Cells(3, 15).Value = GreatestLoss
    ws.Cells(3, 15).NumberFormat = "0.00%"
    ws.Cells(3, 16).Value = GreatestVolume
Next ws

End Sub

