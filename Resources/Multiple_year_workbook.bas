Attribute VB_Name = "Multiple_year_workbook"
Sub multiple_year_stock_data()

For Each ws In Worksheets

Dim tickername As String
Dim tickerresult As Integer
tickerresult = 2

Dim LR As Long
LR = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim open_price As Double
startprice = ws.Cells(2, 3).Value

Dim Close_price As Double
 
Dim yearlychange As Double

Dim percent_change As Double

Dim tickervolume As Double
tickervolume = 0


ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

For I = 2 To LR

    If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
        tickername = ws.Cells(I, 1).Value
        tickervolume = tickervolume + ws.Cells(I, 7).Value
        
        ws.Range("I" & tickerresult).Value = tickername
        
        endprice = ws.Cells(I, 6).Value
        yearlychange = (endprice - startprice)
        ws.Range("J" & tickerresult).Value = yearlychange
        
        If yearlychange >= -0 Then
            ws.Range("J" & tickerresult).Interior.ColorIndex = 4
                Else
            ws.Range("J" & tickerresult).Interior.ColorIndex = 3
        End If
        
        
        If (startprice = 0) Then
            percent_change = 0
        Else
            percent_change = (yearlychange / startprice)
        End If
        
        ws.Range("K" & tickerresult).Value = percent_change
        ws.Range("K" & tickerresult).NumberFormat = "0.00%"
        ws.Range("L" & tickerresult).Value = tickervolume
        
       
        startprice = ws.Cells(I + 1, 3).Value
        tickerresult = tickerresult + 1
        tickervolume = 0
        
    Else
        tickervolume = tickervolume + ws.Cells(I, 7).Value
    
    End If
        
    Next I
    
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        
    Dim Maxincrease As Double
    Dim Minincrease As Double
    Dim Maxvalue As Double
    Dim tickerCellRange As String
    Dim FoundCell As Range
    ' In order to find the row matching the percentage change value
    ' we need to fist set the format to general otherwise it will not find the exact value
    ws.Columns("K").NumberFormat = "General"
    ' Following line builds the string for the range e.g. K2:K3000
    tickerCellRange = "K2" & ":K" & tickerresult
    Maxincrease = WorksheetFunction.Max(ws.Range(tickerCellRange))
    ws.Cells(2, 17).Value = Maxincrease
    ' Now find the row by looking in the range
    Set FoundCell = ws.Range(tickerCellRange).Find(what:=Maxincrease)
    ws.Cells(2, 16).Value = ws.Cells(FoundCell.Row, 9).Value
    Minincrease = WorksheetFunction.Min(ws.Range(tickerCellRange))
    ws.Cells(3, 17).Value = Minincrease
    Set FoundCell = ws.Range(tickerCellRange).Find(what:=Minincrease)
    ws.Cells(3, 16).Value = ws.Cells(FoundCell.Row, 9).Value
    tickerCellRange = "L2" & ":L" & tickerresult
    Maxvalue = WorksheetFunction.Max(ws.Range(tickerCellRange))
    ws.Cells(4, 17).Value = Maxvalue
    Set FoundCell = ws.Range(tickerCellRange).Find(what:=Maxvalue)
    ws.Cells(4, 16).Value = ws.Cells(FoundCell.Row, 9).Value
    ws.Columns("K").NumberFormat = "0.00%"
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
       
    Next ws

End Sub

