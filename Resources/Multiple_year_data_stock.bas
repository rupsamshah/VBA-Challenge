Attribute VB_Name = "Multiple_year_data_stock"
Sub multiple_year_stock_data()


Dim tickername As String
Dim tickerresult As Integer
tickerresult = 2

Dim LR As Long
LR = Cells(Rows.Count, 1).End(xlUp).Row

Dim open_price As Double
startprice = Cells(2, 3).Value

Dim Close_price As Double
 
Dim yearlychange As Double

Dim percent_change As Double

Dim tickervolume As Double
tickervolume = 0


Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

For I = 2 To LR

    If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
        tickername = Cells(I, 1).Value
        tickervolume = tickervolume + Cells(I, 7).Value
        
        Range("I" & tickerresult).Value = tickername
        
        endprice = Cells(I, 6).Value
        yearlychange = (endprice - startprice)
        Range("J" & tickerresult).Value = yearlychange
        
        If yearlychange >= -0 Then
            Range("J" & tickerresult).Interior.ColorIndex = 4
                Else
            Range("J" & tickerresult).Interior.ColorIndex = 3
        End If
        
        
        If (startprice = 0) Then
            percent_change = 0
        Else
            percent_change = (yearlychange / startprice)
        End If
        
        Range("K" & tickerresult).Value = percent_change
        Range("L" & tickerresult).Value = tickervolume
        
       
        startprice = Cells(I + 1, 3).Value
        tickerresult = tickerresult + 1
        tickervolume = 0
        
    Else
        tickervolume = tickervolume + Cells(I, 7).Value
    
    End If
        
    Next I
    
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        
        
    Dim Maxincrease As Double
    Dim Minincrease As Double
    Dim Maxvalue As Double
    Dim tickerCellRange As String
    Dim FoundCell As Range
    ' In order to find the row matching the percentage change value
    ' we need to fist set the format to general otherwise it will not find the exact value
    Columns("K").NumberFormat = "General"
    ' Following line builds the string for the range e.g. K2:K3000
    tickerCellRange = "K2" & ":K" & tickerresult
    Maxincrease = WorksheetFunction.Max(Range(tickerCellRange))
    Cells(2, 17).Value = Maxincrease
    ' Now find the row by looking in the range
    Set FoundCell = Range(tickerCellRange).Find(what:=Maxincrease)
    Cells(2, 16).Value = Cells(FoundCell.Row, 9).Value
    Minincrease = WorksheetFunction.Min(Range(tickerCellRange))
    Cells(3, 17).Value = Minincrease
    Set FoundCell = Range(tickerCellRange).Find(what:=Minincrease)
    Cells(3, 16).Value = Cells(FoundCell.Row, 9).Value
    tickerCellRange = "L2" & ":L" & tickerresult
    Maxvalue = WorksheetFunction.Max(Range(tickerCellRange))
    Cells(4, 17).Value = Maxvalue
    Set FoundCell = Range(tickerCellRange).Find(what:=Maxvalue)
    Cells(4, 16).Value = Cells(FoundCell.Row, 9).Value
    Columns("K").NumberFormat = "0.00%"
      

End Sub

