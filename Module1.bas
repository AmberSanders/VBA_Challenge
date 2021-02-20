Attribute VB_Name = "Module1"
 Sub Stock_Market_Analysis()
    
    For Each WS In Worksheets
    
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Percentage_Change As Double
    Dim Total_Stock_Volume As Double
    Dim Row_Count As Long
    Dim Summary_Row As Long
    Dim Open_Price As Double
    Dim Close_Price As Double
    
    
    WS.Range("H1").Value = "Ticker"
    WS.Range("I1").Value = "Yearly_Change"
    WS.Range("J1").Value = "Percentage_Change"
    WS.Range("K1").Value = "Total_Stock_Volume"
    
    Summary_Row = 2
    Open_Price = WS.Cells(2, 3).Value
    Closed_Price = 0
    Yearly_Change = 0
    Percentage_Change = 0
    
    Total_Stock_Volume = 0
    
    RowCount = WS.Cells(Rows.Count, "A").End(xlUp).Row
    
    
    For i = 2 To RowCount
    
    
    If WS.Cells(i, 1).Value = WS.Cells(i + 1, 1).Value Then
    Total_Stock_Volume = Total_Stock_Volume + WS.Cells(i, 7).Value
    
    
    ElseIf WS.Cells(i, 1).Value <> WS.Cells(i + 1, 1).Value Then
    Total_Stock_Volume = Total_Stock_Volume + WS.Cells(i, 7).Value
    
    
    Ticker = WS.Cells(i, 1).Value
    WS.Cells(Summary_Row, 8).Value = Ticker
    WS.Cells(Summary_Row, 11).Value = Total_Stock_Volume
    
    
    Close_Price = WS.Cells(i, 6).Value
    Yearly_Change = Close_Price - Open_Price
    WS.Cells(Summary_Row, 9).Value = Yearly_Change
    
    
    If Open_Price <> 0 Then
    Percent_Change = (Close_Price - Open_Price) / Open_Price
    WS.Cells(Summary_Row, 10).Value = Percent_Change
    WS.Cells(Summary_Row, 10).NumberFormat = "0.00%"

    
    Summary_Row = Summary_Row + 1
    Total_Stock_Volume = 0
    Open_Price = WS.Cells(i + 1, 3).Value
    Closed_Price = 0
    Yearly_Change = 0
    
    End If
    
    If WS.Cells(Summary_Row, 9).Value > 0 Then
    WS.Cells(Summary_Row, 9).Interior.ColorIndex = 4
    
    Else
    WS.Cells(Summary_Row, 9).Interior.ColorIndex = 3
    
    End If
    
    
    End If
    
    Next i
    Next WS
    
    
    End Sub
    
