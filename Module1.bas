Attribute VB_Name = "Module1"
Sub Stocks()
'Dim Variable
Dim ticker As String
Dim vol As Double
Dim LastRow As Double
Dim i As Double
Dim j As Double
Dim ws As Worksheet
Dim RowDisplay As Integer
Dim TickerOpen As Double
Dim TickerClose As Double
Dim c As Range
    
    
    
    For Each ws In Worksheets
      ws.Activate
      RowDisplay = 2
      LastRow = Cells(Rows.Count, "A").End(xlUp).Row
'Write Header Row
      Cells(1, 9).Value = "Ticker"
      Cells(1, 10).Value = "Yearly Change"
      Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    
      TickerOpen = Range("C2")
   For i = 2 To LastRow
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            TickerClose = Cells(i, 6).Value
            vol = vol + Cells(i, 7).Value
            Range("I" & RowDisplay).Value = Cells(i, 1)
            Range("L" & RowDisplay).Value = vol
            Range("J" & RowDisplay).Value = TickerClose - TickerOpen
            Range("K" & RowDisplay).Value = (TickerClose - TickerOpen) / TickerOpen
            
            
            If Range("J" & RowDisplay).Value < 0 Then
            Range("J" & RowDisplay).Interior.ColorIndex = 3
            Else
            Range("J" & RowDisplay).Interior.ColorIndex = 4
            End If
            
            
'reset Varables for the new stock
             vol = 0
             TickerOpen = Cells(i + 1, 3).Value
             RowDisplay = RowDisplay + 1
        End If
        
'increment the total
             
        vol = vol + Cells(i, 7).Value
        Next i
        
        Range("K2:K" & RowDisplay).NumberFormat = "0.00%"
Next ws

End Sub
