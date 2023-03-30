Attribute VB_Name = "Module1"
Sub stocks()

'Set up some variables
Dim LastRow As Long
Dim TickerName As String
Dim rownum As Integer
Dim Opener As Double
Dim Closer As Double
Dim Change As Double
Dim SumCell As Long
Dim StockTotal As Double
Dim arr(2 To 3001) As Variant
Dim arry(2 To 3001) As Variant
Dim aray(2 To 3001) As Variant

'First we loop through all sheets
For Each ws In Worksheets
    'Find the last row in the sheet
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Next loop through the first column (ticker) and copy it over to the I column
    rownum = 2
    For i = 2 To LastRow
        ws.Cells(1, 9).Value = "Ticker" 'sets first cell in column for title (Ticker)
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then 'check for different row
            TickerName = ws.Cells(i, 1).Value 'grabs name
            ws.Range("I" & rownum).Value = TickerName 'copies name to other column
            rownum = rownum + 1 'goes to the next row
        End If
    Next i
    
    'Next set the "yearly Change" column
    rownum = 2
    For j = 1 To LastRow - 1
        ws.Cells(1, 10).Value = "Yearly Change" 'Title of column
        If ws.Cells(j + 1, 1).Value <> ws.Cells(j, 1).Value Then 'check for different row
            Opener = ws.Cells(j + 1, 3).Value 'grabs the open value
            Closer = ws.Cells(j + 251, 6).Value 'grabs the close value
            Change = (Closer - Opener) 'finds the change
            ws.Range("J" & rownum).Value = Change 'puts change into new column
            If Change > 0 Then 'checks for positive change
                ws.Range("J" & rownum).Interior.ColorIndex = 4 'makes it green
            Else
                ws.Range("J" & rownum).Interior.ColorIndex = 3 'if negative change makes it red
            End If
            rownum = rownum + 1 'goes to the next row
        End If
    Next j
    
    'Next set the Percent Change column
    rownum = 2
    For j = 1 To LastRow - 1
        ws.Cells(1, 11).Value = "Percent Change" 'Title of column
        If ws.Cells(j + 1, 1).Value <> ws.Cells(j, 1).Value Then 'check for different row
            Opener = ws.Cells(j + 1, 3).Value 'grabs the open value
            Closer = ws.Cells(j + 251, 6).Value 'grabs the close value
            Change = ((Closer - Opener) / Opener) * 100 'finds the percent change
            ws.Range("K" & rownum).Value = (Round(Change, 2) & "%") 'puts percent change into new column
            rownum = rownum + 1 'goes to the next row
        End If
    Next j
    
    'Next sum up the total stock
    rownum = 2
    StockTotal = 0
    For i = 2 To LastRow
        ws.Cells(1, 12).Value = "Total Stock Volume" 'sets first cell in column for title
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then 'check for different row
            StockTotal = StockTotal + (Cells(i, 7).Value / 100) 'add to the stock total and reduce number
            ws.Range("L" & rownum).Value = StockTotal 'sets total into new column
            rownum = rownum + 1 'goes to the next row
            StockTotal = 0 'resets total
        Else
            StockTotal = StockTotal + (Cells(i, 7).Value / 100)
        End If
    Next i
    For i = 2 To ws.Cells(Rows.Count, 12).End(xlUp).Row
        ws.Cells(i, 12).Value = ws.Cells(i, 12).Value * 100
    Next i
    ws.Range("L:L").NumberFormat = "0"
    
    'Set up some titles
    ws.Cells(2, 15).Value = "Greatest % increase"
    ws.Cells(3, 15).Value = "Greatest % decrease"
    ws.Cells(4, 15).Value = "Greatest total volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    'Next find the biggest and smallest in the percent column
    For i = 2 To ws.Cells(Rows.Count, 9).End(xlUp).Row
        arr(i) = ws.Cells(i, 11).Value 'makes a list of Percent Change values
        arry(i) = ws.Cells(i, 12).Value 'makes a list of Total Stock Volume values
        aray(i) = ws.Cells(i, 9).Value 'makes a list of Ticker values
    Next i
    
    'Putting max and min values into cells
    ws.Cells(2, 17).Value = WorksheetFunction.Max(arr())
    ws.Cells(3, 17).Value = WorksheetFunction.Min(arr())
    ws.Cells(4, 17).Value = WorksheetFunction.Max(arry())
    ws.Range("Q1:Q2").NumberFormat = "0.00%"
    
    For i = 2 To LastRow
        If ws.Cells(2, 17).Value = ws.Cells(i, 11).Value Then
            ws.Cells(2, 16).Value = ws.Cells(i, 9)
        End If
    Next i
    
    For i = 2 To LastRow
        If ws.Cells(3, 17).Value = ws.Cells(i, 11).Value Then
            ws.Cells(3, 16).Value = ws.Cells(i, 9)
        End If
    Next i
    
    For i = 2 To LastRow
        If ws.Cells(4, 17).Value = ws.Cells(i, 12).Value Then
            ws.Cells(4, 16).Value = ws.Cells(i, 9)
        End If
    Next i
Next ws
    
End Sub
