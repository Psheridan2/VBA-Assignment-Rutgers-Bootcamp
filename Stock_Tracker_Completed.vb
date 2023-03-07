Sub Stock_tracker()
    Dim WS_Count As Integer
    Dim W As Integer
    WS_Count = ActiveWorkbook.Worksheets.count
 
    For W = 1 To WS_Count
        ActiveWorkbook.Worksheets(W).Cells(1, 9).Value = "Ticker"
        ActiveWorkbook.Worksheets(W).Cells(1, 10).Value = "Yearly Change"
        ActiveWorkbook.Worksheets(W).Cells(1, 11).Value = "Percent Change"
        ActiveWorkbook.Worksheets(W).Cells(1, 12).Value = "Total Stock Volume"
        ActiveWorkbook.Worksheets(W).Cells(2, 15).Value = "Greatest % Increase"
        ActiveWorkbook.Worksheets(W).Cells(3, 15).Value = "Greatest % Decrease"
        ActiveWorkbook.Worksheets(W).Cells(4, 15).Value = "Greatest Total Volume"
        ActiveWorkbook.Worksheets(W).Cells(1, 16).Value = "Ticker"
        ActiveWorkbook.Worksheets(W).Cells(1, 17).Value = "Value"
        Dim count As Integer
        count = 1
        Dim num_row, num_row2 As Double
        num_row = ActiveWorkbook.Worksheets(W).Cells(Rows.count, 1).End(xlUp).Row 'Finding last row with a value in Column A
        num_row2 = ActiveWorkbook.Worksheets(W).Cells(Rows.count, 10).End(xlUp).Row 'Same purpose as before but for Columns J-L
        Dim ticker, tickerchange As String
        Dim openval, closeval As Double
        Dim sumvol As Double
        Dim maxpercent, minpercent, maxvol As Double
        sumvol = 0
        tickerchange = ""
        For I = 2 To num_row + 1 'Added plus one to identify first blank row at the end of data set
            ticker = ActiveWorkbook.Worksheets(W).Cells(I, 1).Value
            If ticker <> tickerchange Then 'if A2:A* is not equal to "" then
                count = count + 1
                ActiveWorkbook.Worksheets(W).Cells(count, 9).Value = ticker
                If tickerchange <> "" Then
                    Dim j As Double
                    j = I - 1
                    closeval = ActiveWorkbook.Worksheets(W).Cells(j, 6).Value
                    ActiveWorkbook.Worksheets(W).Cells(count - 1, 10).Value = closeval - openval
                    ActiveWorkbook.Worksheets(W).Cells(count - 1, 12).Value = sumvol
                    sumvol = 0
                    If ActiveWorkbook.Worksheets(W).Cells(count - 1, 10).Value < 0 Then
                        ActiveWorkbook.Worksheets(W).Cells(count - 1, 10).Interior.ColorIndex = 3
                    ElseIf ActiveWorkbook.Worksheets(W).Cells(count - 1, 10).Value >= 0 Then
                        ActiveWorkbook.Worksheets(W).Cells(count - 1, 10).Interior.ColorIndex = 4
                    End If
                    ActiveWorkbook.Worksheets(W).Cells(count - 1, 11).Value = (closeval - openval) / openval
                    ActiveWorkbook.Worksheets(W).Cells(count - 1, 11).NumberFormat = "0.00%"
                End If
                openval = ActiveWorkbook.Worksheets(W).Cells(I, 3).Value 'Openval set to first instance of open (C Column) for each new ticker value
                tickerchange = ticker 'Updating memory for ticker value
            End If
            sumvol = ActiveWorkbook.Worksheets(W).Cells(I, 7).Value + sumvol
 
 
        Next I
        minpercent = 0
        maxpercent = 0
        maxvol = 0
        For y = 2 To num_row2
            If ActiveWorkbook.Worksheets(W).Cells(y, 11).Value > maxpercent Then
 
                maxpercent = ActiveWorkbook.Worksheets(W).Cells(y, 11).Value
                ActiveWorkbook.Worksheets(W).Cells(2, 17).Value = maxpercent
                ActiveWorkbook.Worksheets(W).Cells(2, 16).Value = ActiveWorkbook.Worksheets(W).Cells(y, 9).Value
 
            End If
            If ActiveWorkbook.Worksheets(W).Cells(y, 11).Value < minpercent Then
 
                minpercent = ActiveWorkbook.Worksheets(W).Cells(y, 11).Value
                ActiveWorkbook.Worksheets(W).Cells(3, 17).Value = minpercent
                ActiveWorkbook.Worksheets(W).Cells(3, 16).Value = ActiveWorkbook.Worksheets(W).Cells(y, 9).Value
 
            End If
            If ActiveWorkbook.Worksheets(W).Cells(y, 12).Value > maxvol Then
 
                maxvol = ActiveWorkbook.Worksheets(W).Cells(y, 12).Value
                ActiveWorkbook.Worksheets(W).Cells(4, 17).Value = maxvol
                ActiveWorkbook.Worksheets(W).Cells(4, 16).Value = ActiveWorkbook.Worksheets(W).Cells(y, 9).Value
 
            End If
            ActiveWorkbook.Worksheets(W).Cells(2, 17).NumberFormat = "0.00%"
            ActiveWorkbook.Worksheets(W).Cells(3, 17).NumberFormat = "0.00%"
        Next y
 
 
 
    Next W
End Sub