' Multiple_year_stock_data HW - Mary Phillipuk


Sub LoopThroughWorkBook()
    ' Loop through sheets in this workbook
    
    Dim ws As Worksheet
    ' var for current sheet

    For Each ws In ThisWorkbook.Worksheets
       ' for every sheet is this workbook, do the following
       
       ' if statement below used during dev to check against sample screen shots
       ' If ws.Name = "2015" Then
       
            ws.Activate
            ' activate the sheet
            totalTickerVolumes
            ' read/compute values in that sheet
            showMaxMin
            ' draw table at right that shows max/min vals for the sheet
            
       ' End If

    Next ws

End Sub


Sub totalTickerVolumes()
    ' Go through active sheet and read/aggregate values per ticker symbol
    
    Dim totalVol As Double
    ' sum of volume per ticker symbol
    Dim ticker As String
    ' current ticker
    Dim openVal As Double
    ' opening val for current ticker
    Dim closeVal As Double
    ' closing val for current ticker
    
    Dim lastrow As Long
    ' var for last row of values in sheet
    Dim i As Long
    ' loop iterator
    Dim grow As Integer
    ' var for row to output aggregated ticker vals
    
    ' initialize some vars to set up loop, catch first ticker val
    grow = 1
    ticker = ""
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).row
    
    ' set up header row for ticker aggregated vals
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    For i = 2 To lastrow
        ' loop through rows
        
        If Cells(i, 1).Value <> ticker Then
            ' if current row ticker var doesn't equal last one saved
            
            If (ticker <> "") Then
                ' if the last ticker wasn't the empty string, output vals using printRow subroutine with params
                Call printRow(grow, ticker, openVal, closeVal, totalVol)
            End If
            
            ' set ticker to the new val and reset other vars to compute and output for new ticker
            grow = grow + 1
            ticker = Cells(i, 1).Value
            openVal = Cells(i, 3).Value
            closeVal = Cells(i, 6).Value
            totalVol = Cells(i, 7).Value
            
        Else
            ' otherwise add current row volume total volume for this ticker and set closing val to latest
            closeVal = Cells(i, 6).Value
            totalVol = totalVol + Cells(i, 7).Value
            ' check if new, openVal was zero before
            If openVal = 0 Then
                openVal = Cells(i, 3).Value
            End If
        End If
        
       ' if statement below used during dev to check just a few ticker rows against sample screen shots
       ' If ticker = "ABC" Then
       '     Exit For
        ' End If
          
    Next i

    ' output vals for last ticker saved using printRow subroutine with params
    Call printRow(grow, ticker, openVal, closeVal, totalVol)
    
    ' autofit the columns in ticker output table
    Range("I:L").EntireColumn.AutoFit
End Sub


Sub printRow(row, ticker, openVal, closeVal, vol)
    ' output values given the current output row, ticker, opening value, closing value, and total volume
    
    Dim cindex As Integer
    ' color index var
    Dim dif As Double
    ' difference value, closing - opening values
    Dim perc As Double
    ' difference as a percentage of opening value
    
    ' compute values, prevent zero division
    dif = closeVal - openVal
    If openVal <> 0 Then
        perc = dif / openVal
    Else
        ' default if never had an opening value in year
        perc = 0
    End If
    
    ' output values
    Cells(row, 9).Value = ticker
    Cells(row, 10).Value = dif
    
    
    cindex = CInt((255 - Abs(perc * 255)) / 1.5)
    ' compute color index as a function of percent; expected cindex range = 0-255 with perc value range = 0-1
    ' divide by 1.5 to get more obvious tint (lower cindex = more intense color below)
    
    If (cindex < 0) Then
        ' in case perc > 1 set min cindex value at 0
        cindex = 0
    End If
    
    ' set positive change cell backgrounds to green, negative change cell backgrounds to red
    If (dif > 0) Then
        'Cells(row, 10).Interior.ColorIndex = 4
        Cells(row, 10).Interior.Color = RGB(cindex, 255, cindex)
    ElseIf (dif < 0) Then
        'Cells(row, 10).Interior.ColorIndex = 3
        Cells(row, 10).Interior.Color = RGB(255, cindex, cindex)
    End If
    
    ' output percent and volume values and format percent cells
    Cells(row, 11).Value = perc
    Cells(row, 11).NumberFormat = "0.00%"
    Cells(row, 12).Value = vol
End Sub


Sub showMaxMin()
    ' show max and min percent changes and max total volume for active sheet
    
    Dim lastrow As Long
    ' last row with data
    Dim i As Long
    ' loop iterator
    
    Dim pmaxVal As Double
    ' max percent value (greatest increase)
    Dim pmaxTick As String
    ' ticker symbol for max percent value
    Dim pminVal As Double
    ' min percent value (greatest decrease)
    Dim pminTick As String
    ' ticker symbol for min percent value
    Dim vmaxVal As Double
    ' max volume
    Dim vmaxTick As String
    ' ticker symbol for max volume
    
    Dim perc As Double
    ' current percent val
    Dim tick As String
    ' current ticker symbol
    Dim vol As Double
    ' current volume
    
    ' initialize vars
    lastrow = Cells(Rows.Count, 9).End(xlUp).row
    pmaxVal = 0
    pminVal = 0
    vmaxVal = 0
    
    ' create row and column labels
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    ' loop through aggregated ticker data rows
    For i = 2 To lastrow
        ' read current percent, ticker symbol, and total volume in row
        perc = Cells(i, 11).Value
        tick = Cells(i, 9).Value
        vol = Cells(i, 12).Value
        
        If perc > pmaxVal Then
            ' if current percent > max saved , then set max to current percent and save current ticker as pmaxTick
            pmaxVal = perc
            pmaxTick = tick
        ElseIf perc < pminVal Then
            ' if current percent < min saved , then set min to current percent and save current ticker as pminTick
            pminVal = perc
            pminTick = tick
        End If
        If vol > vmaxVal Then
            ' if current volume > max saved , then set max to current volume and save current ticker as vmaxTick
            vmaxVal = vol
            vmaxTick = tick
        End If
        
    Next i
    
    ' output saved max, min, and ticker values in appropriate cells
    Cells(2, 16).Value = pmaxTick
    Cells(2, 17).Value = pmaxVal
    Cells(3, 16).Value = pminTick
    Cells(3, 17).Value = pminVal
    Cells(4, 16).Value = vmaxTick
    Cells(4, 17).Value = vmaxVal
    
    ' format percent cells and autofit columns
    Range("Q2:Q3").NumberFormat = "0.00%"
    Range("O:Q").EntireColumn.AutoFit
End Sub
