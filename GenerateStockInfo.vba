'***************************************************************************
'Purpose: Loop through all the sheets and calculate the stock information
'   and greatest changes tables
'Inputs: sheet, total rows
'Outputs: none
'***************************************************************************
Sub GenerateStockInformation()
    Dim ws As Worksheet
    Dim totalRows As Long
    
    For Each ws In Worksheets
        'Debug.Print ("Generating headings for worksheet " + ws.Name)
        CreateHeadings ws
        
        'Debug.Print ("Calculating stock information for worksheet " + ws.Name)
        totalRows = CalculateStockInformation(ws)
        
        'Print summary of the greatest changes
        'Debug.Print ("Calculating greatest changes from stock information for worksheet " + ws.Name)
        CalculateGreatestChanges ws, totalRows

    Next ws

End Sub

'***************************************************************************
'Purpose: Calculate the stock information for an individual sheet
'   Create a script that will loop through all the stocks for one year and output the following information.
'   The ticker symbol.
'   Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'   The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'   The total stock volume of the stock.
'   You should also have conditional formatting that will highlight positive change in green and negative change in red.
'Inputs: sheet, total rows
'Outputs: total rows generated in the table
'***************************************************************************
Function CalculateStockInformation(sheet As Worksheet) As Long
    Dim lastRow As Long
    Dim summaryTableRow As Long
    
    Dim tickerSymbol As String
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    
    'Initialize variables
    summaryTableRow = 2
    
    'Initialize the yearly change with the opening price
    yearlyChange = sheet.Cells(2, 3)
    totalVolume = 0
    
    ' Determine the Last Row
    lastRow = sheet.Cells(sheet.Rows.Count, 1).End(xlUp).row
    
    
    'Debug only
    'lastRow = 500
    
    For i = 2 To lastRow
    
        'Check we are within the same stock symbol
        If sheet.Cells(i + 1, 1).Value <> sheet.Cells(i, 1).Value Then
            ' Set the ticker symbol
            tickerSymbol = sheet.Cells(i, 1).Value
            
            'Finalize the yearly change subtracting off the opening price
            If yearlyChange > 0 Then
                percentChange = (sheet.Cells(i, 6).Value - yearlyChange) / yearlyChange
            Else
                percentChange = 0
            End If
            yearlyChange = sheet.Cells(i, 6).Value - yearlyChange
            
            
            sheet.Range("I" & summaryTableRow).Value = tickerSymbol
                        
            sheet.Range("J" & summaryTableRow).Value = yearlyChange
            ' Format the cell
            FormatCell sheet, summaryTableRow, 10
            
            sheet.Range("K" & summaryTableRow).Value = percentChange
            'Number formatting taken from https://www.educba.com/vba-number-format/
            sheet.Range("K" & summaryTableRow).NumberFormat = "0.00%"
            
            sheet.Range("L" & summaryTableRow).Value = totalVolume
        
            summaryTableRow = summaryTableRow + 1
            
            'Initialize the yearly change with the opening price
            yearlyChange = sheet.Cells(i + 1, 3)
            totalVolume = 0
        Else
            totalVolume = totalVolume + sheet.Cells(i, 7).Value
            'Print Tickers
            'Debug.Print Range("A" & i).Value
        End If
                
    Next i
    
    CalculateStockInformation = summaryTableRow - 1
    
End Function

'***************************************************************************
'Purpose: Calculate the �Greatest % increase�, �Greatest % decrease� and �Greatest
'  total volume� for an individual sheet
'Inputs: sheet, total rows
'Outputs: none
'***************************************************************************
Sub CalculateGreatestChanges(sheet As Worksheet, totalRows As Long)
    Dim greatestPercInc As Double
    Dim greatestPercIncTicker As String
    Dim greatestPercDec As Double
    Dim greatestPercDecTicker As String
    Dim greatestTotalVol As Double
    Dim greatestTotalVolString As String
    
    greatestPercInc = sheet.Cells(2, 11).Value
    greatestPercIncTicker = sheet.Cells(2, 9).Value
    greatestPercDec = sheet.Cells(2, 11).Value
    greatestPercDecTicker = sheet.Cells(2, 9).Value
    greatestTotalVol = sheet.Cells(2, 12).Value
    greatestTotalVolTicker = sheet.Cells(2, 9).Value
    
    'Calculate the greatest percent increase, percent decrease and total volume
    For i = 3 To totalRows
        If (sheet.Cells(i, 11).Value > greatestPercInc) Then
            greatestPercInc = sheet.Cells(i, 11).Value
            greatestPercIncTicker = sheet.Cells(i, 9).Value
        End If
        
        If (sheet.Cells(i, 11).Value < greatestPercDec) Then
            greatestPercDec = sheet.Cells(i, 11).Value
            greatestPercDecTicker = sheet.Cells(i, 9).Value
        End If
        
        If (sheet.Cells(i, 12).Value > greatestTotalVol) Then
            greatestTotalVol = sheet.Cells(i, 12).Value
            greatestTotalVolTicker = sheet.Cells(i, 9).Value
        End If

    Next i
        
    'Print the headers
    'Could break this out into separate subroutine too
    sheet.Range("P1").Value = "Ticker"
    sheet.Range("Q1").Value = "Value"
    sheet.Range("O2").Value = "Greatest % Increase"
    sheet.Range("O3").Value = "Greatest % Decrease"
    sheet.Range("O4").Value = "Greatest Total Volume"

    'Print the values
    sheet.Range("P2").Value = greatestPercIncTicker
    sheet.Range("Q2").Value = greatestPercInc
    sheet.Range("Q2").NumberFormat = "0.00%"
    sheet.Range("P3").Value = greatestPercDecTicker
    sheet.Range("Q3").Value = greatestPercDec
    sheet.Range("Q3").NumberFormat = "0.00%"
    sheet.Range("P4").Value = greatestTotalVolTicker
    sheet.Range("Q4").Value = greatestTotalVol

End Sub

'***************************************************************************
'Purpose: Helper method to print the headings for the generated stock
'   information.
'Inputs: sheet
'Outputs: none
'***************************************************************************
Sub CreateHeadings(sheet As Worksheet)
    sheet.Range("I1").Value = "Ticker"
    sheet.Range("J1").Value = "Yearly Change"
    sheet.Range("K1").Value = "Percent Change"
    sheet.Range("L1").Value = "Total Stock Volume"
End Sub

'***************************************************************************
'Purpose: Helper method to format color of a cell to red if negative and
'   green if positive.
'Inputs: sheet, row, column
'Outputs: none
'***************************************************************************
Sub FormatCell(sheet As Worksheet, row As Long, col As Long)
    If (sheet.Cells(row, col).Value < 0) Then
        sheet.Cells(row, col).Interior.ColorIndex = 3
    ElseIf (sheet.Cells(row, col).Value > 0) Then
        sheet.Cells(row, col).Interior.ColorIndex = 4
    End If
End Sub

