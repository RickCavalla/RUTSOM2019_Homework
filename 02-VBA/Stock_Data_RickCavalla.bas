Option Explicit

Sub Stock_Data_Sheet(pvWorksheet As Worksheet)
    'Total Existing Data Rows
    Dim lRowCount As Long
    
    'Existing Data variables
    Dim sTicker As String
    Dim sNextTicker As String
    Dim dOpen As Double
    Dim dClose As Double
    Dim lVol As Long
    
    'New Totals variables
    Dim dYearChangeFlat As Double
    Dim dYearChangePercent As Double
    Dim dVolTotal As Double
    
    'Maxium variables
    Dim dMaxIncrease As Double
    Dim dMaxDecrease As Double
    Dim dMaxVol As Double
    Dim sMaxIncreaseTicker As String
    Dim sMaxDecreaseTicker As String
    Dim sMaxVolTicker As String
    
    'Iterator
    Dim lRowLoop As Long
    
    'Column locations of existing data
    Dim iTickerCol As Integer
    Dim iOpenCol As Integer
    Dim iCloseCol As Integer
    Dim iVolCol As Integer
    
    'Column locations of totals data
    Dim iTickerTotalCol As Integer
    Dim iChangeFlatCol As Integer
    Dim iChangePercentCol As Integer
    Dim iVolTotalCol As Integer
    Dim lCurrentTotalRow As Long
    
    'Starting column of maximums data
    Dim iMaxCol As Integer
    
    'Assign a bunch of column variables
    iTickerCol = 1
    iOpenCol = 3
    iCloseCol = 6
    iVolCol = 7
    iTickerTotalCol = 9
    iChangeFlatCol = 10
    iChangePercentCol = 11
    iVolTotalCol = 12
    iMaxCol = 14
    
    'Sequence to increment for writing totals data
    lCurrentTotalRow = 1
    
    'Set column headers of totals data
    pvWorksheet.Cells(lCurrentTotalRow, iTickerTotalCol).Value = "Ticker"
    pvWorksheet.Cells(lCurrentTotalRow, iChangeFlatCol).Value = "Yearly Change"
    pvWorksheet.Cells(lCurrentTotalRow, iChangePercentCol).Value = "Percent Change"
    pvWorksheet.Cells(lCurrentTotalRow, iVolTotalCol).Value = "Total Volume"
    
    'Set formats of totals data so it is easier to read
    pvWorksheet.Columns(iChangeFlatCol).ClearFormats
    pvWorksheet.Columns(iChangePercentCol).ClearFormats
    pvWorksheet.Columns(iChangeFlatCol).NumberFormat = "0.00"
    pvWorksheet.Columns(iChangePercentCol).Style = "Percent"
    pvWorksheet.Columns(iChangePercentCol).NumberFormat = "0.00%"
    
    'Used "Record Macro" to trick Excel into writing conditional format code below
    'Just tweaked to use my variables instead of "Selection"
    
    'Set conditional format for positive changes (green)
    pvWorksheet.Columns(iChangeFlatCol).FormatConditions.Add _
        Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
        
    With pvWorksheet.Columns(iChangeFlatCol).FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(0, 255, 0)
        .TintAndShade = 0
    End With
    
    'Set conditional format for negative changes (red)
    pvWorksheet.Columns(iChangeFlatCol).FormatConditions.Add _
        Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
        
    With pvWorksheet.Columns(iChangeFlatCol).FormatConditions(2).Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(255, 0, 0)
        .TintAndShade = 0
    End With
    
    'Clear format of header rows so they do not color red or green
    pvWorksheet.Cells(1, iChangeFlatCol).ClearFormats
    pvWorksheet.Cells(1, iChangePercentCol).ClearFormats
    
    'Set initial opening price
    dOpen = pvWorksheet.Cells(2, iOpenCol)
    
    'I googled a way to get row count that was not dependent on cursor movement
    lRowCount = pvWorksheet.UsedRange.Rows.Count
    
    'loop through all existing data rows
    For lRowLoop = 2 To lRowCount
        'Grab existing data values and increment volume total
        sTicker = pvWorksheet.Cells(lRowLoop, iTickerCol)
        sNextTicker = pvWorksheet.Cells(lRowLoop + 1, iTickerCol)
        lVol = pvWorksheet.Cells(lRowLoop, iVolCol)
        dVolTotal = dVolTotal + lVol
        
        'If transitioning between tickers, write totals data
        If sTicker <> sNextTicker Then
            dClose = pvWorksheet.Cells(lRowLoop, iCloseCol)
            dYearChangeFlat = dClose - dOpen
            
            'Avoid division by zero
            If dOpen <> 0 Then
                dYearChangePercent = dYearChangeFlat / dOpen
            Else
                dYearChangePercent = 0
            End If
            
            'Increment where we are printing next batch of totals
            lCurrentTotalRow = lCurrentTotalRow + 1
            
            'Check if we have new maximums in current totals
            If dYearChangePercent > dMaxIncrease Then
                dMaxIncrease = dYearChangePercent
                sMaxIncreaseTicker = sTicker
            End If
            
            If dYearChangePercent < dMaxDecrease Then
                dMaxDecrease = dYearChangePercent
                sMaxDecreaseTicker = sTicker
            End If
            
            If dVolTotal > dMaxVol Then
                dMaxVol = dVolTotal
                sMaxVolTicker = sTicker
            End If
            
            'Write totals data
            pvWorksheet.Cells(lCurrentTotalRow, iTickerTotalCol).Value = sTicker
            pvWorksheet.Cells(lCurrentTotalRow, iChangeFlatCol).Value = dYearChangeFlat
            pvWorksheet.Cells(lCurrentTotalRow, iChangePercentCol).Value = dYearChangePercent
            pvWorksheet.Cells(lCurrentTotalRow, iVolTotalCol).Value = dVolTotal
            
            'Reset some variables before starting next ticker
            dVolTotal = 0
            dOpen = pvWorksheet.Cells(lRowLoop + 1, iOpenCol)
        End If
    Next lRowLoop
    
    'Write maximums
    pvWorksheet.Cells(1, iMaxCol) = "Greatest % Increase"
    pvWorksheet.Cells(1, iMaxCol + 1) = sMaxIncreaseTicker
    pvWorksheet.Cells(1, iMaxCol + 2) = dMaxIncrease
    pvWorksheet.Cells(1, iMaxCol + 2).Style = "Percent"
    pvWorksheet.Cells(1, iMaxCol + 2).NumberFormat = "0.00%"
    
    pvWorksheet.Cells(2, iMaxCol) = "Greatest % Decrease"
    pvWorksheet.Cells(2, iMaxCol + 1) = sMaxDecreaseTicker
    pvWorksheet.Cells(2, iMaxCol + 2) = dMaxDecrease
    pvWorksheet.Cells(2, iMaxCol + 2).Style = "Percent"
    pvWorksheet.Cells(2, iMaxCol + 2).NumberFormat = "0.00%"
    
    pvWorksheet.Cells(3, iMaxCol) = "Greatest Total Volume"
    pvWorksheet.Cells(3, iMaxCol + 1) = sMaxVolTicker
    pvWorksheet.Cells(3, iMaxCol + 2) = dMaxVol
    
    'Expand columns so all data is visible
    pvWorksheet.Columns().AutoFit
End Sub

Sub Stock_Data_All()
    Dim vWorkbook As Workbook
    Dim vWorksheet As Worksheet
    Dim iWorksheetLoop As Integer
    
    Set vWorkbook = ActiveWorkbook
    
    'Loop through all sheets in current workbook
    'Call the subroutine that handles an individual sheet
    For iWorksheetLoop = 1 To ActiveWorkbook.Worksheets.Count
        Set vWorksheet = ActiveWorkbook.Worksheets(iWorksheetLoop)
        Call Stock_Data_Sheet(vWorksheet)
    Next iWorksheetLoop
End Sub
