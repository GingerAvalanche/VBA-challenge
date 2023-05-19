Sub stock_summary_optimized():
    For Each sheet In Worksheets
        ' Init basic header stuff
        sheet.Range("I1").Value = "Ticker"
        sheet.Range("J1").Value = "Yearly Change"
        sheet.Range("K1").Value = "Percent Change"
        sheet.Range("K:K").NumberFormat = "0.00%"
        sheet.Range("L1").Value = "Total Stock Volume"
        sheet.Range("P1").Value = "Ticker"
        sheet.Range("Q1").Value = "Value"
        sheet.Range("O2").Value = "Greatest % Increase"
        sheet.Range("O3").Value = "Greatest % Decrease"
        sheet.Range("O4").Value = "Greatest Total Volume"
        sheet.Range("Q2:Q3").NumberFormat = "0.00%"
        
        ' Init variables for summary
        ' Explicitly set ticker_count to 0, as the Dim statement seems not to do it
        '  on subsequent iterations? And it's the only dependent variable
        Dim ticker As String
        Dim yearly_change As Double
        Dim year_start_value As Double
        year_start_value = sheet.Cells(2, 3).Value
        Dim year_end_value As Double
        Dim row, first_row, last_row As Long
        row = 2
        first_row = 2
        Dim ticker_count As Integer
        ticker_count = 2
        
        Do While sheet.Cells(row, 1).Value <> ""
            ticker = sheet.Cells(row, 1).Value
            
            If Not (sheet.Cells(row + 1, 1).Value = ticker) Then
                year_end_value = sheet.Cells(row, 6).Value
                yearly_change = year_end_value - year_start_value
                sheet.Cells(ticker_count, 9).Value = ticker
                sheet.Cells(ticker_count, 10).Value = yearly_change
                sheet.Cells(ticker_count, 11).Value = yearly_change / year_start_value
                last_row = row
                sheet.Cells(ticker_count, 12).Value = Application.WorksheetFunction.Sum(sheet.Range("G" & first_row & ":G" & last_row))
                first_row = row + 1
                year_start_value = sheet.Cells(row + 1, 3).Value
                ticker_count = ticker_count + 1
            End If
            
            row = row + 1
        Loop
        
        ' Format the yearly and percentage changed ranges
        Dim range_str As String
        Dim format_range1, format_range2 As Range
        Set format_range1 = sheet.Range("J2:J" & Trim(Str(row)))
        Set format_range2 = sheet.Range("K2:K" & Trim(Str(row)))
        
        format_range1.FormatConditions.Delete
        format_range1.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:=0
        format_range1.FormatConditions(1).Interior.Color = RGB(255, 0, 0)
        format_range1.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:=0
        format_range1.FormatConditions(2).Interior.Color = RGB(0, 255, 0)
        format_range2.FormatConditions.Delete
        format_range2.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:=0
        format_range2.FormatConditions(1).Interior.Color = RGB(255, 0, 0)
        format_range2.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:=0
        format_range2.FormatConditions(2).Interior.Color = RGB(0, 255, 0)
        
        ' Retrieve the most-positive and most-negative % changes, and highest volume
        ' Explicitly set storage variables to 0, as testing shows the Dim statements
        '  don't do that for you in VBA, for some reason
        row = 2
        Dim change, increase, decrease As Double
        Dim vol_change, volume As Single
        increase = 0
        decrease = 0
        volume = 0
        Do While sheet.Cells(row, 9).Value <> ""
            ticker = sheet.Cells(row, 9).Value
            change = sheet.Cells(row, 11).Value
            vol_change = sheet.Cells(row, 12).Value
            
            If increase < change Then
                increase = change
                sheet.Range("P2").Value = ticker
                sheet.Range("Q2").Value = increase
            End If
            
            If decrease > change Then
                decrease = change
                sheet.Range("P3").Value = ticker
                sheet.Range("Q3").Value = decrease
            End If
            
            If volume < vol_change Then
                volume = vol_change
                sheet.Range("P4").Value = ticker
                sheet.Range("Q4").Value = volume
            End If
            
            row = row + 1
        Loop
        
        ' AutoFit data/header columns, ignoring ticker columns
        sheet.Columns("J:L").AutoFit
        sheet.Columns("O").AutoFit
        ' sheet.Columns("Q").AutoFit ' Ignore this one, the screenshots didn't make this one fit
    Next sheet
End Sub
