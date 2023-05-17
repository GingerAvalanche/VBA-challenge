Sub stock_summary():
    Dim sheet As Worksheet
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
        ReDim tickers(0) As String
        Dim ticker As String
        Dim ticker_loc As Integer
        ReDim ticker_volume(0) As Double
        Dim yearly_change As Double
        ReDim year_start_date(0) As Long
        ReDim year_start_value(0) As Double
        ReDim year_end_date(0) As Long
        ReDim year_end_value(0) As Double
        Dim row As Long
        row = 2
        Dim ticker_count As Integer
        ticker_count = 0
        
        Do While sheet.Cells(row, 1).Value <> ""
            ticker = sheet.Cells(row, 1).Value
            ticker_loc = IndexOf(ticker, tickers)
            
            If ticker_loc = -1 Then
                ' If arrays aren't big enough, double their size
                If UBound(tickers) < ticker_count Then
                    ReDim Preserve tickers(ticker_count * 2)
                    ReDim Preserve ticker_volume(ticker_count * 2)
                    ReDim Preserve year_start_date(ticker_count * 2)
                    ReDim Preserve year_start_value(ticker_count * 2)
                    ReDim Preserve year_end_date(ticker_count * 2)
                    ReDim Preserve year_end_value(ticker_count * 2)
                End If
                ' Init non-default values for new ticker
                ticker_loc = ticker_count
                tickers(ticker_loc) = ticker
                year_start_date(ticker_loc) = 99999999
                ticker_count = ticker_count + 1
            End If
            
            ' Add volume to running tally for this ticker
            ticker_volume(ticker_loc) = ticker_volume(ticker_loc) + sheet.Cells(row, 7).Value
            
            ' Store open price if this date is earlier than current open price date
            If year_start_date(ticker_loc) > sheet.Cells(row, 2).Value Then
                year_start_date(ticker_loc) = sheet.Cells(row, 2).Value
                year_start_value(ticker_loc) = sheet.Cells(row, 3).Value
            End If
            
            ' Store close price if this date is later than current close price date
            If year_end_date(ticker_loc) < sheet.Cells(row, 2).Value Then
                year_end_date(ticker_loc) = sheet.Cells(row, 2).Value
                year_end_value(ticker_loc) = sheet.Cells(row, 6).Value
            End If
            
            row = row + 1
        Loop
        
        ' Set the summary data for each ticker
        ' Make sure to check for empty tickers, as the arrays are fixed-length
        '  and probably contain empty values
        For i = 0 To UBound(tickers)
            If tickers(i) = "" Then
                Exit For
            End If
            row = i + 2
            sheet.Cells(row, 9).Value = tickers(i)
            yearly_change = year_end_value(i) - year_start_value(i)
            sheet.Cells(row, 10).Value = yearly_change
            sheet.Cells(row, 11).Value = (yearly_change / year_start_value(i))
            sheet.Cells(row, 12).Value = ticker_volume(i)
        Next i
        
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

Function IndexOf(search As String, arr() As String) As Integer
    Dim found As Boolean
    Dim loc As Integer
    loc = 0
    
    Do While loc <= UBound(arr)
        If arr(loc) = search Then
            found = True
            IndexOf = loc
            Exit Do
        End If
        loc = loc + 1
    Loop
    
    If Not (found) Then
        IndexOf = -1
    End If
End Function
