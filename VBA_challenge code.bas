Attribute VB_Name = "Module1"
Sub year()

Dim Current As Worksheet

         For Each Current In Worksheets
         
    Dim LastRow As Long
    Dim Ticker As String
    Dim OpeningPrice As Double
    Dim YearlyChange As Double
    Dim PercentageChange As Double
    Dim OutputRow As Long
    
    LastRow = Current.Cells(Rows.Count, 1).End(xlUp).Row
    OutputRow = 2 ' Starting row for output
    StockTotal = 0
    
    ' Output headers
    Current.Range("I1").Value = "Ticker"
    Current.Range("J1").Value = "Yearly Change"
    Current.Range("K1").Value = "Percentage Change"
    
    ' Loop through the data
    OpeningPrice = Current.Cells(2, 3).Value
    Ticker = Current.Cells(2, 1).Value
    For i = 2 To LastRow
    StockTotal = StockTotal + Current.Cells(i, 7).Value
        If Ticker <> Current.Cells(i + 1, 1).Value Then
            ' New ticker, calculate yearly change and percentage change, and output
            ' If Ticker <> "" Then
                YearlyChange = Current.Cells(i, 6).Value - OpeningPrice
                PercentageChange = (YearlyChange / OpeningPrice)
                Current.Range("I" & OutputRow).Value = Ticker
                Current.Range("J" & OutputRow).Value = YearlyChange
                Current.Range("J" & OutputRow).NumberFormat = "0.00"
                Current.Range("K" & OutputRow).NumberFormat = "0.00%"
                ' Range.DisplayFormat.Interior.Color
                ' 3 is red
                ' 4 is green
            If YearlyChange > 0 Then
                Current.Range("J" & OutputRow).Interior.ColorIndex = 4
                
            ElseIf YearlyChange < 0 Then
                Current.Range("J" & OutputRow).Interior.ColorIndex = 3
                
                End If
                
                Current.Range("K" & OutputRow).Value = PercentageChange
                Current.Range("L" & OutputRow).Value = StockTotal
                OutputRow = OutputRow + 1
                
            ' End If
            StockTotal = 0
            Ticker = Current.Cells(i, 1).Value
            OpeningPrice = Current.Cells(i, 3).Value
        End If
    Next i
    
    ' Calculate and output the last ticker's yearly change and percentage change
    YearlyChange = Current.Cells(LastRow, 6).Value - OpeningPrice
    PercentageChange = (YearlyChange / OpeningPrice) * 100
    Current.Range("I" & OutputRow).Value = Ticker
    Current.Range("J" & OutputRow).Value = YearlyChange
    Current.Range("K" & OutputRow).Value = PercentageChange
    
    Next
    
End Sub
