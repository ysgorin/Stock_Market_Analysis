Attribute VB_Name = "Module1"
Sub StockAnalysis()

' Declare variables to retrieve: ticker symbol, total stock volume (tsv), year open price, and year close price.
Dim Symbol As String
Dim TSV As LongLong
Dim YearOpen As Double
Dim YearClose As Double

' Declare variables to store calculated yearly price change and percent change.
Dim YearlyChange As Double
Dim PercentChange As Double

' Declare variables to store greatest percent increase, decrease, and total volume and their respective symbols.
Dim GIncr As Double
Dim GDecr As Double
Dim GIncrSymbol As String
Dim GDecrSymbol As String
Dim GTV As LongLong
Dim GTVSymbol As String

' Declare variable for row counter for output to summary table.
Dim RowCounter As Long

Dim ws As Worksheet
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

For Each ws In Worksheets
  
    RowCounter = 2
    StockVolume = 0

    ' Set up ticker symbol summary table on each shset
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ' Set up additional summary table on each sheet
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"

    ' Format titles bold.
    ws.Range("I1:L1,O2:O4,P1:Q1").Font.Bold = True
    
    ' Create a script that loops through all the stocks for one year
    For i = 2 To LastRow

        ' Make a running sum of total stock volume for ticker symbol.
        TSV = TSV + ws.Cells(i, 7).Value

        ' Identify first record for ticker symbol.
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then

            ' Record opening price for ticker symbol
            YearOpen = ws.Cells(i, 3).Value

        End If

        ' Identify last record for each ticker symbol
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
            ' Record ticker symbol.
            Symbol = ws.Cells(i, 1).Value

            ' Output ticker symbol to summary table.
            ws.Range("I" & RowCounter).Value = Symbol
  
            ' Record closing price for ticker symbol.
            YearClose = ws.Cells(i, 6).Value
    
            ' Calculate yearly price change.
            YearlyChange = YearClose - YearOpen

            ' Output price change to summary table.
            ws.Range("J" & RowCounter).Value = YearlyChange
    
            ' Format positive price change in green and negative in red.
            If YearlyChange > 0 Then
                ws.Range("J" & RowCounter).Interior.ColorIndex = 4
    
            ElseIf YearlyChange < 0 Then
                ws.Range("J" & RowCounter).Interior.ColorIndex = 3
    
            End If
    
            ' Calculate yearly percent change.
            PercentChange = (YearlyChange / YearOpen)
    
            ' Output formatted percent change to summary table.
            ws.Range("K" & RowCounter).Value = FormatPercent(PercentChange, 2, vbTrue, vbFalse)
            
            ' Format positive percent change in green and negative in red.
            If PercentChange > 0 Then
                ws.Range("K" & RowCounter).Interior.ColorIndex = 4
    
            ElseIf PercentChange < 0 Then
                ws.Range("K" & RowCounter).Interior.ColorIndex = 3
    
            End If
            
            ' Output total stock volume to summary table.
            ws.Range("L" & RowCounter).Value = TSV
    
            ' Reset total stock volume variable for next symbol.
            TSV = 0

            ' Add to row counter for next summary table output.
            RowCounter = RowCounter + 1

        End If

    Next i
    
    ' Set greatest variables at zero.
    GDecr = 0
    GIncr = 0
    GTV = 0

    ' Create a script that loops through the created summary table.
    For i = 2 To LastRow

        ' Identify the greatest percent decrease and its ticker symbol.
        If ws.Cells(i, 11).Value < GDecr Then
            GDecr = ws.Cells(i, 11).Value
            GDecrSymbol = ws.Cells(i, 9).Value

        ' Identify the greatest percent increase and its ticker symbol.
        ElseIf ws.Cells(i, 11).Value > GIncr Then
            GIncr = ws.Cells(i, 11).Value
            GIncrSymbol = ws.Cells(i, 9).Value

        End If

        ' Identify the greatest total volume and its ticker symbol.
        If ws.Cells(i, 12).Value > GTV Then
            GTV = ws.Cells(i, 12).Value
            GTVSymbol = ws.Cells(i, 9).Value

        End If

    Next i

    ' Output results to additional summary table.
    ws.Range("P2").Value = GIncrSymbol
    ws.Range("Q2").Value = FormatPercent(GIncr, 2, vbTrue, vbFalse)
    ws.Range("P3").Value = GDecrSymbol
    ws.Range("Q3").Value = FormatPercent(GDecr, 2, vbTrue, vbFalse)
    ws.Range("P4").Value = GTVSymbol
    ws.Range("Q4").Value = GTV
    ws.Columns("A:Q").AutoFit
       
Next ws

End Sub
