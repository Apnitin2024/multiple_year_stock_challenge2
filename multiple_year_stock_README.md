VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub mutlyiple_year_stock_data()
    Dim ws As Worksheet

    For Each ws In Worksheets
        ws.Activate
        Debug.Print ws.Name

        ws.Range("K1").Value = "Ticker"
        ws.Range("L1").Value = "Quaterly change"
        ws.Range("M1").Value = "Percentage change"
        ws.Range("N1").Value = "Total stock volume"
                ws.Range("Q2").Value = "Greatest % Increase"
        ws.Range("Q3").Value = "Greatest % Decrease"
        ws.Range("Q4").Value = "Greatest Total Volume"
        ws.Range("R1").Value = "Ticker"
        ws.Range("S1").Value = "Value"
        
    Next ws
    
    
End Sub

Sub Summary_table()
    On Error Resume Next

    Dim ws As Worksheet
    Dim ticker_name As String
    Dim Ticker_Total As Double
    Dim open_price As Double
    Dim close_price As Double
    Dim quarterly_change As Double
    Dim Percent_change As Double
    Dim i As Long
    Dim lastrow As Long
    Dim Summary_table_row As Integer
    Dim startRow As Long

    For Each ws In Worksheets
        ' Determining the Last Row
        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' Initialize the summary_table_row starting row
        Summary_table_row = 2

        ' Initialize variables (ticker_name)
        ticker_name = ws.Cells(2, 1).Value
        startRow = 2
    
        'Intializing the integer
        For i = 2 To lastrow
            
            ' Accumulate the total volume (ticker_total and ticker_name)
            Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value

            ' Check if we are still within the same Ticker, if not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Or i = lastrow Then
                
                ' Set the ticker name
                ticker_name = ws.Cells(i, 1).Value

                ' Get the open price from the first row of the quarter
                open_price = ws.Cells(startRow, 3).Value

                ' Get the close price from the last row of the quarter
                close_price = ws.Cells(i, 6).Value

                ' Calculate the quarterly change and percent change
                quarterly_change = close_price - open_price
                Percent_change = Round(((quarterly_change / open_price) * 100), 2)

                ' Print the Ticker name in the Summary Table
                ws.Range("K" & Summary_table_row).Value = ticker_name

                ' Print the quarterly change in the Summary Table
                ws.Range("L" & Summary_table_row).Value = quarterly_change

                ' Print the percent change in the Summary Table
                ws.Range("M" & Summary_table_row).Value = Percent_change & "%"

                ' Print the Volume of stocks to the Summary Table
                ws.Range("N" & Summary_table_row).Value = Ticker_Total

                ' Conditional color formatting for quarterly_change
                If ws.Range("L" & Summary_table_row).Value > 0 Then
                    ws.Range("L" & Summary_table_row).Interior.ColorIndex = 4
                ElseIf ws.Range("L" & Summary_table_row).Value < 0 Then
                    ws.Range("L" & Summary_table_row).Interior.ColorIndex = 3
                Else
                    ws.Range("L" & Summary_table_row).Interior.ColorIndex = 0
                End If

                ' Add one to the summary table row
                Summary_table_row = Summary_table_row + 1

                ' Reset the Ticker_Total and startRow for the next ticker
                Ticker_Total = 0
                startRow = i + 1
            End If
        Next i
        
        ' Code to find the max and min
        Dim maxPercentChange As Double
        Dim minPercentChange As Double

        maxPercentChange = WorksheetFunction.Max(ws.Range("M2:M" & Summary_table_row - 1))
        minPercentChange = WorksheetFunction.Min(ws.Range("M2:M" & Summary_table_row - 1))

        ' Find the greatest total volume
        Dim maxVolume As Double
        maxVolume = WorksheetFunction.Max(ws.Range("N2:N" & Summary_table_row - 1))

        ' Print the max and min percentage change and the greatest total volume in the Summary Table
        ws.Range("S2").Value = maxPercentChange
        ws.Range("S3").Value = minPercentChange
        ws.Range("S4").Value = maxVolume

        Dim maxTicker As String
        Dim minTicker As String
        Dim max_volume As String

        ' Find the tickers corresponding to the max and min percentage changes
        For i = 2 To Summary_table_row - 1
            If ws.Cells(i, 13).Value = maxPercentChange Then
                maxTicker = ws.Cells(i, 11).Value
            End If
            
            If ws.Cells(i, 13).Value = minPercentChange Then
                minTicker = ws.Cells(i, 11).Value
            End If
            
            If ws.Cells(i, 14).Value = maxVolume Then
                max_volumeticker = ws.Cells(i, 11).Value
            End If
        Next i

        ' Print the max and min percentage change and their corresponding tickers in cells R2 & R3
        ws.Range("R2").Value = maxTicker
        ws.Range("R3").Value = minTicker
        ws.Range("R4").Value = max_volumeticker
        
    Next ws
    
    End Sub
    



         
  
        

        
        
        
        











