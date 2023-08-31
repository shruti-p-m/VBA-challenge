Attribute VB_Name = "Module1"
Sub stock_analysis()

'declaring variables for columns

Dim Ticker As String
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total_Volume As LongLong
Dim Opening As Double
Dim Closing As Double

'to loop throught all the sheets

For Each ws In Worksheets
Dim lastRow As LongLong
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'adding columns headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Value"

    'declaring variable for for loop
    Dim i As LongLong
    Dim j As Integer
    'initializing ticker counter before loop
    j = 1
    'getting first ticker's opening value
    Opening = ws.Cells(2, 3).Value
    
    'intializing total volume
    Total_Volume = 0
    
    For i = 2 To lastRow
    'sums the stock value
    Total_Volume = Total_Volume + ws.Cells(i, 7).Value
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            'keeps track of what number ticker
            j = j + 1
            'assigning value for ticker, and adding the value in the new ticker column
            Ticker = ws.Cells(i, 1).Value
            ws.Cells(j, 9).Value = Ticker
        
            'gathering closing value of the pervious ticker
        
            Closing = ws.Cells(i, 6).Value
            'calculating yearly change and adding it to new column
            Yearly_Change = Closing - Opening
            ws.Cells(j, 10).Value = Yearly_Change
        
            'Assigning color to yearly change
            If Yearly_Change < 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 3
        
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 4
        
            End If
            Percent_Change = Yearly_Change / Opening
            ' https://www.mrexcel.com/board/threads/vba-change-number-of-decimal-places-of-a-percentage.521221/
            'formats the yearly change to a percent with 0.00% format
            ws.Cells(j, 11).NumberFormat = "0.00%"
            ws.Cells(j, 11).Value = Percent_Change
            
            'assigns the total volume in the column when the ticker changes
            ws.Cells(j, 12).Value = Total_Volume
            
            'resets the total volume for the next ticker
            Total_Volume = 0
            
            'gets opening value for the next ticker
            Opening = ws.Cells(i + 1, 3).Value
        End If
    Next i

    'variable to find greatest increase and decrease number and ticker
    Dim compare_increase As Double
    Dim compare_decrease As Double
    Dim compare_volume As LongLong
    Dim ticker_increase As String
    Dim ticker_decrease As String
    Dim ticker_volume As String
    compare_increase = ws.Cells(2, 11).Value
    compare_decrease = ws.Cells(2, 11).Value
    compare_volume = ws.Cells(2, 12).Value
    
    Dim ticker_number As Integer
    ticker_number = j + 1
    
    'find the greatest percent change
    For i = 3 To ticker_number
        If ws.Cells(i, 11).Value > compare_increase Then
            compare_increase = ws.Cells(i, 11).Value
            ticker_increase = ws.Cells(i, 9).Value
        End If
    Next i
    
    'find the smallest percent change
    For i = 3 To ticker_number
        If ws.Cells(i, 12).Value > compare_volume Then
            compare_volume = ws.Cells(i, 12).Value
            ticker_volume = ws.Cells(i, 9).Value
        End If
    Next i
    
    'find the greatest total stock volume
     For i = 3 To (j + 1)
        If ws.Cells(i, 11).Value < compare_decrease Then
            compare_decrease = ws.Cells(i, 11).Value
            ticker_decrease = ws.Cells(i, 9).Value
        End If
    Next i
    
    'formatting and putting the greatest percent change in the sheet
    ws.Range("P2").Value = ticker_increase
    ws.Range("Q2").Value = compare_increase
    ws.Range("Q2").NumberFormat = "0.00%"
    
    'formatting and putting the smallest percent change in the sheet
    ws.Range("P3").Value = ticker_decrease
    ws.Range("Q3").Value = compare_decrease
    ws.Range("Q3").NumberFormat = "0.00%"
    
    'formatting and putting the greatest total stock volume in the sheet
    ws.Range("P4").Value = ticker_volume
    ws.Range("Q4").Value = compare_volume
    ws.Range("Q4").NumberFormat = Number
    
    
Next ws
End Sub
