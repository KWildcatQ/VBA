Sub StockChallenge()
    
'Declare Variables
Dim Tickers As Integer
Dim Cumulative As Double
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim YearlyChange As Double

'Column labels

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'Column labels on Multiple Sheets
Dim wsSheet As Worksheet

For Each wsSheet In ThisWorkbook.Worksheets
wsSheet.Range("A1:L1").Value = Worksheets("2014").Range("A1:L1").Value
Next wsSheet

'Compile Tickers
Dim r As Long
r = Cells(Rows.Count, 1).End(xlUp).Row

Cumulative = 0
Tickers = 0

' Set open price for first ticker
OpenPrice = Cells(2, 3).Value

For i = 2 To r

    
    'Percent Change of Yearly Open and Closing Price
    Cells(Tickers + 2, 11).Value = PercentChange
            
   'Total Volume
   Cumulative = Cells(i, 7).Value + Cumulative
      
    'Comparison to row above cell
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        'New symbol has been found
        Tickers = Tickers + 1
        Cells(Tickers + 1, 9).Value = Cells(i, 1).Value
        Cells(Tickers + 1, 12).Value = Cumulative
        
        ' Set close price for current ticker
        ClosePrice = Cells(i, 6).Value
        
        'YearlyChange
        YearlyChange = ClosePrice - OpenPrice
        Cells(Tickers + 1, 10).Value = YearlyChange
        
        'Set Conditional Formatting
        If Cells(Tickers + 1, 10).Value >= 0 Then
        Cells(Tickers + 1, 10).Interior.ColorIndex = 4
        Else
        Cells(Tickers + 1, 10).Interior.ColorIndex = 3
        End If
            
        If OpenPrice = 0 Then
        PercentChange = 0
        Else
        PercentChange = (ClosePrice - OpenPrice) / OpenPrice
        End If

        ' Set Summary Percentage Change
        Cells(Tickers + 1, 11).Value = PercentChange

        ' Set open price for next ticker
        OpenPrice = Cells(i + 1, 3).Value
    
        Else
        If OpenPrice = 0 Then
        OpenPrice = Cells(i, 3).Value
        
        End If
                
        Cumulative = 0
        
        End If

 Next i
  
    'Formatting Percent Change Column as Percent
    Columns(11).NumberFormat = "0.00%"
            
End Sub

