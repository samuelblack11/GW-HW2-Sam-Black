Sub hw2Final()

 ' Loop through all of the worksheets in the active workbook.
For Each ws In Worksheets
Dim worksheetName As String

lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
worksheetName = ws.Name

' Set an initial variable for holding the ticker symbol
Dim ticker As String

' Set an initial variable for holding the total per ticker
Dim sumVolume As LongLong
sumVolume = 0

' Set an initial variable for holding the total per credit card brand
Dim iTickerCount As Integer
iTickerCount = 2

' Create variable for year's closing price
Dim closePrice As Double
Dim openPrice As Double
Dim annualPerformance As Double
Dim percentPerformance As Double

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Total Stock Volume"
ws.Cells(1, 11).Value = "Annual Performance"
ws.Cells(1, 12).Value = "Percent Peformance"

' Loop through all tickers per sheet
   For I = 2 To lastRow
    
    If ws.Cells(I, 1).Value <> ws.Cells(I - 1, 1) Then
    openPrice = ws.Cells(I, 3).Value

End If

' Starting at A2, for column A, iTicker = the ticker symbol as String
    If ws.Cells(I, 1).Value <> ws.Cells(I + 1, 1) Then

    ' Set ticker name
    ticker = ws.Cells(I, 1).Value
    closePrice = ws.Cells(I, 6).Value
    
    ' Add to sumVolume
    sumVolume = ws.Cells(I, 7).Value + sumVolume

    ' Print Ticker in Summary Table
    ws.Cells(iTickerCount, 9).Value = ws.Cells(I, 1).Value

    ' Print Total Volume for Ticker in Summary Table
   ws.Cells(iTickerCount, 10).Value = sumVolume
    
     annualPerformance = closePrice - openPrice
    ws.Cells(iTickerCount, 11).Value = annualPerformance
    
    ' Use Conditonal Formatting to display postive change in green and negative change in red
    If annualPerformance < 0 Then
    ws.Cells(iTickerCount, 11).Interior.ColorIndex = 3
    Else
    ws.Cells(iTickerCount, 11).Interior.ColorIndex = 4
    End If
    
    ' Calculate % Performance
    
    If openPrice = 0 Then
    percentPerformance = 0
    
    Else
    percentPerformance = (annualPerformance / openPrice)
    ' Print % Performance in Cell
   ws.Cells(iTickerCount, 12).Value = percentPerformance

   End If

    ' Add one to summary table row
    iTickerCount = iTickerCount + 1
    sumVolume = 0

    Else

    sumVolume = sumVolume + ws.Cells(I, 7).Value

    End If
    Next I

   ws.Range("L1:L10000").NumberFormat = "0.00%"
    
    ws.Cells(1, 14).Value = "Metric"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    
Dim range1 As Range
Dim range2 As Range
Dim max As Double
Dim min As Double
Dim highestVolume As LongLong
Dim e As String
Dim f As String
Dim g As String

'Set range from which to determine smallest value
Set range1 = ws.Range("L2:L5000")
Set range2 = ws.Range("J2:J5000")

' Find Max of range in loop through Cells(a,12).Value
max = Application.WorksheetFunction.max(range1)
ws.Cells(2, 16).Value = max
ws.Range("P2").NumberFormat = "0.00%"
' Left Lookup for Ticker Symbol
e = Application.WorksheetFunction.Index(ws.Range("I2:I5000"), Application.WorksheetFunction.Match(max, ws.Range("L2:L5000"), 0))
ws.Cells(2, 15).Value = e

' Find Min of range in loop through Cells(a,12).Value
min = Application.WorksheetFunction.min(range1)
ws.Cells(3, 16).Value = min
ws.Range("P3").NumberFormat = "0.00%"
' Left Lookup for Ticker Symbol
f = Application.WorksheetFunction.Index(ws.Range("I2:I5000"), Application.WorksheetFunction.Match(min, ws.Range("L2:L5000"), 0))
ws.Cells(3, 15).Value = f

highestVolume = Application.WorksheetFunction.max(range2)
ws.Cells(4, 16).Value = highestVolume
ws.Range("J2:J5000").NumberFormat = "0"
g = Application.WorksheetFunction.Index(ws.Range("I2:I5000"), Application.WorksheetFunction.Match(highestVolume, ws.Range("J2:J5000"), 0))
ws.Cells(4, 15).Value = g

ws.Columns.AutoFit

  MsgBox ws.Name
     Next ws
     
End Sub


