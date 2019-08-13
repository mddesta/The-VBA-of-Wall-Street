Sub Stock():
On Error Resume Next
'Loop through all worksheets in the workbook
num = ActiveWorkbook.Worksheets.Count
For j = 1 To num
Dim SheetName As String
SheetName = ActiveWorkbook.Worksheets(j).Name
Sheets(SheetName).Activate

'Summuraize for each stock.
Dim lastrow As Long
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Dim amount As Double
Dim ticker As Double
Dim Opening As Double
Dim Closing As Double
Dim YC As Double
Dim PC As Double
Dim PCV As String

ticker = 1
amount = 0


    For i = 1 To lastrow
    'adding the volume for identical tickers
        If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
            amount = amount + Cells(i, 7).Value
        End If
    'calculating the summary per ticker
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Cells(ticker, 9).Value = Cells(i, 1).Value
            amount = amount + (Cells(i, 7).Value)
            Cells(ticker, 12).Value = amount
            Closing = Cells(i, 6).Value
            YC = Closing - Opening
            Cells(ticker, 10).Value = YC
            PC = (YC / Opening)
                If Opening <> 0 Then
                Cells(ticker, 13).Value = PC
                End If
            PCV = Format(PC, "Percent")
            Cells(ticker, 11).Value = PCV
            'resetting total volume to zero
            amount = 0
            ticker = ticker + 1
            Opening = Cells(i + 1, 3).Value
        End If
    Next i
    'creating the headers for the summary
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "value"
    Cells(2, 15).Value = "Greatest % increase"
    Cells(3, 15).Value = "Greatest % decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 13).Value = ""

    
    'formating the % change
    Dim greatest As Variant
    Dim lowest As Variant
    Dim Largest As Variant
    
    
    
    greatest = 0
    lowest = 0
    Largest = 0
    
    For a = 2 To ticker
        If Cells(a, 10).Value > 0 Then
        Cells(a, 11).Interior.ColorIndex = 4
        End If
        If Cells(a, 10).Value < 0 Then
        Cells(a, 11).Interior.ColorIndex = 3
        End If
    'calculating the largest
        If Cells(a, 12).Value > Largest Then
            Largest = Cells(a, 12).Value
            Cells(4, 16).Value = Cells(a, 9).Value
            Cells(4, 17).Value = (Cells(a, 12).Value)
        End If
     'Calculating the greatest and least %change
        If Cells(a, 13).Value > greatest Then
            greatest = Cells(a, 13).Value
            Cells(2, 16).Value = Cells(a, 9).Value
            Cells(2, 17).Value = Format(greatest, "Percent")
        End If
        If Cells(a, 13).Value < lowest Then
            lowest = Cells(a, 13).Value
            Cells(3, 16).Value = Cells(a, 9).Value
            Cells(3, 17).Value = Format(lowest, "Percent")
        End If
    Next a
'clear contents of column 13
Range("M1:M" & lastrow).ClearContents

Next j

End Sub





