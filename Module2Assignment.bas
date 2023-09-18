Attribute VB_Name = "Module1"
Sub sort()

For Each ws In Worksheets

Dim WorksheetName As String
WorksheetName = ws.Name

'defining all used variables
Dim ticker As String
Dim tickerbig As String

Dim start As Double
Dim last As Double

Dim vol As Double
Dim volbig As Double

Dim holder As Integer

Dim great As Double
Dim least As Double

Dim greattick As String
Dim leasttick As String


'resetting values to proper starting values for each ws
great = 0
least = 0
start = 0
last = 0
vol = 0
volbig = 0
holder = 2
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


'setting up columns for processed data
ws.Range("J" & 1).Value = "Ticker"
ws.Range("K" & 1).Value = "Yearly Change"
ws.Range("L" & 1).Value = "Percent Change"
ws.Range("M" & 1).Value = "Total Stock Volume"


For i = 2 To lastrow
    'if the current line is a differnt ticker from the prior line, mark start value and begin keeping track of volume
    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
        start = ws.Cells(i, 3).Value
        vol = vol + ws.Cells(i, 7).Value
    'else if current line is differnt from the following line, mark end value, calculate total volume, as well as last stock value
    ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ticker = ws.Cells(i, 1).Value
        vol = vol + ws.Cells(i, 7).Value
        'if statment to keep track of biggest volume for the year
        If vol > volbig Then
                volbig = vol
                tickerbig = ticker
        End If
        
        last = ws.Cells(i, 6).Value
        
        'place yearly totals in correct columns of yearly calculations
        ws.Range("J" & holder).Value = ticker
        ws.Range("K" & holder).Value = (last - start)
        ws.Range("L" & holder).Value = ((last - start) / start)
        
            ' adjusts color for positive/negative change for the year
            If ws.Range("K" & holder).Value > 0 Then
                ws.Range("K" & holder).Interior.ColorIndex = 4
            Else
                ws.Range("K" & holder).Interior.ColorIndex = 3
            End If
        
        ws.Range("M" & holder).Value = vol
        'reset volume for each itteration and add one to place holder for final data
        vol = 0
        holder = holder + 1
    Else
        'if same ticker, just keep track of volume
        vol = vol + ws.Cells(i, 7).Value
    End If
Next i

'itteration goes through proccesed data, looking for biggest increase,decrease and total volume
lastrow2 = ws.Cells(Rows.Count, 12).End(xlUp).Row
 For j = 2 To lastrow2
    If ws.Range("L" & j).Value > great Then
        great = ws.Range("L" & j).Value
        greattick = ws.Range("J" & j).Value
    ElseIf ws.Range("L" & j).Value < least Then
        least = ws.Range("L" & j).Value
        leasttick = ws.Range("J" & j).Value
    End If
Next j

'setting up columns for final data
ws.Range("O" & 1).Value = "Greatest % Increase"
ws.Range("O" & 2).Value = "Greatest % Decrease"
ws.Range("O" & 3).Value = "Greatest Total Volume"
ws.Range("P" & 1).Value = greattick
ws.Range("P" & 2).Value = leasttick
ws.Range("P" & 3).Value = tickerbig
ws.Range("Q" & 1).Value = great
ws.Range("Q" & 2).Value = least
ws.Range("Q" & 3).Value = volbig

'change format for percents, and autofit columns
ws.Range("L2:L" & lastrow).NumberFormat = "0.00%"
ws.Cells(1, 17).NumberFormat = "0.00%"
ws.Cells(2, 17).NumberFormat = "0.00%"

ws.Columns("A:Q").AutoFit

Next ws

End Sub

