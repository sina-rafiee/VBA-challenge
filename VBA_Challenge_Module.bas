Attribute VB_Name = "Module1"


Sub StockSort()

For Sheet = 1 To ThisWorkbook.Worksheets.Count
Worksheets(Sheet).Activate

'Establishes most variables.

Dim TickerCount As Integer
Dim YROpeningVal
Dim YRClosingVal
Dim EndofTable As Long

YROpeningVal = 0
YRClosingVal = 0
TickerCount = 2
EndofTable = Range("A1").End(xlDown).Row

'Adds a temporary associable end to the table.

If Cells(EndofTable, 1).Value <> 0 Then
    For Col = 1 To 7
        Cells(EndofTable + 1, Col).Value = 0
    Next Col
End If

'Runs through the initial table, creating a new one based on parameters set.

For Row = 2 To (EndofTable + 1)

    If IsNumeric(Cells(Row, 7).Value) = True Then
        SumVolume = SumVolume + Cells(Row - 1, 7).Value
    End If
       
    If Cells(Row, 1).Value <> Cells(Row - 1, 1).Value = True Then
           
        If Cells(Row, 1).Value <> 0 Then
            Cells(TickerCount, 9).Value = Cells(Row, 1).Value
        End If
           
        If IsNumeric(Cells(Row - 1, 6).Value) = True Then
            YRClosingVal = Cells(Row - 1, 6).Value
            Cells(TickerCount - 1, 10).Value = YRClosingVal - YROpeningVal
            Cells(TickerCount - 1, 11).Value = YRClosingVal / YROpeningVal - 1
            Cells(TickerCount - 1, 12).Value = SumVolume
        End If
           
        SumVolume = 0
        TickerCount = TickerCount + 1
        YROpeningVal = Cells(Row, 3).Value
    End If

Next Row

' Cleans up the End of the original Table.
   
For Col = 1 To 7
    Cells(EndofTable + 1, Col).Clear
Next Col

'Sets the Headers for the new table.

Cells(1, 9).Value = "Tickers"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

' Changes Column Width to be more appropriate

Range("I:I").ColumnWidth = 9
Range("J:K").ColumnWidth = 15
Range("L:L").ColumnWidth = 19

Range("K:K").NumberFormat = "0.00%"
Range("L:L").NumberFormat = "###,###,###,###,###,###"
'Sets Conditional Formatting For Column J

Dim LTFormat As FormatCondition
Dim GTFormat As FormatCondition
Dim YearlyValues As Range
Set YearlyValues = Range("J2:J" & TickerCount)

YearlyValues.FormatConditions.Delete
Set LTFormat = YearlyValues.FormatConditions.Add(xlCellValue, xlLess, "0")
Set GTFormat = YearlyValues.FormatConditions.Add(xlCellValue, xlGreater, "0")

With LTFormat
    .Interior.Color = RGB(275, 160, 180)
End With
With GTFormat
    .Interior.Color = RGB(180, 255, 130)
End With
   
'Resets Variables
SumVolume = Empty
TickerCount = Empty
YROpeningVal = Empty
YRClosingVal = Empty
EndofTable = Empty
   

' BONUS


'Sets all Variables needed while reading through Table2
Dim GI As Double, GD As Double, GV As Double
Dim GIT As String, GDT As String, GVT As String
   
Dim EndofTable2 As Long
EndofTable2 = Range("I1").End(xlDown).Row

'Reads Table2 for Values needed.
For Table2Row = 2 To EndofTable2
    If Cells(Table2Row, 11).Value > GI Then
    GI = Cells(Table2Row, 11).Value
    GIT = Cells(Table2Row, 9).Value
    End If
    
    If Cells(Table2Row, 11).Value < GD Then
    GD = Cells(Table2Row, 11).Value
    GDT = Cells(Table2Row, 9).Value
    End If
    
    If Cells(Table2Row, 12).Value > GV Then
    GV = Cells(Table2Row, 12).Value
    GVT = Cells(Table2Row, 9).Value
    End If
Next Table2Row

'Places all Values
Cells(2, 16).Value = GIT
Cells(3, 16).Value = GDT
Cells(4, 16).Value = GVT

Cells(2, 17).Value = GI
Cells(3, 17).Value = GD
Cells(4, 17).Value = GV

' Set Headers for both Rows and Columns of the 3rd table.
   
' Row Headers
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
   
' Column Headers
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

' Change Column Widths for 3rd Table
Range("O:O").ColumnWidth = 20
Range("P:Q").ColumnWidth = 17

'Formats Values
Range("Q2:Q3").NumberFormat = "0.00%"
Range("Q4").NumberFormat = "###,###,###,###,###,###"

'Resets Variables
GI = Empty
GD = Empty
GV = Empty
GIT = Empty
GDT = Empty
GVT = Empty

Next Sheet

MsgBox ("Sub Complete!")

End Sub
