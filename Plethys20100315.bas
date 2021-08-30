Attribute VB_Name = "PlethysNew"
Option Explicit
Dim total As Double
Sub Analyze()
Call prepare

Call check
Call delete
Call apneas
Call summary_data
End Sub


Sub check()
Worksheets(3).Activate
total = 0
Dim CellRef As Range
Dim last As Boolean
last = False

'Old
'Modify
'Set CellRef = Range("A26605:B26610")
'/Modify


'Commented 12/22

Dim i As Long
'
'Dim tRow As Boolean
'tRow = False
'Dim row2 As Long
'row2 = 2
'
'Do While Not tRow
'    If Cells(row2, 1) = "Times" Then
'        tRow = True
'    End If
'    row2 = row2 + 1
'Loop
'
Worksheets(3).Activate
Cells(1, 1).Select
Range(Selection, Selection.End(xlDown)).Select
Set CellRef = Range(Selection, Selection.End(xlToRight))


Dim row3 As Long
row3 = 1

Do While Worksheets(3).Cells(row3, 1) <> ""
    Worksheets(3).Cells(row3, 3).Value = "=RC[-1]-RC[-2]"
    row3 = row3 + 1
Loop

Worksheets(3).Cells(row3, 2).Value = "Total"
Worksheets(3).Cells(row3, 3).Value = "=sum(R[" & -1 * row3 + 1 & "]C:R[-1]C)"
total = Worksheets(3).Cells(row3, 3).Value


Dim j As Long
j = 2

Worksheets("Quiet Breathing Times").Activate

Do While (Not last)
    For i = 1 To CellRef.Rows.Count
        If ActiveSheet.Cells(j, 10).Value > CellRef.Cells(i, 1) And ActiveSheet.Cells(j, 10).Value < CellRef.Cells(i, 2) Then
            ActiveSheet.Cells(j, 11).Value = "y"
        End If
    Next i
    j = j + 1
    If ActiveSheet.Cells(j, 10) = "" Then
        last = True
    End If
Loop


End Sub

Sub delete()
Dim i As Long
Dim CellRef As Range

i = 2

Do While Cells(i, 8).Value <> ""
    If Cells(i, 11) = "" Then
        Cells(i, 11).Select
        Set CellRef = Range(Selection, Selection.End(xlDown)).Resize(Range(Selection, Selection.End(xlDown)).Rows.Count - 1)
        CellRef.Select
        Selection.EntireRow.delete
    Else
        i = i + 1
    End If
Loop

End Sub

Sub apneas()
Dim row As Long
row = 2

Do While Cells(row, 13) <> ""
    row = row + 1
Loop

row = row + 1

Cells(row - 1, 19).Value = "=2*average(R[" & 3 - row & "]C[0]:R[-1]C[0])"
Cells(row, 19).Value = "00:00:" & Cells(row - 1, 19).Value
Cells(row, 19).NumberFormat = "General"

Dim i As Long
Dim CellRef As Range

i = 3

Do While Cells(i, 8).Value <> ""
    If Cells(i, 9) > Cells(row, 19) Then
        Cells(i, 14).Value = "y"
    End If
    i = i + 1
Loop

Call sort

If Cells(row - 2, 14).Value = "y" Then
    Cells(row - 2, 14).Select
    If Cells(row - 3, 14).Value = "y" Then
        Range(Selection, Selection.End(xlUp)).EntireRow.Select
        Else
            Selection.EntireRow.Select
    End If
    Selection.Cut
    Worksheets("Apneas").Activate
    Worksheets("Apneas").Cells(2, 1).Select
    ActiveSheet.Paste
End If

Worksheets("Quiet Breathing Times").Activate
Cells(1, 1).EntireRow.Select
Selection.Copy
Worksheets("Apneas").Activate
Worksheets("Apneas").Cells(1, 1).Select
ActiveSheet.Paste

Worksheets("Quiet Breathing Times").Activate
Cells(2, 13).Select
Selection.End(xlDown).Select
Cells(ActiveCell.row + 1, ActiveCell.Column).Select
Range(Selection, Selection.End(xlDown)).Resize(Range(Selection, Selection.End(xlDown)).Rows.Count - 2).EntireRow.delete

Call sort_again


End Sub
Sub sort()
    Cells(2, 9).Select
    Range(Selection, Selection.End(xlDown)).EntireRow.Select
    Selection.sort Key1:=Range("I2"), Order1:=xlAscending, Header:= _
        xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
'    ActiveWorkbook.Worksheets("Quiet Breathing Times").sort.SortFields.Clear
'    ActiveWorkbook.Worksheets("Quiet Breathing Times").sort.SortFields.Add Key:=Range("I1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
'        xlSortNormal
'    With ActiveWorkbook.Worksheets("Quiet Breathing Times").sort
'        .SetRange Range(Selection, Selection.End(xlDown)).EntireRow
'        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
End Sub
Sub sort_again()
    Cells(1, 12).Select
    Range(Selection, Selection.End(xlDown)).EntireRow.Select
    Selection.sort Key1:=Range("L2"), Order1:=xlAscending, Header:= _
        xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
'    ActiveWorkbook.Worksheets("Quiet Breathing Times").sort.SortFields.Clear
'    ActiveWorkbook.Worksheets("Quiet Breathing Times").sort.SortFields.Add Key:=Range("L1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
'        xlSortNormal
'    With ActiveWorkbook.Worksheets("Quiet Breathing Times").sort
'        .SetRange Range(Selection, Selection.End(xlDown)).EntireRow
'        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
End Sub

Sub prepare()

    Dim CellRef As Range

    Sheets("WBP_Compensated1_Data").Activate
    Cells(1, 1).Select
    Sheets("WBP_Compensated1_Data").Select
    Sheets("WBP_Compensated1_Data").Copy After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Select
    Sheets(Sheets.Count).Name = "All Data with Gaps"
    
'New Stuff
    Range("I1").Select
    Selection.EntireColumn.Insert
    ActiveCell.FormulaR1C1 = "Gap Time"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]-(R[-1]C[-1]+(R[-1]C[5]+R[-1]C[6])/3600/24)"
    Range("H2").Select
    Set CellRef = Range(Selection, Selection.End(xlDown)).Offset(0, 1)
    Range("I2").Select
    Selection.AutoFill Destination:=CellRef
    
    Columns("I:I").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.NumberFormat = "s.000"
        
'/New Stuff

'Irregularity
    Range("AE1").Select
    ActiveCell.FormulaR1C1 = "Irr"
    Range("AE2").Select
    ActiveCell.FormulaR1C1 = "=ABS((RC[-17]+RC[-16])-(R[-1]C[-17]+R[-1]C[-16]))/(R[-1]C[-17]+R[-1]C[-16])"
    Range("AD2").Select
    Set CellRef = Range(Selection, Selection.End(xlDown)).Offset(0, 1)
    Range("AE2").Select
    Selection.AutoFill Destination:=CellRef
    
    Columns("AE:AE").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'/Irregularity
    
    Sheets(Sheets.Count).Activate
    Cells(1, 1).Select
    Sheets(Sheets.Count).Select
    Sheets(Sheets.Count).Copy After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Select
    Sheets(Sheets.Count).Name = "Quiet Breathing Times"
    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Select
    Sheets(Sheets.Count).Name = "Apneas"
    Sheets("Quiet Breathing Times").Select
        With ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
        End With
        ActiveWindow.FreezePanes = True
    Range("J1").Select
    Selection.EntireColumn.Insert
    ActiveCell.FormulaR1C1 = "[m]:ss.0"
    Range("J2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "[m]:ss.0"
    Range("K1").Select
    Selection.EntireColumn.Insert
    ActiveCell.FormulaR1C1 = "Include"
    Range("M1").Select
    Selection.EntireColumn.Insert
    ActiveCell.FormulaR1C1 = "60/f"
    Range("N1").Select
    Selection.EntireColumn.Insert
    ActiveCell.FormulaR1C1 = "Apnea"
    
    
    
   

    Cells(2, 10).Select
    ActiveCell.FormulaR1C1 = "=RC[-2]"
    Range("I2").Select
    Set CellRef = Range(Selection, Selection.End(xlDown)).Offset(0, 1)
    Cells(2, 10).Select
    Selection.AutoFill Destination:=CellRef
    
    Cells(2, 13).Select
    ActiveCell.FormulaR1C1 = "=60/RC[-1]"
    Range("L2").Select
    Set CellRef = Range(Selection, Selection.End(xlDown)).Offset(0, 1)
    Cells(2, 13).Select
    Selection.AutoFill Destination:=CellRef
    
'Deemed Unnecessary 12/22/09
    
'Dim row As Long
'row = 2
'
'Do While Cells(row, 13) <> ""
'    row = row + 1
'Loop
'
'row = row + 1
'
'Cells(row, 1).Value = "Times"
'Cells(row + 1, 1).Select
'Range(Selection, Selection.End(xlDown)).Select
'Selection.NumberFormat = "[m]:ss.0"
'Cells(row + 1, 2).Select
'Range(Selection, Selection.End(xlDown)).Select
'Selection.NumberFormat = "[m]:ss.0"


End Sub

Sub summary_data()

Dim qRow As Long
Dim aRow As Long

Worksheets("Quiet Breathing Times").Activate
Dim row As Long
row = 2

Do While Cells(row, 12) <> ""
    row = row + 1
Loop

row = row + 1

Cells(row, 11).Value = "Average"
Cells(row + 1, 11).Value = "SD"

Cells(row, 12).Value = "=average(R[" & 2 - row & "]C[0]:R[-2]C[0])"
Cells(row + 1, 12).Value = "=stdev(R[" & 1 - row & "]C[0]:R[-3]C[0])"

Cells(row, 13).Value = "=average(R[" & 2 - row & "]C[0]:R[-2]C[0])"
Cells(row + 1, 13).Value = "=stdev(R[" & 1 - row & "]C[0]:R[-3]C[0])"

Cells(row, 18).Value = "=average(R[" & 2 - row & "]C[0]:R[-2]C[0])"
Cells(row + 1, 18).Value = "=stdev(R[" & 1 - row & "]C[0]:R[-3]C[0])"

Cells(row, 19).Value = "=average(R[" & 2 - row & "]C[0]:R[-2]C[0])"
Cells(row + 1, 19).Value = "=stdev(R[" & 1 - row & "]C[0]:R[-3]C[0])"

Cells(row, 29).Value = "=average(R[" & 2 - row & "]C[0]:R[-2]C[0])"
Cells(row + 1, 29).Value = "=stdev(R[" & 1 - row & "]C[0]:R[-3]C[0])"

Cells(row, 35).Value = "=average(R[" & 2 - row & "]C[0]:R[-2]C[0])"
Cells(row + 1, 35).Value = "=stdev(R[" & 1 - row & "]C[0]:R[-3]C[0])"

qRow = row

Worksheets("Apneas").Activate

row = 2

Do While Cells(row, 12) <> ""
    row = row + 1
Loop

row = row + 1

Cells(row, 12).Value = "Total Time"
Cells(row + 1, 12).Value = "Minutes"
Cells(row + 2, 12).Value = "Apneas"
Cells(row + 3, 12).Value = "Apneas/min"
Cells(row + 4, 12).Value = "Ave. Apnea"
Cells(row + 5, 12).Value = "SD Apnea"

Cells(row, 13).Value = total
Cells(row, 13).Select
Selection.NumberFormat = "[m]:ss.0"
Cells(row + 1, 13).Value = "=minute(R[-1]C[0])+second(R[-1]C[0])/60"
Cells(row + 2, 13).Value = row - 3
Cells(row + 3, 13).Value = "=R[-1]C[0]/R[-2]C[0]"
Cells(row + 4, 13).Value = "=average(R[" & -2 - row & "]C[-4]:R[-6]C[-4])"
Cells(row + 5, 13).Value = "=stdev(R[" & -3 - row & "]C[-4]:R[-7]C[-4])"
Cells(row + 5, 13).NumberFormat = "s.000"

aRow = row

'Chart
    Call doChart
'/Chart

'Summary Sheet
    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Select
    Sheets(Sheets.Count).Name = "Summary"
    Sheets("Summary").Activate
    
    Cells(1, 1).Value = "Total Time"
    Cells(1, 2).Value = "Frequency"
    Cells(1, 3).Value = "Frequency SD"
    Cells(1, 4).Value = "Frequency CV"
    Cells(1, 5).Value = "Apneas/min."
    Cells(1, 6).Value = "Apnea Length"
    Cells(1, 7).Value = "Apnea Length SD"
    Cells(1, 8).Value = "Ti"
    Cells(1, 9).Value = "Te"
    Cells(1, 10).Value = "Penh"
    Cells(1, 11).Value = "Irr"
    
    Cells(2, 1).Value = Worksheets("Apneas").Cells(aRow, 13)
        Cells(2, 1).NumberFormat = "[m]:ss"
    Cells(2, 2).Value = Worksheets("Quiet Breathing Times").Cells(qRow, 12)
    Cells(2, 3).Value = Worksheets("Quiet Breathing Times").Cells(qRow + 1, 12)
    Cells(2, 4).Value = "=R[0]C[-1]/R[0]C[-2]"
    Cells(2, 5).Value = Worksheets("Apneas").Cells(aRow + 3, 13)
    Cells(2, 6).Value = Worksheets("Apneas").Cells(aRow + 4, 13)
        Cells(2, 6).NumberFormat = "s.000"
    Cells(2, 7).Value = Worksheets("Apneas").Cells(aRow + 5, 13)
        Cells(2, 7).NumberFormat = "s.000"
    Cells(2, 8).Value = Worksheets("Quiet Breathing Times").Cells(qRow, 18)
    Cells(2, 9).Value = Worksheets("Quiet Breathing Times").Cells(qRow, 19)
    Cells(2, 10).Value = Worksheets("Quiet Breathing Times").Cells(qRow, 29)
    Cells(2, 11).Value = Worksheets("Quiet Breathing Times").Cells(qRow, 35)
    
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A2:K2").Select
    Selection.Copy
'/Summary Sheet

Charts("Chart").Activate
End Sub
Sub doChart()
Worksheets("Quiet Breathing Times").Activate
    Dim row2 As Integer
    row2 = 2
    Do While Cells(row2, 1) <> ""
        row2 = row2 + 1
    Loop
row2 = row2 + 1
    
Cells(row2, 9).Value = "Min"
Cells(row2 + 1, 9).Value = "Max"

Cells(row2, 10).Value = "=min(R[" & 2 - row2 & "]C[0]:R[-2]C[0])"
Cells(row2 + 1, 10).Value = "=max(R[" & 1 - row2 & "]C[0]:R[-3]C[0])"

Dim low As Double
Dim high As Double

low = Cells(row2, 10).Value
high = Cells(row2 + 1, 10).Value
    
    
    Charts.Add
    ActiveChart.ChartType = xlXYScatter
    ActiveChart.SetSourceData Source:=Sheets("Quiet Breathing Times").Range("L1") _
        , PlotBy:=xlColumns
    ActiveChart.SeriesCollection(1).XValues = _
        "='Quiet Breathing Times'!R2C10:R" & row2 - 2 & "C10"
    ActiveChart.SeriesCollection(1).Values = _
        "='Quiet Breathing Times'!R2C12:R" & row2 - 2 & "C12"
    ActiveChart.SeriesCollection(1).Name = "='Quiet Breathing Times'!R1C12"
    ActiveChart.Location Where:=xlLocationAsNewSheet, Name:="Chart"
    With ActiveChart
        .HasTitle = False
        .Axes(xlCategory, xlPrimary).HasTitle = False
        .Axes(xlValue, xlPrimary).HasTitle = False
    End With
    With ActiveChart.Axes(xlCategory)
        .MinimumScale = low
        .MaximumScale = high
        .MinorUnitIsAuto = True
        .MajorUnitIsAuto = True
        .Crosses = xlAutomatic
        .ReversePlotOrder = False
        .ScaleType = xlLinear
        .DisplayUnit = xlNone
    End With
    ActiveChart.HasLegend = False
End Sub
