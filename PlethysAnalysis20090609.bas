Attribute VB_Name = "PlethysAnalysis"
Option Explicit
Dim total As Double
Sub step1()
Call prepare
End Sub

Sub step2()
Call check
Call delete
Call apneas
Call summary_data
End Sub


Sub check()
Worksheets("Quiet Breathing Times").Activate
total = 0
Dim CellRef As Range
Dim last As Boolean
last = False

'Old
'Modify
'Set CellRef = Range("A26605:B26610")
'/Modify

Dim i As Long

Dim tRow As Boolean
tRow = False
Dim row2 As Long
row2 = 2

Do While Not tRow
    If Cells(row2, 1) = "Times" Then
        tRow = True
    End If
    row2 = row2 + 1
Loop

Cells(row2, 1).Select
Range(Selection, Selection.End(xlDown)).Select
Set CellRef = Range(Selection, Selection.End(xlToRight))


Dim row3 As Long
row3 = row2

Do While Cells(row3, 1) <> ""
    Cells(row3, 3).Value = "=RC[-1]-RC[-2]"
    row3 = row3 + 1
Loop

Cells(row3, 2).Value = "Total"
Cells(row3, 3).Value = "=sum(R[" & row2 - row3 & "]C:R[-1]C)"
total = Cells(row3, 3).Value


Dim j As Long
j = 2

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

Cells(row - 1, 13).Value = "=2*average(R[" & 3 - row & "]C[0]:R[-1]C[0])"
Cells(row, 13).Value = "00:00:" & Cells(row - 1, 13).Value
Cells(row, 13).NumberFormat = "General"

Dim i As Long
Dim CellRef As Range

i = 3

Do While Cells(i, 8).Value <> ""
    If Cells(i, 9) > Cells(row, 13) Then
        Cells(i, 14).Value = "y"
    End If
    i = i + 1
Loop

Call sort

Cells(row - 2, 14).Select
Range(Selection, Selection.End(xlUp)).EntireRow.Select
Selection.Cut
Worksheets("Apneas").Activate
Worksheets("Apneas").Cells(2, 1).Select
ActiveSheet.Paste

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
    ActiveWorkbook.Worksheets("Quiet Breathing Times").sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Quiet Breathing Times").sort.SortFields.Add Key:=Range("I1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Quiet Breathing Times").sort
        .SetRange Range(Selection, Selection.End(xlDown)).EntireRow
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub sort_again()
    Cells(1, 12).Select
    Range(Selection, Selection.End(xlDown)).EntireRow.Select
    ActiveWorkbook.Worksheets("Quiet Breathing Times").sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Quiet Breathing Times").sort.SortFields.Add Key:=Range("L1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Quiet Breathing Times").sort
        .SetRange Range(Selection, Selection.End(xlDown)).EntireRow
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
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
    ActiveCell.FormulaR1C1 = "=RC[-1]-R[-1]C[-1]"
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
    
Dim row As Long
row = 2

Do While Cells(row, 13) <> ""
    row = row + 1
Loop

row = row + 1

Cells(row, 1).Value = "Times"
Cells(row + 1, 1).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.NumberFormat = "[m]:ss.0"
Cells(row + 1, 2).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.NumberFormat = "[m]:ss.0"


End Sub

Sub summary_data()
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


End Sub
