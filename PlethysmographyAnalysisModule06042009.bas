Attribute VB_Name = "Module1"
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
        If ActiveSheet.Cells(j, 9).Value > CellRef.Cells(i, 1) And ActiveSheet.Cells(j, 9).Value < CellRef.Cells(i, 2) Then
            ActiveSheet.Cells(j, 10).Value = "y"
        End If
    Next i
    j = j + 1
    If ActiveSheet.Cells(j, 9) = "" Then
        last = True
    End If
Loop


End Sub

Sub delete()
Dim i As Long
Dim CellRef As Range

i = 2

Do While Cells(i, 8).Value <> ""
    If Cells(i, 10) = "" Then
        Cells(i, 10).Select
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

Do While Cells(row, 12) <> ""
    row = row + 1
Loop

row = row + 1

Cells(row, 12).Value = "=2*average(R[" & 2 - row & "]C[0]:R[-2]C[0])"

Dim i As Long
Dim CellRef As Range

i = 2

Do While Cells(i, 8).Value <> ""
    If Cells(i, 12) > Cells(row, 12) Then
        Cells(i, 13).Value = "y"
    End If
    i = i + 1
Loop

Call sort

Cells(row - 2, 13).Select
Range(Selection, Selection.End(xlUp)).EntireRow.Select
Selection.Cut
Worksheets("Apnea").Activate
Worksheets("Apnea").Cells(2, 1).Select
ActiveSheet.Paste

Worksheets("Quiet Breathing").Activate
Cells(1, 1).EntireRow.Select
Selection.Copy
Worksheets("Apnea").Activate
Worksheets("Apnea").Cells(1, 1).Select
ActiveSheet.Paste

Worksheets("Quiet Breathing").Activate
Cells(2, 12).Select
Selection.End(xlDown).Select
Cells(ActiveCell.row + 1, ActiveCell.Column).Select
Range(Selection, Selection.End(xlDown)).Resize(Range(Selection, Selection.End(xlDown)).Rows.Count - 2).EntireRow.delete

Call sort_again


End Sub
Sub sort()
    Cells(2, 12).Select
    Range(Selection, Selection.End(xlDown)).EntireRow.Select
    ActiveWorkbook.Worksheets("Quiet Breathing").sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Quiet Breathing").sort.SortFields.Add Key:=Range("L2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Quiet Breathing").sort
        .SetRange Range(Selection, Selection.End(xlDown)).EntireRow
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub sort_again()
    Cells(2, 11).Select
    Range(Selection, Selection.End(xlDown)).EntireRow.Select
    ActiveWorkbook.Worksheets("Quiet Breathing").sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Quiet Breathing").sort.SortFields.Add Key:=Range("K2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Quiet Breathing").sort
        .SetRange Range(Selection, Selection.End(xlDown)).EntireRow
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub prepare()
'
' Macro10 Macro
'

'
    Sheets("WBP_Compensated1_Data").Activate
    Cells(1, 1).Select
    Sheets("WBP_Compensated1_Data").Select
    Sheets("WBP_Compensated1_Data").Copy After:=Sheets(Sheets.Count)
    Sheets("WBP_Compensated1_Data (2)").Select
    Sheets("WBP_Compensated1_Data (2)").Name = "Quiet Breathing"
    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Select
    Sheets(Sheets.Count).Name = "Apnea"
    Sheets("Quiet Breathing").Select
        With ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
        End With
        ActiveWindow.FreezePanes = True
    Range("I1").Select
    Selection.EntireColumn.Insert
    ActiveCell.FormulaR1C1 = "[m]:ss.0"
    Range("I2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "[m]:ss.0"
    Range("J1").Select
    Selection.EntireColumn.Insert
    ActiveCell.FormulaR1C1 = "Include"
    Range("L1").Select
    Selection.EntireColumn.Insert
    ActiveCell.FormulaR1C1 = "60/f"
    Range("M1").Select
    Selection.EntireColumn.Insert
    ActiveCell.FormulaR1C1 = "Apnea"
    
    
    
    Dim CellRef As Range

    Cells(2, 9).Select
    ActiveCell.FormulaR1C1 = "=RC[-1]"
    Range("H2").Select
    Set CellRef = Range(Selection, Selection.End(xlDown)).Offset(0, 1)
    Cells(2, 9).Select
    Selection.AutoFill Destination:=CellRef
    
    Cells(2, 12).Select
    ActiveCell.FormulaR1C1 = "=60/RC[-1]"
    Range("K2").Select
    Set CellRef = Range(Selection, Selection.End(xlDown)).Offset(0, 1)
    Cells(2, 12).Select
    Selection.AutoFill Destination:=CellRef
    
Dim row As Long
row = 2

Do While Cells(row, 12) <> ""
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
Worksheets("Quiet Breathing").Activate
Dim row As Long
row = 2

Do While Cells(row, 12) <> ""
    row = row + 1
Loop

row = row + 1

Cells(row, 10).Value = "Average"
Cells(row + 1, 10).Value = "SD"

Cells(row, 11).Value = "=average(R[" & 2 - row & "]C[0]:R[-2]C[0])"
Cells(row + 1, 11).Value = "=stdev(R[" & 1 - row & "]C[0]:R[-3]C[0])"

Cells(row, 12).Value = "=average(R[" & 2 - row & "]C[0]:R[-2]C[0])"
Cells(row + 1, 12).Value = "=stdev(R[" & 1 - row & "]C[0]:R[-3]C[0])"

Cells(row, 17).Value = "=average(R[" & 2 - row & "]C[0]:R[-2]C[0])"
Cells(row + 1, 17).Value = "=stdev(R[" & 1 - row & "]C[0]:R[-3]C[0])"

Cells(row, 18).Value = "=average(R[" & 2 - row & "]C[0]:R[-2]C[0])"
Cells(row + 1, 18).Value = "=stdev(R[" & 1 - row & "]C[0]:R[-3]C[0])"

Cells(row, 28).Value = "=average(R[" & 2 - row & "]C[0]:R[-2]C[0])"
Cells(row + 1, 28).Value = "=stdev(R[" & 1 - row & "]C[0]:R[-3]C[0])"

Worksheets("Apnea").Activate

row = 2

Do While Cells(row, 12) <> ""
    row = row + 1
Loop

row = row + 1

Cells(row, 11).Value = "Total Time"
Cells(row + 1, 11).Value = "Minutes"
Cells(row + 2, 11).Value = "Apneas"
Cells(row + 3, 11).Value = "Apneas/min"
Cells(row + 4, 11).Value = "Ave. Apnea"
Cells(row + 5, 11).Value = "SD Apnea"

Cells(row, 12).Value = total
Cells(row, 12).Select
Selection.NumberFormat = "[m]:ss.0"
Cells(row + 1, 12).Value = "=minute(R[-1]C[0])+second(R[-1]C[0])/60"
Cells(row + 2, 12).Value = row - 3
Cells(row + 3, 12).Value = "=R[-1]C[0]/R[-2]C[0]"
Cells(row + 4, 12).Value = "=average(R[" & -2 - row & "]C[0]:R[-6]C[0])"
Cells(row + 5, 12).Value = "=stdev(R[" & -3 - row & "]C[0]:R[-7]C[0])"


End Sub
