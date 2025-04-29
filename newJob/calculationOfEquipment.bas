Attribute VB_Name = "calculationOfEquipment"
Option Explicit
Dim lastCell As Integer
Dim seachRange As Range
Dim seachString As String
Dim еquipment As New collection
Dim i As Integer
Dim еquipmentSummary As New collection
Dim j As Integer
Dim stringOfFormula As String


Sub calculationOfEquipment()
'расчет оборудования Россолимо

j = 1

lastCell = seachLastCell()

Set seachRange = Range("A1:C" & lastCell)
seachString = "ОБОРУДОВАНИЕ:*"
Set еquipment = Estimate.Seach(seachString, seachRange)
seachString = "Оборудование"
Set еquipmentSummary = Estimate.Seach(seachString, seachRange)
Call quickSort.quickSort(еquipment, 1, еquipment.Count)

For j = 1 To еquipment.Count
    If j = 1 Then
        If IsEmpty(Cells(еquipment(j) + 2, 12)) Then
            stringOfFormula = "L" & еquipment(j) + 3
        Else
            stringOfFormula = "L" & еquipment(j) + 2
        End If
    Else
        If IsEmpty(Cells(еquipment(j) + 2, 12)) Then
            stringOfFormula = stringOfFormula & "+L" & еquipment(j) + 3
        Else
            stringOfFormula = stringOfFormula & "+L" & еquipment(j) + 2
        End If
    End If
Next

Cells(еquipmentSummary(еquipmentSummary.Count - 1), 12).Formula = "=" & stringOfFormula



End Sub

Function seachLastCell()
' поиск последней непустой ячейки в столбцах с 1-го по 12-й
    Dim c(12) As Integer
    For i = 1 To 12
        c(i) = Cells(Rows.Count, i).End(xlUp).row
    Next
    seachLastCell = WorksheetFunction.Max(c)
End Function
Function Seach(seachStr, seachRange, token) As collection
'поиск по строке и сохранение номера ряда в коллекцию
Dim foundCell As Range
Dim firstFoundCell As Range

Set Seach = New collection

Set foundCell = seachRange.Find(seachStr, LookIn:=xlValues, MatchCase:=True)
Set firstFoundCell = foundCell

If firstFoundCell Is Nothing Then
    MsgBox (seachStr & " не найдено")
    Exit Function
End If

Do
    Set foundCell = seachRange.FindNext(After:=foundCell)
    If token = "row" Then
        Seach.Add foundCell.row
    Else
        Seach.Add foundCell.Offset(0, token).Value
    End If
    
Loop While foundCell.Address <> firstFoundCell.Address

End Function

