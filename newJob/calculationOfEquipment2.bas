Attribute VB_Name = "calculationOfEquipment2"
Option Explicit
Dim lastCell As Integer
Dim seachRange As Range
Dim seachString As String
Dim еquipment As New collection
Dim i As Integer
Dim еquipmentSummary As New collection
Dim j As Integer
Dim stringOfFormula As String


Sub calculationOfEquipment2()
'расчет оборудования Реконструкция
'рабочая

j = 1
i = 1


lastCell = seachLastCell()
Set seachRange = Range("A1:N" & lastCell)

For i = 1 To lastCell
    If Cells(i, 1).Value Like "*О" Then
       еquipment.Add i
    End If
Next

seachString = "Оборудование"
Set еquipmentSummary = Seach(seachString, seachRange)
Call quickSort.quickSort(еquipment, 1, еquipment.Count)

For j = 1 To еquipment.Count
    If j = 1 Then
        stringOfFormula = "N" & еquipment(j)
    Else
        stringOfFormula = stringOfFormula & "+N" & еquipment(j)
    End If
Next

Cells(еquipmentSummary(еquipmentSummary.Count - 1), 14).Formula = "=" & stringOfFormula

Set еquipment = New collection
Set еquipmentSummary = New collection


End Sub

Function seachLastCell()
' поиск последней непустой ячейки в столбцах с 1-го по 12-й
    Dim c(12) As Integer
    For i = 1 To 12
        c(i) = Cells(Rows.Count, i).End(xlUp).row
    Next
    seachLastCell = WorksheetFunction.Max(c)
End Function
Function Seach(seachStr, seachRange) As collection
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
    Seach.Add foundCell.row
Loop While foundCell.Address <> firstFoundCell.Address

End Function

