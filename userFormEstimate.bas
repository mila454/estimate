Attribute VB_Name = "userFormEstimate"
Option Explicit

Dim lastRow As Integer
Dim seachRange As Range
Dim seachStr As String
Dim totalEstimate As New Collection
Dim compile As New Collection

Sub userFormEstimate()

prepareEstimate.Show

End Sub

Sub nds()

Call activateSheet("Смета *")
Call clearTail



End Sub

Function seachLastCell()
' поиск последней непустой ячейки в столбцах с 1-го по 11-й
    Dim c(11) As Integer
    Dim i As Variant
    
    For i = 1 To 11
        c(i) = Cells(Rows.Count, i).End(xlUp).Row
    Next
    seachLastCell = WorksheetFunction.Max(c)
End Function

Sub activateSheet(sheetName)
' активирование листа
Dim currSheet As Variant

For Each currSheet In Worksheets
    If currSheet.Name Like sheetName Then
        currSheet.Activate
        sheetName = ActiveSheet.Name
    End If
Next

End Sub

Function Seach(seachStr, seachRange) As Collection
'поиск по строке и сохранение номера ряда в коллекцию
Dim foundCell As Range
Dim firstFoundCell As Range

Set Seach = New Collection

Set foundCell = seachRange.Find(seachStr, LookIn:=xlValues)
Set firstFoundCell = foundCell

If firstFoundCell Is Nothing Then
    MsgBox (seachStr & " не найдено")
    Exit Function
End If

Do
    Set foundCell = seachRange.FindNext(After:=foundCell)
    Seach.Add foundCell.Row
Loop While foundCell.Address <> firstFoundCell.Address

End Function

Sub quickSort(coll As Collection, first As Long, last As Long)

Dim centreVal As Variant, temp As Variant
Dim low As Long
Dim high As Long

If last = 0 Then Exit Sub

low = first
high = last
centreVal = coll((first + last) \ 2)

Do While low <= high
    Do While coll(low) < centreVal And low < last
        low = low + 1
    Loop
    Do While centreVal < coll(high) And high > first
        high = high - 1
    Loop
    If low <= high Then
    ' Поменять значения
        temp = coll(low)
        coll.Add coll(high), After:=low
        coll.Remove low
        coll.Add temp, Before:=high
        coll.Remove high + 1
        ' Перейти к следующим позициям
        low = low + 1
        high = high - 1
    End If
    Loop
    If first < high Then quickSort coll, first, high
    If low < last Then quickSort coll, low, last
End Sub

Sub clearTail()
'очистка, удаление объединения между Итого по локальной смете... и Составил-Проверил
lastRow = seachLastCell() + 1
Set seachRange = Range(Cells(1, 1), Cells(lastRow, 9))
seachStr = "Итого по*смете*"
Set totalEstimate = Seach(seachStr, seachRange)
Call quickSort(totalEstimate, 1, totalEstimate.Count)
seachStr = "Составил"
Set compile = Seach(seachStr, seachRange)
Range("A" & totalEstimate(1) + 1 & ":A" & compile(1) - 1).EntireRow.Hidden = False
Range("A" & totalEstimate(1) + 1 & ":A" & compile(1) - 1).EntireRow.Delete
End Sub
