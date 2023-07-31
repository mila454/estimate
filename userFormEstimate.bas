Attribute VB_Name = "userFormEstimate"
Option Explicit

Dim lastRow As Integer
Dim seachRange As Range
Dim seachStr As String
Dim letterCol As String
Dim numberCol As Integer
Dim totalEstimate As New Collection 'номер строки итого по смете
Dim compile As New Collection 'номер строки составил
Dim tableCap As New Collection 'номер строки шапки таблицы


Sub userFormEstimate()

prepareEstimate.Show

End Sub

Sub nds(typeEstimate)

If typeEstimate = "ТСН" Then
    numberCol = 11
    letterCol = "K"
ElseIf typeEstimate = "СН" Then
    numberCol = 11
    letterCol = "K"
End If

Call activateSheet("Смета *")

Call clearTail

Set seachRange = Range(Cells(1, 1), Cells(lastRow, 10))
seachStr = "№ п/п"
Set tableCap = Seach(seachStr, seachRange)

Range("A" & totalEstimate(1) + 1 & ":A" & totalEstimate(1) + 3).EntireRow.Insert

If typeEstimate = "ТСН" Then
    Range("H" & totalEstimate(1) & ":I" & totalEstimate(1)).UnMerge
    Cells(totalEstimate(1), 9).formula = "=sum(O" & tableCap(1) + 2 & ":O" & totalEstimate(1) - 1 & ")"
    Cells(totalEstimate(1), 8).Clear
    Call ndsTotal(totalEstimate(1), 9, "I")
    Call setFormat(totalEstimate(1), 9, totalEstimate(1) + 2, 9)
End If

Range(Split(Range(letterCol & totalEstimate(1)).Offset(, -1).Address, "$")(1) & totalEstimate(1) & ":" & letterCol & totalEstimate(1)).UnMerge
Cells(totalEstimate(1), numberCol).formula = "=sum(P" & tableCap(1) + 2 & ":P" & totalEstimate(1) - 1 & ")"
Cells(totalEstimate(1), numberCol - 1).Clear

Call ndsTotal(totalEstimate(1), numberCol, letterCol)
Columns(letterCol & ":" & letterCol).EntireColumn.AutoFit

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
If totalEstimate(1) + 1 < compile(1) - 1 Then
    Range("A" & totalEstimate(1) + 1 & ":A" & compile(1) - 1).EntireRow.Hidden = False
    Range("A" & totalEstimate(1) + 1 & ":A" & compile(1) - 1).EntireRow.Delete
End If

End Sub

Sub ndsTotal(r, numberCol, letterCol)
'расчет и вывод НДС и итого с НДС

Cells(r + 1, 1).Value = "НДС 20%"
Cells(r + 1, numberCol).formula = "=round(" & letterCol & r & "*0.2,2)"
Cells(r + 2, 1) = "Итого с НДС 20%"
Cells(r + 2, numberCol).formula = "=" & letterCol & r & "+" & letterCol & (r + 1)
Range("A" & r + 1 & ":A" & r + 2).WrapText = False

End Sub

Sub setFormat(row1Format, col1Format, row2Format, col2Format)
'форматирование диапазона
With Range(Cells(row1Format, col1Format), Cells(row2Format, col2Format))
    .Font.Bold = True
    .NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
End With

End Sub
