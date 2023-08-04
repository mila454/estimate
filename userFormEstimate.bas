Attribute VB_Name = "userFormEstimate"
Option Explicit

Dim lastRow As Integer
Dim seachRange As Range
Dim seachStr As String
Dim letterCol As String
Dim numberCol As Integer
Dim currYear As Integer
Dim signer As String 'ФИО утверждающего
Dim position As String 'должность утверждающего
Dim typeEstimate As String 'тип сметы: ТСН или СН
Dim objectName As String 'наименование сметы
Dim totalEstimate As New Collection 'номер строки итого по смете
Dim approve As New Collection 'номер строки СОГЛАСОВАНО
Dim smetaName As New Collection 'наименование локальных смет

Sub userFormEstimate()

prepareEstimate.Show

End Sub

Sub nds()
Dim i As Variant
Dim item As Variant

Call activateSheet("Смета *")
lastRow = seachLastCell() + 1
Set seachRange = Range(Cells(1, 1), Cells(lastRow, 11))
Call determinationEstimateType
Call header
Call clearTail

If totalEstimate.Count > 1 Then
    For i = 1 To totalEstimate.Count - 1
        Range("A" & totalEstimate(i) + 1 & ":A" & totalEstimate(i) + 3).EntireRow.Insert
        Call ndsTotal(totalEstimate(i), numberCol, letterCol)
        totalEstimate.Add totalEstimate(i + 1) + i * 3, , , i
        totalEstimate.Remove (i + 2)
    Next
End If
Range("A" & totalEstimate(totalEstimate.Count) + 1 & ":A" & totalEstimate(totalEstimate.Count) + 3).EntireRow.Insert
Call ndsTotal(totalEstimate(totalEstimate.Count), numberCol, letterCol)

For Each item In totalEstimate
    Call cancelMerge(item)
Next

If typeEstimate = "ТСН" Then
    Call ndsTotal(totalEstimate(1), 9, "I")
    Call setFormat(totalEstimate(1), 9, totalEstimate(1) + 2, 9)
End If




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

Set foundCell = seachRange.Find(seachStr, LookIn:=xlValues) 'Cells(20, 6).Font.color = 16711680
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
'быстрая сортировка элеиентов коллекции
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
Dim i As Variant

Call activateSheet("Смета *")
Set seachRange = Range(Cells(1, 1), Cells(lastRow, 9))
seachStr = "Итого по*смете*"
Set totalEstimate = Seach(seachStr, seachRange)
Call quickSort(totalEstimate, 1, totalEstimate.Count)
If smetaName.Count > 1 Then
    Range("A" & totalEstimate(1) + 1 & ":A" & approve(1) + 5).EntireRow.Hidden = False
    Range("A" & totalEstimate(1) + 1 & ":A" & approve(1) + 5).EntireRow.Delete
    totalEstimate.Add totalEstimate(2) - 11, , , 1
    totalEstimate.Add totalEstimate(4) - 11, , , 2
    totalEstimate.Remove (totalEstimate.Count)
    totalEstimate.Remove (totalEstimate.Count)
    Range("A" & totalEstimate(2) + 1 & ":A" & totalEstimate(3) - 1).EntireRow.Hidden = False
    If (totalEstimate(3) - 1) - (totalEstimate(2) + 1) > 1 Then
        Range("A" & totalEstimate(2) + 1 & ":A" & totalEstimate(3) - 1).EntireRow.Delete
        totalEstimate.Add totalEstimate(3) - 2, , 3
        totalEstimate.Remove (totalEstimate.Count)
    End If

End If
Range("A" & totalEstimate(totalEstimate.Count) + 1 & ":A" & lastRow).EntireRow.Hidden = False
Range("A" & totalEstimate(totalEstimate.Count) + 1 & ":A" & lastRow).EntireRow.Delete

End Sub

Sub ndsTotal(r, numberCol, letterCol)
'расчет и вывод НДС и итого с НДС

Cells(r + 1, 1).Value = "НДС 20%"
Cells(r + 1, numberCol).formula = "=round(" & letterCol & r & "*0.2,2)"
Cells(r + 2, 1) = "Итого с НДС 20%"
Cells(r + 2, numberCol).formula = "=" & letterCol & r & "+" & letterCol & (r + 1)
Range("A" & r + 1 & ":A" & r + 2).WrapText = False
Call setFormat(r, numberCol, r + 2, numberCol)

End Sub

Sub setFormat(row1Format, col1Format, row2Format, col2Format)
'форматирование диапазона
With Range(Cells(row1Format, col1Format), Cells(row2Format, col2Format))
    .Font.Bold = True
    .NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
End With

End Sub

Sub determinationEstimateType()
'определение типа сметы: ТСН или СН
Dim currSheet As Variant

sheetName = "Смета*"
For Each currSheet In Worksheets
    If currSheet.Name Like sheetName Then
        If InStr(currSheet.Name, "ТСН") = 0 Then
            typeEstimate = "СН"
            numberCol = 10
            letterCol = "J"
        Else
            typeEstimate = "ТСН"
            numberCol = 11
            letterCol = "K"
        End If
    End If
Next

End Sub

Sub header()
'Формирование шапки, наименования сметы

Call activateSheet("Смета *")
seachStr = "СОГЛАСОВАНО"
Set approve = Seach(seachStr, seachRange)
If approve.Count <> 0 Then
    Rows("3:8").Delete
End If
Rows("3:8").Insert
signer = "Е.И. Новощинская"
position = "Заместитель директора ГКУ г.Москвы " & Chr(34) & "Дирекция Мосприроды" & Chr(34)
Cells(3, 3) = Chr(34) & "УТВЕРЖДАЮ" & Chr(34)
Cells(3, 3).HorizontalAlignment = xlLeft
Cells(5, 2) = "Заказчик:"
Cells(6, 2) = position
Cells(7, 2) = "_________________________ " & signer
currYear = Format(Date, "yyyy")
Cells(8, 2) = Chr(34) & "_____" & Chr(34) & "___________________ " & currYear & " г."
Cells(3, 3).Font.Bold = True
With Range("B6:E6, B7:D7, B8:D8")
    .MergeCells = True
    .WrapText = True
    .HorizontalAlignment = xlLeft
End With
Call heightAdjustment(Range("B6:D6"))
Call countEstimate

Range("A10") = objectName
Range("A15") = smetaName(2)
Call heightAdjustment(Range("A10:K10"))
Call heightAdjustment(Range("A15:K15"))

Cells(13, 1) = "ЛОКАЛЬНАЯ СМЕТА № 1"
With Range(Cells(3, 1), Cells(7, 2))
    .Font.Name = "Times New Roman"
    .Font.Size = 13
End With


End Sub

Sub heightAdjustment(mergedRange)
'автоподбор высоты объединенных ячеек
Dim myCell As Range, myLen As Integer, _
myWidth As Single, k As Single, n As Single

With mergedRange
    'Задаем объединенной ячейке перенос текста
    .WrapText = True
    'Задаем объединенной ячейке такую высоту строки, чтобы умещалась одна строка текста
    .RowHeight = Cells(mergedRange.Row, mergedRange.Column).Font.Size * 1.3
End With
myLen = Len(CStr(Cells(mergedRange.Row, mergedRange.Column)))
For Each myCell In mergedRange
    myWidth = myWidth + myCell.ColumnWidth
Next
n = 10
k = Cells(mergedRange.Row, mergedRange.Column).Font.Size / n
mergedRange.RowHeight = mergedRange.RowHeight * WorksheetFunction.RoundUp(myLen * k / myWidth, 0)

End Sub

Sub countEstimate()
'определение количества смет и их наменований
Dim temp As New Collection
Dim tempLastRow As Integer
Dim item As Variant

Call activateSheet("Source")
objectName = Cells(20, 7).Value
tempLastRow = Cells(Rows.Count, 6).End(xlUp).Row
seachStr = "Новая локальная смета*"
Set seachRange = Range(Cells(1, 6), Cells(tempLastRow, 6))
Set temp = Seach(seachStr, seachRange)

For Each item In temp
    If Cells(item, 6).Font.color = 16711680 Then
        smetaName.Add Cells(item, 7).Value
    End If
Next
Call activateSheet("Смета *")

End Sub

Sub cancelMerge(numberRow)
'отмена объединения ячеек и перенос формулы

If typeEstimate = "ТСН" Then
    Range("H" & numberRow & ":I" & numberRow).UnMerge
    Cells(numberRow, 9).formula = Cells(numberRow, 8).formula
    Cells(numberRow, 8).Clear
    Call setFormat(numberRow, 9, numberRow, 9)

End If

Range(Split(Range(letterCol & numberRow).Offset(, -1).Address, "$")(1) & numberRow & ":" & letterCol & numberRow).UnMerge
Cells(numberRow, numberCol).formula = Cells(numberRow, numberCol - 1).formula
Cells(numberRow, numberCol - 1).Clear
Call setFormat(numberRow, numberCol, numberRow, numberCol)

Columns(letterCol & ":" & letterCol).EntireColumn.AutoFit


End Sub
