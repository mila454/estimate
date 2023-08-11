Attribute VB_Name = "userFormEstimate"
Option Explicit

Dim lastRow As Integer
Dim seachRange As Range
Dim seachStr As String
Dim letterCol As String
Dim numberCol As Integer
Dim currYear As Integer
Dim answer As Variant 'ответ на вопрос об НДС в том числе
Dim signer As String 'ФИО утверждающего
Dim position As String 'должность утверждающего
Dim typeEstimate As String 'тип сметы: ТСН или СН
Dim totalEstimate As New Collection 'номер строки итого по смете
Dim rowForCoefficient As New Collection 'номер строки для вывода коэффициента
Dim namePosition As New Dictionary 'ключ:номер ЛОКАЛЬНАЯ СМЕТА № значение: наименование локальной сметы
Dim nameLocation As Variant 'номер ряда расположения наименования сметы
Dim smetaName As Variant 'наименование сметы
Dim numberEstimates As Integer 'количество смет
Dim coefficientName As String 'наименование коэффициета
Dim coefficient As Variant 'значение коэффициента

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
        Range("A" & totalEstimate(i) & ":H" & totalEstimate(i)).Value = "Итого по локальной смете №" & i & ": " & smetaName(i - 1)
        Call heightAdjustment(Range("A" & totalEstimate(i) & ":H" & totalEstimate(i)))
        If answer = 6 And Cells(totalEstimate(i), 1).Value Like "*Стоимость*посадочного*материала*" Then
            Call NDSIncluding(totalEstimate(i) + 1)
        Else
            Call ndsTotal(totalEstimate(i), numberCol, letterCol)
        End If
        totalEstimate.Add totalEstimate(i + 1) + i * 3, , , i
        totalEstimate.Remove (i + 2)
    Next
End If

Range("A" & totalEstimate(totalEstimate.Count) + 1 & ":A" & totalEstimate(totalEstimate.Count) + 3).EntireRow.Insert
Range("A" & totalEstimate(totalEstimate.Count) & ":H" & totalEstimate(totalEstimate.Count)).Value = "Итого по смете:" & smetaName(0)
If answer = 6 Then
    Call NDSIncluding(totalEstimate(totalEstimate.Count) + 1)
Else
    Call ndsTotal(totalEstimate(totalEstimate.Count), numberCol, letterCol)
End If
Call heightAdjustment(Range("A" & totalEstimate(totalEstimate.Count) & ":H" & totalEstimate(totalEstimate.Count)))
For Each item In totalEstimate
    Call cancelMerge(item)
Next

If typeEstimate = "ТСН" Then
    Call ndsTotal(totalEstimate(1), 9, "I")
    Call setFormat(totalEstimate(1), 9, totalEstimate(1) + 2, 9)
End If

Call finishOrGoToMainMenu

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

For i = 1 To totalEstimate.Count
    If i = totalEstimate.Count Then
        If numberEstimates = 1 Then
            Range("A" & totalEstimate(i) & ":A" & lastRow).EntireRow.Delete
            totalEstimate.Remove (totalEstimate.Count)
        Else
            Range("A" & totalEstimate(i) + 1 & ":A" & lastRow).EntireRow.Hidden = False
            Range("A" & totalEstimate(i) + 1 & ":A" & lastRow).EntireRow.Delete
        End If
    Else
        If numberEstimates > 1 Then
            If i < numberEstimates Then
                Range("A" & totalEstimate(i) + 1 & ":A" & nameLocation(i)).EntireRow.Hidden = False
                Range("A" & totalEstimate(i) + 1 & ":A" & nameLocation(i)).EntireRow.Delete
                totalEstimate.Add totalEstimate(i + 1) - ((nameLocation(i)) - (totalEstimate(i))), , , i
                totalEstimate.Add totalEstimate(i + 3) - ((nameLocation(i)) - (totalEstimate(i))), , , i + 1
                totalEstimate.Remove (totalEstimate.Count)
               
            Else
                Range("A" & totalEstimate(i) + 1 & ":A" & totalEstimate(i + 1)).EntireRow.Hidden = False
                Range("A" & totalEstimate(i) + 1 & ":A" & totalEstimate(i + 1) - 1).EntireRow.Delete
                totalEstimate.Add totalEstimate(i + 1) - ((totalEstimate(i + 1) - 1) - (totalEstimate(i))), , , i
                totalEstimate.Add totalEstimate(i + 3) - ((totalEstimate(i + 1) - 1) - (totalEstimate(i))), , , i + 1
                totalEstimate.Remove (totalEstimate.Count)
                totalEstimate.Remove (totalEstimate.Count)
                totalEstimate.Remove (totalEstimate.Count)
            End If
        Else
            Range("A" & totalEstimate(i) + 1 & ":A" & totalEstimate(i + 1)).EntireRow.Hidden = False
            Range("A" & totalEstimate(i) + 1 & ":A" & totalEstimate(i + 1) - 1).EntireRow.Delete
            totalEstimate.Add totalEstimate(i + 1) - ((totalEstimate(i + 1) - 1) - (totalEstimate(i))), , , i
            totalEstimate.Remove (totalEstimate.Count)
        End If
    End If
Next

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
Dim item As Variant
Dim i As Variant

Call activateSheet("Смета *")

For Each item In Range("A1:K" & lastRow)
    If item Like "*СОГЛАСОВАНО*" Then
        Rows("" & item.Row & ":" & item.Row + 5).Delete
        
    End If
Next
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

For i = 0 To numberEstimates - 1
    Cells(nameLocation(i) + 3, 1).Value = Cells(nameLocation(i), 1).Value & i + 1
    Cells(nameLocation(i) + 5, 1).Value = smetaName(i)
    Range(Cells(nameLocation(i), 1), Cells(nameLocation(i) + 1, numberCol + 1)).Clear
Next

Call heightAdjustment(Range("A10:K10"))
Call heightAdjustment(Range("A15:K15"))


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
Dim item As Variant
Dim quiestion As String

Set namePosition = New Dictionary

For Each item In Range("A1:K" & lastRow)
    If item Like "*ЛОКАЛЬНАЯ СМЕТА №*" Then
        namePosition.Add item.Row, Cells((item.Row + 3), item.Column).Value
        
    End If
Next

item = ""

nameLocation = namePosition.Keys
smetaName = namePosition.Items
numberEstimates = UBound(smetaName) + 1
For Each item In smetaName
    If item Like "*Стоимость*посадочного*материала" Then
        quiestion = "В смете: " & item & " НДС в том числе?"
        answer = MsgBox(quiestion, vbYesNo)
    End If
Next
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

Sub NDSIncluding(numberRow)
'НДС: в том числе

Cells(numberRow, 1).Value = "В том числе НДС 20%"
Cells(numberRow, numberCol).formula = "=round(" & letterCol & numberRow - 1 & "*20/120, 2)"
Range("A" & numberRow & ":A" & numberRow).WrapText = False
Call setFormat(numberRow, numberCol, numberRow, numberCol)

End Sub

Sub finishOrGoToMainMenu()
'закончить оформление сметы или перейти в основное меню
Dim passing As Integer

passing = MsgBox("Нажмите OK для продолжения оформления сметы или Cancel для завершения", vbOKCancel)

If passing = 1 Then
    Call userFormEstimate
Else
    Exit Sub
End If

End Sub

Sub coefBudgetFinancing()

coefficientName = "коэффициентом бюджетного финансирования"
Call completeAddCoef

End Sub

Sub addCoef(numberRow)
'добавление формулы для коэффициента
Rows((numberRow + 1) & ":" & (numberRow + 2)).Insert

Cells(numberRow + 1, 1).Value = "Итого с " & coefficientName & " K =" & coefficient
Cells(numberRow + 1, numberCol).formula = "=round(" & letterCol & numberRow & "*" & coefficient & ",2)"
Call NDSIncluding(numberRow + 2)

End Sub

Sub completeAddCoef()
'добавление коэффициента
Dim i As Variant


If typeEstimate = "" Then
    Call determinationEstimateType
End If

coefficient = InputBox("Перейдите на английский и введите значение коэффициента")

Call activateSheet("*Смета*")
If lastRow = 0 Then lastRow = seachLastCell() + 1

Set seachRange = Range(Cells(1, 1), Cells(lastRow, 11))
seachStr = "?????* НДС*"
Set rowForCoefficient = Seach(seachStr, seachRange)
Call quickSort(rowForCoefficient, 1, rowForCoefficient.Count)

For i = 0 To rowForCoefficient.Count - 1
    Call addCoef(rowForCoefficient(i + 1) + i * 2)
    
Next

Call activateSheet("ПНЦ")

End Sub

