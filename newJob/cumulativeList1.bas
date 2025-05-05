Attribute VB_Name = "cumulativeList1"

Option Explicit
Dim lastCell As Integer
Dim totalByPosition As New collection
Dim beginningOfSection As New collection
Dim totalForSection As New collection
Dim totalByEstimate As New collection
Dim coefMat As New collection
Dim coefMeh As New collection
Dim coefEquip As New collection
Dim coefTransp As New collection
Dim seachRange As Range
Dim seachString As String
Dim initialPosition As Integer
Dim i As Integer
Dim j As Integer
Dim ans As Integer
Dim numberColumn As Integer
Dim letterColumn As String
Dim ans2 As Integer
Dim ans3 As Integer
Dim item As Variant
Dim smetaName As New collection
Dim еquipment As New collection
Dim еquipmentSummary As New collection
Dim stringOfFormula As String
Dim shift As Integer

'универсальная: выбор объекта Россолимо или реконструкция или кс2 по Россолимо
'черновик
'в накопительной по актам не заполняет итоги по разделам, по  смете, хвост
'в накопительной скрытие столбцов твердо установлено
Sub selectMode()

ans = MsgBox("Объект - Общежитие Россолимо?", 4)
If ans = 6 Then
    ans = MsgBox("Объект - смета - Общежитие Россолимо?", 4)
    If ans = 6 Then
        numberColumn = 12
        letterColumn = "L"
        Call initialDate
    Else
        Call initialDate
    End If
Else
    Call cumulativeList4.initialDate
End If

End Sub

Sub initialDate()
'Создание исходных данных '

lastCell = seachLastCell()

Set seachRange = Range("A1:N" & lastCell)
seachString = "Раздел: *"
Set beginningOfSection = Estimate.Seach(seachString, seachRange)
seachString = "Итого по разделу *"
Set totalForSection = Seach(seachString, seachRange, "row")
seachString = "Всего по позиции"
Set totalByPosition = Seach(seachString, seachRange, "row")
seachString = "ВСЕГО по смете*"
Set totalByEstimate = Seach(seachString, seachRange, "row")
Call quickSort.quickSort(beginningOfSection, 1, beginningOfSection.Count)
For Each item In Range("B1:B" & totalByPosition(1))
    If item.Value Like "Обоснование" Then
        initialPosition = item.row + 6
    End If
Next
 totalByPosition.Add initialPosition

Call quickSort.quickSort(totalForSection, 1, totalForSection.Count)
Call quickSort.quickSort(totalByPosition, 1, totalByPosition.Count)
Call quickSort.quickSort(totalByEstimate, 1, totalByEstimate.Count)


Set seachRange = Range("A1:N" & lastCell)
Set coefMeh = Seach("эксплуатация машин и механизмов", seachRange, 3)
Set coefMat = Seach("материальные ресурсы", seachRange, 3)
Set coefTransp = Seach("перевозка", seachRange, 3)
Set coefEquip = Seach("Всего оборудование", seachRange, 3)


Call removeItemsFromCollection(coefMeh)
Call removeItemsFromCollection(coefMat)
Call removeItemsFromCollection(coefTransp)

ans = MsgBox("Есть необходимость заполнения графы Сметная стоимость в текущем уровне цен?", 4)
If ans = 6 Then
    Call filTotalForPosition
Else
    Call cumulativeList
End If

End Sub

Sub filTotalForPosition()
'заполнение итого по позиции в текущих ценах'
Dim beginning As Integer
Dim k As Integer

For j = 2 To totalByPosition.Count
    beginning = totalByPosition(j - 1)
    For k = 1 To beginningOfSection.Count
        If beginningOfSection(k) > totalByPosition(j - 1) And beginningOfSection(k) < totalByPosition(j) Then
            beginning = beginningOfSection(k)
            Exit For
        End If
    Next
    For i = beginning To totalByPosition(j)
        If Cells(i, 3).Value Like "Перевозка *" Or Cells(i, 3).Value Like "Погрузка *" Then
            Cells(i, numberColumn - 1).Value2 = coefMeh(1)
            Cells(i, numberColumn).Formula = "=round(K" & i & "*J" & i & ",2)"
        Else
            Call filCurrentPrices(i)
        End If
    Next
    Cells(totalByPosition(j), numberColumn).Formula = "= SUM(L" & beginning + 1 & ":L" & totalByPosition(j) - 1 & ")"
    If Cells(totalByPosition(j), numberColumn).MergeCells = True Then
        Call cancelMerge("K", totalByPosition(j), "L", totalByPosition(j), 1)
    End If
    
    Cells(totalByPosition(j), numberColumn + 1).Formula = "= L" & totalByPosition(j)
Next

'итого по разделам
For i = 1 To totalForSection.Count
    Cells(totalForSection(i), numberColumn).Formula = "= SUM(M" & beginningOfSection(i) & ":M" & totalForSection(i) - 1 & ")"
    Cells(totalForSection(i), numberColumn).NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
Next

'итого по смете
Cells(totalByEstimate(2), numberColumn + 1).Formula = "= SUM(M" & totalByPosition(1) & ":M" & totalByEstimate(2) - 1 & ")"
Cells(totalByEstimate(2), numberColumn + 1).NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
Columns("M:M").EntireColumn.AutoFit

Cells(totalByEstimate(2), 13).Select

ans = MsgBox("Проверьте итоговую сумму сметы", 1)
If ans = 1 Then
    'Cells(totalByEstimate(2), numberColumn).Formula = "= M" & totalByEstimate(2)
    ans3 = MsgBox("Нужен расчет оборудования?", 4)
    If ans3 = 6 Then
        Call calculationOfEquipment
    End If
    ans2 = MsgBox("Создать накопительную?", 4)
    If ans2 = 6 Then
        Call cumulativeList
    Else
        Exit Sub
    End If
Else
    Exit Sub
End If

End Sub
Sub cumulativeList()
Set seachRange = Range("A1:L" & lastCell)
seachString = "(наименование конструктивного решения)*"
Set smetaName = Seach(seachString, seachRange, "row")

'скрыть шапку и лишние столбцы
Range("A1:A" & smetaName(1) - 2).EntireRow.Hidden = True
Range("A" & smetaName(1) + 1 & ":A" & totalByPosition(1) - 7).EntireRow.Hidden = True
Columns("E:F").Hidden = True
Columns("H:I").Hidden = True
Columns("K:K").Hidden = True
Columns("M:M").Hidden = True

For j = 2 To totalByPosition.Count
    Range("A" & totalByPosition(j - 1) + 2 & ":A" & totalByPosition(j) - 1).EntireRow.Hidden = True
Next

Call insertCol("Акт № 1", numberColumn + 3, initialPosition - 6, lastCell)
Cells(totalByEstimate(2), numberColumn + 3).Formula = "= SUM(O" & totalByPosition(1) + 1 & ":O" & totalByEstimate(2) - 1 & ")"
Call fillTail(numberColumn + 3)

Call insertCol("Акт № 2", numberColumn + 5, initialPosition - 6, lastCell)
Cells(totalByEstimate(2), numberColumn + 5).Formula = "= SUM(Q" & totalByPosition(1) + 1 & ":Q" & totalByEstimate(2) - 1 & ")"
Call fillTail(numberColumn + 5)

Call insertCol("ИТОГО по Актам", numberColumn + 7, initialPosition - 6, lastCell, "255 250 205")

For Each item In totalByPosition
    Cells(item, numberColumn + 7).Formula = "=O" & item & "+Q" & item
Next
Cells(totalByEstimate(2), numberColumn + 7).Formula = "= SUM(R" & totalByPosition(1) + 1 & ":R" & totalByEstimate(2) - 1 & ")"
Call fillTail(numberColumn + 7)

Call insertCol("Остаток", numberColumn + 9, initialPosition - 6, lastCell, "240 230 140")
item = 0
For Each item In totalByPosition
    Cells(item, numberColumn + 9).Formula = "=L" & item & "-S" & item
Next
Cells(totalByEstimate(2), numberColumn + 9).Formula = "= SUM(U" & totalByPosition(1) + 1 & ":U" & totalByEstimate(2) - 1 & ")"
Call fillTail(numberColumn + 9)

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

Sub removeItemsFromCollection(coll)

For i = coll.Count To 1 Step -1
    If coll(i) = Empty Then
        coll.Remove (i)
    End If
Next


End Sub

Sub filCurrentPrices(i)
'заполнение сметная стоимость в текущем уровне цен, руб. ЭМ и М

'заполнение материалов в составе позиции
If Cells(i, 1).Value Like "*,*" And Not Cells(i, 3).Value Like "*ОБОРУДОВАНИЕ*" Then
    Cells(i, numberColumn - 1).Value2 = coefMat(1)
    Cells(i, numberColumn).Formula = "=round(K" & i & "*J" & i & ",2)"
End If

'заполнение материалов отдельной строкой

If Cells(i, 2).Value Like "#*" And IsEmpty(Cells(i + 1, 3)) And Not Cells(i, 3).Value Like "*ОБОРУДОВАНИЕ*" Then
    Cells(i, numberColumn - 1).Value2 = coefMat(1)
    Cells(i, numberColumn).Formula = "=round(K" & i & "*J" & i & ",2)"
End If

'заполнение оборудования

If Cells(i, 3).Value Like "*ОБОРУДОВАНИЕ*" And IsEmpty(Cells(i, 11)) Then
    Cells(i, numberColumn - 1).Value2 = coefEquip(1)
    Cells(i, numberColumn).Formula = "=round(K" & i & "*J" & i & ",2)"
End If

    Select Case Cells(i, 3).Value
        Case "ЭМ"
            Cells(i, numberColumn - 1).Value2 = coefMeh(1)
                Cells(i, numberColumn).Formula = "=round(K" & i & "*J" & i & ",2)"
        Case "М"
                Cells(i, numberColumn - 1).Value2 = coefMat(1)
                Cells(i, numberColumn).Formula = "=round(K" & i & "*J" & i & ",2)"
        Case "в т.ч. ОТм", "ФОТ"
                Cells(i, numberColumn).ClearContents
    End Select


End Sub

Sub insertCol(col_Name, col_ins, numberRow, lastCell, Optional fillCol = "255 255 255")
'вставка двух колонок с соответствующими названиями, форматирование
Dim range1 As Range
Dim fill_color() As String

Cells(, col_ins).EntireColumn.Insert
Cells(, col_ins).EntireColumn.Insert
Range((Cells(numberRow, (col_ins - 1))), Cells(numberRow, col_ins)).HorizontalAlignment = xlCenterAcrossSelection
Cells(numberRow, col_ins - 1).Value = col_Name
Cells(numberRow, col_ins - 1).WrapText = True
Cells(numberRow + 1, col_ins - 1).Value = "Кол-во"
Cells(numberRow + 1, col_ins - 1).HorizontalAlignment = xlCenter
Cells(numberRow + 1, col_ins).Value = "Стоимость, руб."
Cells(initialPosition - 7, col_ins - 1).VerticalAlignment = xlCenter
Set range1 = Range((Cells(numberRow, (col_ins - 1))), Cells(lastCell, col_ins))
If fillCol <> "" Then
    Call fillColor(range1, fillCol)
End If
With range1
    .Font.Size = 11
    .Borders.LineStyle = xlContinuous
    .ColumnWidth = 16
End With

End Sub

Sub fillTail(coll)
'заполнение хвоста

For Each item In Range("L" & totalByEstimate(2) + 1 & ":L" & lastCell)
    If item.HasFormula Then
        Cells(item.row, numberColumn).Copy
        Cells(item.row, coll).PasteSpecial xlFormulas
    End If
Next

With Range(Cells(totalByEstimate(2), coll), Cells(lastCell, coll))
    .Font.Bold = True
    .NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
End With

End Sub

Sub cancelMerge(col1, row1, col2, row2, transferStatus)
' Отмена объединения ячеек и перенос данных
Range(col1 & row1 & ":" & col2 & row2).UnMerge

If transferStatus = 1 Then
    Range(col1 & row1).Copy
    Range(col2 & row2).PasteSpecial (xlPasteValuesAndNumberFormats)
        Range(col1 & row1).Clear
Else
    Range(col1 & row1 & ":" & col2 & row2).Clear
End If

End Sub

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
    If Cells(еquipment(j), 1).Value Like "*,*" Then
        shift = 0
    Else
        If IsEmpty(Cells(еquipment(j) + 2, 12)) Then
            shift = 3
        Else
            shift = 2
        End If
    End If


    If j = 1 Then
        stringOfFormula = "L" & еquipment(j) + shift
    Else
        stringOfFormula = stringOfFormula & "+L" & еquipment(j) + shift
    End If
Next

If еquipmentSummary.Count = 1 Then
    Cells(еquipmentSummary(еquipmentSummary.Count), 12).Formula = "=" & stringOfFormula
Else
    Cells(еquipmentSummary(еquipmentSummary.Count - 1), 12).Formula = "=" & stringOfFormula
End If

End Sub











