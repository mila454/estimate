Attribute VB_Name = "cumulativeList5"

Option Explicit
Dim lastCell As Integer
Dim totalByPosition As New collection
Dim beginningOfSection As New collection
Dim beginningOfPosition As New collection
Dim beginningOfPosition2 As New collection
Dim totalForSection As New collection
Dim totalsForSection As New collection
Dim totalByEstimate As New collection
Dim totalsByEstimate As New collection
Dim coefMat As New collection
Dim coefMeh As New collection
Dim coefTransp As New collection
Dim seachRange As Range
Dim seachString As String
Dim initialPosition As Integer
Dim i As Integer
Dim j As Integer
Dim ans As Integer
Dim ans2 As Integer
Dim item As Variant
Dim smetaName As New collection
'создание накопительной с заполнением текущих цен
'рабочая версия
'объект Реконструкция здания. Отделение медицинской реабилитации

Sub cumulativeList5()

lastCell = seachLastCell()

Set seachRange = Range("A1:N" & lastCell)
seachString = "Раздел *"
Set beginningOfSection = Seach(seachString, seachRange, "row")
seachString = "Итого по разделу *"
Set totalForSection = Seach(seachString, seachRange, "row")
seachString = "Всего по позиции"
Set totalByPosition = Seach(seachString, seachRange, "row")
seachString = "ВСЕГО по смете*"
Set totalByEstimate = Seach(seachString, seachRange, "row")
seachString = "Итоги по разделу*"
Set totalsForSection = Seach(seachString, seachRange, "row")
seachString = "Итоги по смете*"
Set totalsByEstimate = Seach(seachString, seachRange, "row")
Set seachRange = Range("A1:N" & lastCell)
Set coefMeh = Seach("эксплуатация машин и механизмов", seachRange, 2)
Set coefMat = Seach("материалы", seachRange, 2)
Call quickSort.quickSort(totalsForSection, 1, totalsForSection.Count)
Call quickSort.quickSort(totalsByEstimate, 1, totalsByEstimate.Count)
Call quickSort.quickSort(beginningOfSection, 1, beginningOfSection.Count)
Call quickSort.quickSort(totalForSection, 1, totalForSection.Count)
Call quickSort.quickSort(totalByEstimate, 1, totalByEstimate.Count)
For Each item In Range("B1:B" & totalByPosition(1))
    If item.Value Like "Обоснование" Then
        initialPosition = item.row + 5
    End If
Next
Set seachRange = Range("B" & initialPosition & ":B" & lastCell)
seachString = "Ф*"
Set beginningOfPosition = Seach(seachString, seachRange, "row")
seachString = "ТЦ*"
Set beginningOfPosition2 = Seach(seachString, seachRange, "row")
For i = 1 To beginningOfPosition2.Count
    beginningOfPosition.Add beginningOfPosition2(i)
Next
Set beginningOfPosition2 = New collection

Call quickSort.quickSort(beginningOfPosition, 1, beginningOfPosition.Count)
Call quickSort.quickSort(totalByPosition, 1, totalByPosition.Count)

Set seachRange = Range("A1:N" & lastCell)
seachString = "(наименование работ и затрат*"
Set smetaName = Seach(seachString, seachRange, "row")

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

For j = 1 To totalByPosition.Count
    beginning = beginningOfPosition(j)
    For k = 1 To beginningOfSection.Count
        If beginningOfSection(k) > beginningOfPosition(j) And beginningOfSection(k) < totalByPosition(j) Then
            beginning = beginningOfSection(k)
            Exit For
        End If
    Next
    For i = beginning To totalByPosition(j)
        Call filCurrentPrices(i)
    Next
    If (totalByPosition(j) - beginning) > 1 Then
        Cells(totalByPosition(j), 14).Formula = "= SUM(N" & beginning + 1 & ":N" & totalByPosition(j) - 1 & ")"
    Else
        Cells(totalByPosition(j), 14).Formula = "=N" & totalByPosition(j) - 1
    End If
    
    Cells(totalByPosition(j), 15).Formula = "= N" & totalByPosition(j)
Next

For i = 1 To totalForSection.Count
    Cells(totalForSection(i), 14).Formula = "= SUM(O" & beginningOfSection(i) & ":O" & totalForSection(i) - 1 & ")"
    Cells(totalForSection(i), 14).NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
Next

Cells(totalByEstimate(1), 15).Formula = "= SUM(O" & totalByPosition(1) & ":O" & totalByEstimate(1) - 1 & ")"
Cells(totalByEstimate(1), 15).NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
Columns("O:O").EntireColumn.AutoFit

Cells(totalByEstimate(1), 14).Select

ans = MsgBox("Проверьте итоговую сумму сметы", 1)
If ans = 1 Then
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

seachString = " (наименование работ и затрат)*"
Set smetaName = Seach(seachString, seachRange, "row")

'скрыть шапку и лишние столбцы
Range("A1:A" & smetaName(1) - 2).EntireRow.Hidden = True
Range("A" & smetaName(1) + 1 & ":A" & initialPosition - 5).EntireRow.Hidden = True
Columns("G:H").Hidden = True
Columns("J:M").Hidden = True
For j = 1 To totalByPosition.Count
    If (totalByPosition(j) - beginningOfPosition(j)) > 1 Then
        Range("A" & beginningOfPosition(j) + 1 & ":A" & totalByPosition(j) - 1).EntireRow.Hidden = True
    End If
Next

If totalsForSection.Count > 0 Then
    For j = 1 To totalForSection.Count
        Range("A" & totalsForSection(j) & ":A" & totalForSection(j) - 1).EntireRow.Hidden = True
    Next
End If

For j = 1 To totalByEstimate.Count
    Range("A" & totalsByEstimate(j) & ":A" & totalByEstimate(j) - 1).EntireRow.Hidden = True
Next

Call insertCol("Акт № 1", 15, smetaName(1), lastCell)
Cells(totalByEstimate(1), 16).Formula = "= SUM(P" & totalByPosition(1) & ":P" & totalByEstimate(1) - 1 & ")"
Call fillTail(16)

Call insertCol("Акт № 2", 17, smetaName(1), lastCell)
Cells(totalByEstimate(1), 18).Formula = "= SUM(R" & totalByPosition(1) & ":R" & totalByEstimate(1) - 1 & ")"
Call fillTail(18)

Call insertCol("ИТОГО по Актам", 19, smetaName(1), lastCell, "255 250 205")

For Each item In totalByPosition
    Cells(item, 20).Formula = "=P" & item & "+R" & item
    Cells(item, 19).Formula = "=O" & item & "+Q" & item
Next
Cells(totalByEstimate(1), 20).Formula = "= SUM(T" & totalByPosition(1) & ":T" & totalByEstimate(1) & ")"
Call fillTail(20)

Call insertCol("Остаток", 21, smetaName(1), lastCell, "240 230 140")
item = 0
For Each item In totalByPosition
    Cells(item, 22).Formula = "=N" & item & "-T" & item
Next
Cells(totalByEstimate(1), 22).Formula = "= SUM(V" & totalByPosition(1) & ":V" & totalByEstimate(1) - 1 & ")"
Call fillTail(22)

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
'значение берется из колонки, сдвинутой на token от найденной
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

If Cells(i + 1, 2).Value Like "ФССЦ-*" Then
    Cells(i + 1, 13).Value2 = coefMat(1)
    Cells(i + 1, 14).Formula = "=round(L" & i + 1 & "*M" & i + 1 & ",2)"
End If
    
    Select Case Cells(i, 3).Value
        Case "ЭМ"
            Cells(i, 13).Value2 = coefMeh(1)
                Cells(i, 14).Formula = "=round(L" & i & "*M" & i & ",2)"
        Case "М"
                Cells(i, 13).Value2 = coefMat(1)
                Cells(i, 14).Formula = "=round(L" & i & "*M" & i & ",2)"
        Case "в т.ч. ОТм", "ФОТ"
                Cells(i, 14).ClearContents
    End Select


End Sub

Sub insertCol(col_Name, col_ins, numberRow, lastCell, Optional fillCol = "255 255 255")
'вставка двух колонок с соответствующими названиями, форматирование
Dim range1 As Range
Dim fill_color() As String

Cells(, col_ins).EntireColumn.Insert
Cells(, col_ins).EntireColumn.Insert
Range((Cells(numberRow - 1, (col_ins))), Cells(numberRow - 1, col_ins + 1)).HorizontalAlignment = xlCenterAcrossSelection
Cells(numberRow - 1, col_ins).Value = col_Name
Cells(numberRow, col_ins).Value = "Кол-во"
Cells(numberRow, col_ins).HorizontalAlignment = xlCenter

With Cells(numberRow, col_ins + 1)
    .Value = "Стоимость, руб."
    .WrapText = True
    .EntireRow.AutoFit
End With


Set range1 = Range((Cells(numberRow, (col_ins))), Cells(lastCell, col_ins + 1))
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

For Each item In Range("N" & totalByEstimate(1) + 1 & ":N" & lastCell)
    If item.HasFormula Then
        Cells(item.row, 14).Copy
        Cells(item.row, coll).PasteSpecial xlFormulas
    End If
Next

With Range(Cells(totalByEstimate(1), coll), Cells(lastCell, coll))
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



