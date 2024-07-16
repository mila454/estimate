Attribute VB_Name = "LSR_LZU_12"

Option Explicit
Dim lastCell As Integer
Dim totalByPosition As New collection
Dim beginningOfSection As New collection
Dim totalForSection As New collection
Dim totalByEstimate As New collection
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
Dim forDelete As Range
Dim fot As Variant


'для лср лэу 12гр'
Sub initialDate()
'Создание исходных данных '

lastCell = seachLastCell()
Set seachRange = Range("A1:L" & lastCell)

Set forDelete = seachRange.Find("Составил*", LookIn:=xlValues, MatchCase:=True)
Rows(forDelete.row).Delete
Set forDelete = seachRange.Find("Проверил**", LookIn:=xlValues, MatchCase:=True)
Rows(forDelete.row).Delete

lastCell = seachLastCell()
Set seachRange = Range("A1:L" & lastCell)

'вставка Итого по разделу'
Set seachRange = Range("A1:L" & lastCell)
seachString = "Раздел *"
Set beginningOfSection = Seach(seachString, seachRange, "row")
Call quickSort.quickSort(beginningOfSection, 1, beginningOfSection.Count)
For i = beginningOfSection(1) + 1 To lastCell
    If Cells(i + 1, 1).Value <> 0 And Cells(i, 1).Value <> 0 And IsNumeric(Cells(i, 1)) And IsNumeric(Cells(i + 1, 1)) Then
        Rows(i + 1).Insert Shift:=xlDown
        With Cells(i + 1, 3)
            .Value = "Всего по позиции"
            .Font.Size = 11
            .Font.Bold = True
        End With
        lastCell = lastCell + 1
    End If
    
Next
If Cells(lastCell, 3).Value <> "Всего по позиции" Then
        With Cells(lastCell + 1, 3)
            .Value = "Всего по позиции"
            .Font.Size = 11
            .Font.Bold = True
        End With
End If
lastCell = seachLastCell()

Set seachRange = Range("A1:L" & lastCell)
seachString = "Раздел *"
Set beginningOfSection = Seach(seachString, seachRange, "row")
Call quickSort.quickSort(beginningOfSection, 1, beginningOfSection.Count)

beginningOfSection.Remove (1)
beginningOfSection.Add lastCell + 1
For i = 1 To beginningOfSection.Count
    If i > 1 Then
        Rows(beginningOfSection(i) + i - 1).Insert Shift:=xlDown
        With Cells(beginningOfSection(i) + i - 1, 2)
            .Value = "Итого по разделу"
            .Font.Size = 11
            .Font.Bold = True
        End With
        
    Else
        Rows(beginningOfSection(i)).Insert Shift:=xlDown
        With Cells(beginningOfSection(i), 2)
            .Value = "Итого по разделу"
            .Font.Size = 11
            .Font.Bold = True
        End With
    End If
    
    If i = beginningOfSection.Count Then
        With Cells(beginningOfSection(i) + i, 2)
            .Value = "ВСЕГО по смете"
            .Font.Size = 11
            .Font.Bold = True
        End With
    End If
    
Next

lastCell = seachLastCell()
Set seachRange = Range("A1:L" & lastCell)

seachString = "Итого по разделу*"
Set totalForSection = Seach(seachString, seachRange, "row")
seachString = "ВСЕГО по смете*"
Set totalByEstimate = Seach(seachString, seachRange, "row")
Call quickSort.quickSort(totalForSection, 1, totalForSection.Count)
Call quickSort.quickSort(totalByEstimate, 1, totalByEstimate.Count)

lastCell = seachLastCell()
Set seachRange = Range("A1:L" & lastCell)
seachString = "Всего по позиции*"

Set totalByPosition = Seach(seachString, seachRange, "row")
Call quickSort.quickSort(totalByPosition, 1, totalByPosition.Count)

For Each item In Range("B1:B" & totalByPosition(1))
    If item.Value Like "Шифр расценки и коды ресурсов  (обоснование коэффициента)" Then
        initialPosition = item.row + 6
    End If
Next
 totalByPosition.Add initialPosition
Call quickSort.quickSort(totalByPosition, 1, totalByPosition.Count)

ans = MsgBox("Есть необходимость заполнения графы Сметная стоимость в текущем уровне цен?", 4)
If ans = 6 Then
    Call filTotalForPosition
Else
    Call cumulativeList
End If

End Sub

Sub filTotalForPosition()
'заполнение итого по позиции в текущих ценах'


i = 0
For j = 2 To totalByPosition.Count

    
    For i = totalByPosition(j - 1) + 1 To totalByPosition(j)
        Call filCurrentPrices(i)
    Next
    Cells(totalByPosition(j), 11).Formula = "= SUM(K" & totalByPosition(j - 1) + 1 & ":K" & totalByPosition(j) - 1 & ")"
    Cells(totalByPosition(j), 13).Formula = "= K" & totalByPosition(j)
Next

Cells(totalByEstimate(1), 11).Formula = "= SUM(M" & totalByPosition(1) & ":M" & totalByEstimate(1) - 1 & ")"
Cells(totalByEstimate(1), 11).NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
Columns("K:K").EntireColumn.AutoFit

Cells(totalByEstimate(1), 11).Select

ans = MsgBox("Проверьте итоговую сумму сметы", 1)
If ans = 1 Then
    Cells(totalByEstimate(1), 12).Formula = "= M" & totalByEstimate(1)
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

seachString = "(наименование конструктивного решения*"
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

Call insertCol("Акт № 1", 15, initialPosition - 6, lastCell)
Cells(totalByEstimate(2), 15).Formula = "= SUM(O" & totalByPosition(1) + 1 & ":O" & totalByEstimate(2) - 1 & ")"
Call fillTail(15)

Call insertCol("Акт № 2", 17, initialPosition - 6, lastCell)
Cells(totalByEstimate(2), 17).Formula = "= SUM(Q" & totalByPosition(1) + 1 & ":Q" & totalByEstimate(2) - 1 & ")"
Call fillTail(17)

Call insertCol("ИТОГО по Актам", 19, initialPosition - 6, lastCell, "255 250 205")

For Each item In totalByPosition
    Cells(item, 19).Formula = "=O" & item & "+Q" & item
Next
Cells(totalByEstimate(2), 19).Formula = "= SUM(R" & totalByPosition(1) + 1 & ":R" & totalByEstimate(2) - 1 & ")"
Call fillTail(19)

Call insertCol("Остаток", 21, initialPosition - 6, lastCell, "240 230 140")
item = 0
For Each item In totalByPosition
    Cells(item, 21).Formula = "=L" & item & "-S" & item
Next
Cells(totalByEstimate(2), 21).Formula = "= SUM(U" & totalByPosition(1) + 1 & ":U" & totalByEstimate(2) - 1 & ")"
Call fillTail(21)

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


Select Case Cells(i, 3).Value
    Case "ЗТР"
        Cells(i, 11).ClearContents
    Case "ЭМ", "МР"
        Cells(i, 11).Formula = "=round(H" & i & "*J" & i & ",2)"
    Case "ЗП"
        Cells(i, 11).Formula = "=round(H" & i & "*J" & i & ",2)"
        fot = fot + Cells(i, 11).Value
    Case "в т.ч. ЗПМ"
        Cells(i, 11).ClearContents
        If Cells(i, 8).Value <> 0 Then
            fot = fot + Split(Split(Cells(i, 8), "(")(1), ")")(0) * Cells(i, 10)
        End If
    Case "НР от ФОТ"
        Cells(i, 11).Value = fot * Cells(i, 10).Value / 100
    Case "СП от ФОТ"
        Cells(i, 11).Value = fot * Cells(i, 10).Value / 100
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
        Cells(item.row, 12).Copy
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


