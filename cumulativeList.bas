Attribute VB_Name = "cumulativeList"

Option Explicit
Dim currWB As Workbook
Dim lastCell As Integer
Dim totalByTypeOfWork As New Collection
Dim totalByPosition As New Collection
Dim beginningOfSection As New Collection
Dim totalForSection As New Collection
Dim totalByEstimate As New Collection
Dim coefFinanc As New Collection
Dim coefDecline As New Collection
Dim numberCoefFinance As Variant
Dim numberCoefDecline As Variant
Dim NDSInclud As New Collection


Sub cumulativeList()
'составление накопительной ведомости
Dim ws As Worksheet
Dim currRow As Variant
Dim seachRange As Range
Dim seachString As String
Dim KBRow As Integer
Dim tempSeach As New Collection
Dim tempRange As Range
Dim i As Long
Dim Shift As Integer
Set currWB = ActiveWorkbook
' скрыть все листы, кроме Смета...
For Each ws In currWB.Worksheets
    If InStr(ws.Name, "Смета") > 0 Then
        ws.Visible = xlSheetVisible
    Else
        ws.Visible = xlSheetHidden
    End If
Next ws

'скрыть шапку и лишние столбцы
Range("A1:A12").EntireRow.Hidden = True
Range("A15:A28").EntireRow.Hidden = True
With Range("A14:K14")
    .UnMerge
    .HorizontalAlignment = xlCenterAcrossSelection
End With
Columns("G:J").Hidden = True
Columns("K:AR").Hidden = False
lastCell = ContractEstimate.seachLastCell()
Range("L:M").Clear
Columns("L:L").Hidden = True

Set seachRange = Range("A1:K" & lastCell)
seachString = "Раздел: *"
Set beginningOfSection = Estimate.Seach(seachString, seachRange)
seachString = "Итого по разделу: *"
Set totalForSection = Estimate.Seach(seachString, seachRange)
seachString = "Всего по позиции:"
Set totalByPosition = Estimate.Seach(seachString, seachRange)
seachString = "Итого по локальной смете*"
Set totalByEstimate = Estimate.Seach(seachString, seachRange)
Call quickSort.quickSort(beginningOfSection, 1, beginningOfSection.Count)
Call quickSort.quickSort(totalForSection, 1, totalForSection.Count)
Call quickSort.quickSort(totalByPosition, 1, totalByPosition.Count)
Call quickSort.quickSort(totalByEstimate, 1, totalByEstimate.Count)

ActiveSheet.PageSetup.PrintArea = "$A$1:$AD$484"
'MsgBox rowWithTotal.Count
'Dim i As Long
'For i = 1 To rowWithTotal.Count
'    MsgBox rowWithTotal(i)
'    MsgBox lastRow(i)
'Next i

'удаление скрытых строк
For i = 1 To 5
    If Range("A" & totalByEstimate(1) + 1).EntireRow.Hidden Then
        With Range("A" & totalByEstimate(1) + 1)
            .EntireRow.Hidden = False
            .EntireRow.Delete
        End With
    End If
Next



Set seachRange = Range("A" & totalByEstimate(1) + 1 & ":K" & lastCell)
seachString = "Посадка деревьев*"
Set tempSeach = Estimate.Seach(seachString, seachRange)
totalByTypeOfWork.Add tempSeach(1)
seachString = "Восстановление отпада*"
Set tempSeach = Estimate.Seach(seachString, seachRange)
totalByTypeOfWork.Add tempSeach(1)
seachString = "Уходные работы*"
Set tempSeach = Estimate.Seach(seachString, seachRange)
totalByTypeOfWork.Add tempSeach(1)
totalByTypeOfWork.Add tempSeach(2)
totalByTypeOfWork.Add tempSeach(3)
Call quickSort.quickSort(totalByTypeOfWork, 1, totalByTypeOfWork.Count)
Set coefFinanc = Estimate.Seach("*коэффициент*финансиро*", seachRange)
Set coefDecline = Estimate.Seach("*коэффициент*снижен*", seachRange)
Set NDSInclud = Estimate.Seach("*в том числе НДС*", seachRange)
Call quickSort.quickSort(coefFinanc, 1, coefFinanc.Count)
Call quickSort.quickSort(coefDecline, 1, coefDecline.Count)
Call quickSort.quickSort(NDSInclud, 1, NDSInclud.Count)

If coefFinanc.Count = 0 Then
    numberCoefDecline = Replace(Left(Split(Cells(coefDecline(1), 1).Value, "=")(1), 13), ",", ".")
Else
    numberCoefDecline = Replace(Left(Split(Cells(coefDecline(1), 1).Value, "=")(1), 13), ",", ".")
    numberCoefFinance = Replace(Left(Split(Cells(coefFinanc(1), 1).Value, "=")(1), 13), ",", ".")
End If

Call insertCol("Исполнительная смета", 14, 32, lastCell, "255 250 205") 'LemonChiffon цвет
Call transferValues(totalByPosition, 14)
Call displayEstimate(14, "N")

Call insertCol("Отклонения", 16, 32, lastCell)
Call difference(totalByPosition, 10, "J", 14, "N", 16)
Call displayEstimate(16, "P")

Call insertCol("Акт № 1", 18, 32, lastCell)


Call displayEstimate(18, "R")

Call insertCol("Акт № 2", 20, 32, lastCell)


Call displayEstimate(20, "T")

Call insertCol("Акт № 3", 22, 32, lastCell)


Call displayEstimate(22, "V")

Call insertCol("Акт № 4", 24, 32, lastCell)

Call displayEstimate(24, "X")


Call insertCol("ИТОГО по Актам", 26, 32, lastCell, "255 215 0")

For Each currRow In totalByPosition
    Cells(currRow, 26).formula = "=R" & currRow & "+T" & currRow & "+V" & currRow & "+X" & currRow
    Call Estimate.setFormat(currRow, 26, currRow, 26)
Next

Call displayEstimate(26, "Z")



Call insertCol("Остаток по контрактной смете", 28, 32, lastCell)
currRow = 0
For Each currRow In totalByPosition
    Cells(currRow, 28).formula = "=J" & currRow & "-Z" & currRow
    Call Estimate.setFormat(currRow, 28, currRow, 28)
Next
Call displayEstimate(28, "AB")

Call insertCol("Остаток по исполнительной смете", 30, 32, lastCell, "255 250 205")
currRow = 0
For Each currRow In totalByPosition
    Cells(currRow, 30).formula = "=N" & currRow & "-Z" & currRow
    Call Estimate.setFormat(currRow, 30, currRow, 30)
Next
Call displayEstimate(30, "AD")

currRow = 0
For Each currRow In totalForSection
    Set tempRange = Range("A" & currRow & ":AD" & currRow)
    Call fillColor(tempRange, "143 188 139")
Next
currRow = 0
For Each currRow In totalByEstimate
    Set tempRange = Range("A" & currRow & ":AD" & currRow)
    Call fillColor(tempRange, "255 218 185")
Next

Range("33:33,A:K").Select
Range("A13").Activate
ActiveWindow.FreezePanes = True



'скрыть подстроки в расценках
currRow = 0

Range("A" & beginningOfSection(1) + 2 & ":A" & totalByPosition(1) - 1).EntireRow.Hidden = True
For i = 2 To totalByPosition.Count
    Shift = 3
    If Cells((totalByPosition(i - 1) + Shift), 1).Value Like "Итого по разделу*" Then
        Shift = 9
    End If
    If (totalByPosition(i - 1) + Shift) <> (totalByPosition(i)) Then
        Range("A" & totalByPosition(i - 1) + Shift & ":A" & totalByPosition(i) - 1).EntireRow.Hidden = True
    End If
Next





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
Cells(32, col_ins - 1).VerticalAlignment = xlCenter
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

Sub fillColor(fillRange, color)
'заполнение области цветом
Dim fill_color() As String

fill_color = Split(color)
fillRange.Interior.color = RGB(fill_color(0), fill_color(1), fill_color(2))

End Sub

Sub transferValues(rangeValue, col_copy)
'перенос данных
Dim currRow As Variant

For Each currRow In totalByPosition
    Cells(currRow, col_copy) = Cells(currRow, 10).Value
    Call Estimate.setFormat(currRow, col_copy, currRow, col_copy)
Next

End Sub
    
Sub outputTotalForSection(numberCol, letterCol, Optional ByVal numberRow As Integer = 0)
'вывод формулы итого по разделу
Dim i As Long

i = 1
If numberRow = 0 Then
    Do While i <= totalForSection.Count
        Cells(totalForSection(i), numberCol).formula = "=Sum(" & letterCol & beginningOfSection(i) & ":" & letterCol & (totalForSection(i) - 1) & ")"
        Call Estimate.setFormat(totalForSection(i), numberCol, totalForSection(i), numberCol)
        i = i + 1
    Loop
End If

End Sub

Sub difference(rangeDiff, col1, letterCol1, col2, letterCol2, targetCol)
'вывод формулы разности
Dim currRow As Variant

For Each currRow In rangeDiff
    Cells(currRow, targetCol).formula = "=" & letterCol1 & currRow & "-" & letterCol2 & currRow
    Call Estimate.setFormat(currRow, col2, currRow, col2)
Next

End Sub

Sub totalWithNDS(numberRow, numberCol, letterCol)
'вывод формулы итого с НДС
Cells(numberRow + 1, numberCol).formula = "=round(" & letterCol & (numberRow) & "*0.2,2)"
Cells(numberRow + 2, numberCol).formula = "=" & letterCol & numberRow & "+" & letterCol & numberRow + 1
Call Estimate.setFormat(numberRow + 1, numberCol, numberRow + 2, numberCol)
End Sub

Sub displayEstimate(numberCol, letterCol)
'вывод хвоста сметы
Dim rowCoef As Variant
Dim rowType As Variant
Dim rowNDS As Variant
Dim formul As String
Dim i As Long

formul = "="
Call outputTotalForSection(numberCol, letterCol)
Cells(totalByEstimate(1), numberCol).Value = Estimate.formulaTotal(totalForSection, letterCol)
Call totalWithNDS(totalByEstimate(1), numberCol, letterCol)

For i = 1 To totalByTypeOfWork.Count
    For Each rowType In totalForSection
        If Cells(rowType, 1).Value Like "*Посадка*" Then
            formul = formul & "+" & letterCol & rowType
        End If
    Next
    Cells(totalByTypeOfWork(i), numberCol).Value = formul
    formul = "="
    i = i + 1
    For Each rowType In totalForSection
        If Cells(rowType, 1).Value Like "*Уходные работы*2-й этап*" Then
            formul = formul & "+" & letterCol & rowType
        End If
    Next
    Cells(totalByTypeOfWork(i), numberCol).Value = formul
    formul = "="
    i = i + 1
    For Each rowType In totalForSection
        If Cells(rowType, 1).Value Like "*Восстановление*" Then
            formul = formul & "+" & letterCol & rowType
        End If
    Next
    Cells(totalByTypeOfWork(i), numberCol).Value = formul
    formul = "="
    i = i + 1
    For Each rowType In totalForSection
        If Cells(rowType, 1).Value Like "*Уходные работы*3-й этап*" Then
            formul = formul & "+" & letterCol & rowType
        End If
    Next
    Cells(totalByTypeOfWork(i), numberCol).Value = formul
    formul = "="
    i = i + 1
    For Each rowType In totalForSection
        If Cells(rowType, 1).Value Like "*Уходные работы*4-й этап*" Then
            formul = formul & "+" & letterCol & rowType
        End If
    Next
    Cells(totalByTypeOfWork(i), numberCol).Value = formul
    formul = "="
    i = i + 1
Next

For i = 1 To totalByTypeOfWork.Count
    Call totalWithNDS(totalByTypeOfWork(i), numberCol, letterCol)
Next

For Each rowCoef In coefFinanc
    Cells(rowCoef, numberCol).formula = "=round(" & letterCol & rowCoef - 1 & "*" & numberCoefFinance & ",2)"
Next

rowCoef = 0

For Each rowCoef In coefDecline
    Cells(rowCoef, numberCol).formula = "=round(" & letterCol & rowCoef - 1 & "*" & numberCoefDecline & ",2)"
Next

rowCoef = 0


For Each rowCoef In NDSInclud
    Call estimateSN.NDSIncluding(rowCoef, numberCol, letterCol)
Next

Call Estimate.setFormat(totalByEstimate(1), numberCol, lastCell, numberCol)

End Sub


