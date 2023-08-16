Attribute VB_Name = "breakdownByMonths"
Option Explicit
Dim lastCell As Integer
Dim totalForSection As New Collection
Dim beginningOfSection As New Collection



Sub breakdownByMonths()

'НДС 20% должен быть сразу после Итого по локальной смете
'разбивка по месяцам сметы ТСН обслуживание

Dim currWB As Workbook
Dim ws As Worksheet
Dim currentRow As Variant
Dim rowPosition As New Collection
Dim seachString As String
Dim totalByEstimate As New Collection
Dim coefFinanc As New Collection
Dim coefDecline As New Collection
Dim seachRange As Range
Dim numberCoefFinance As Variant
Dim numberCoefDecline As Variant
Dim rowNDS As New Collection
Dim firstRow As Integer
Dim lastRow As Integer
Dim periodCoef As Integer
Dim i As Integer
Dim j As Integer
Dim currentRow2 As Variant
Dim letterCol(13 To 39) As String
Dim formula As String
Dim namesOfColunms(14) As String
Dim totalIncludTaxesAndFees As New Collection

'заполнение массива буквами из наименования столбцов
For i = 13 To 39 Step 2
    letterCol(i) = Split(Cells(1, i).Address, "$")(1)
Next

namesOfColunms(1) = "Январь"
namesOfColunms(2) = "Февраль"
namesOfColunms(3) = "Март"
namesOfColunms(4) = "Апрель"
namesOfColunms(5) = "Май"
namesOfColunms(6) = "Июнь"
namesOfColunms(7) = "Июль"
namesOfColunms(8) = "Август"
namesOfColunms(9) = "Сентябрь"
namesOfColunms(10) = "Октябрь"
namesOfColunms(11) = "Ноябрь"
namesOfColunms(12) = "Декабрь"
namesOfColunms(13) = "Сумма по актам"
namesOfColunms(14) = "Остатки"

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
Range("A15:A31").EntireRow.Hidden = True
With Range("A14:K14")
    .UnMerge
    .HorizontalAlignment = xlCenterAcrossSelection
End With
Columns("F:I").Hidden = True
Columns("K:K").Hidden = True

lastCell = ContractEstimate.seachLastCell()

Set seachRange = Range("A33:H" & lastCell)

seachString = "Итого по разделу: *"
Set totalForSection = Estimate.Seach(seachString, seachRange)
Call quickSort.quickSort(totalForSection, 1, totalForSection.Count)

seachString = "Итого по локальной смете*"
Set totalByEstimate = Estimate.Seach(seachString, seachRange)

Set coefFinanc = Estimate.Seach("*коэффициент*финансиро*", seachRange)
If coefFinanc.Count <> 0 Then
     numberCoefFinance = Replace(Left(Split(Cells(coefFinanc(1), 1).Value, "=")(1), 13), ",", ".")
End If
Set coefDecline = Estimate.Seach("*коэффициент*снижен*", seachRange)
If coefDecline.Count <> 0 Then
     numberCoefDecline = Replace(Left(Split(Cells(coefDecline(1), 1).Value, "=")(1), 13), ",", ".")
End If

Set seachRange = Range("A33:H" & totalByEstimate(1) + 1)
Set rowNDS = Estimate.Seach("НДС 20*", seachRange)
Set totalIncludTaxesAndFees = Estimate.Seach("*Итого с учетом налогов и сборов*", seachRange)

Set seachRange = Range("A32:K" & lastCell)
seachString = "Раздел: *"
Set beginningOfSection = Estimate.Seach(seachString, seachRange)
Call quickSort.quickSort(beginningOfSection, 1, beginningOfSection.Count)


Range("L32").Select
ActiveWindow.FreezePanes = True

'вставка столбцов
i = 1
For j = 13 To 39 Step 2
    If j = 37 Then
        Call cumulativeList.insertCol(namesOfColunms(i), j, 13, lastCell, "230 230 250")
    ElseIf j = 39 Then
        Call cumulativeList.insertCol(namesOfColunms(i), j, 13, lastCell, "144 238 144")
    Else
        Call cumulativeList.insertCol(namesOfColunms(i), j, 13, lastCell)
    End If
        Call outputTotalForSection(j, letterCol(j))
        formula = Estimate.formulaTotal(totalForSection, letterCol(j))
        Cells(totalByEstimate(totalByEstimate.Count), j).Value = formula
        Cells(totalByEstimate(totalByEstimate.Count), j).Font.Bold = True
        formula = ""
        If rowNDS.Count <> 0 Then
            Call cumulativeList.totalWithNDS(rowNDS(rowNDS.Count) - 1, j, letterCol(j))
            If coefFinanc.Count <> 0 Then
                Cells(coefFinanc(1), j).formula = "=round(" & letterCol(j) & totalByEstimate(totalByEstimate.Count) + 2 & "*" & numberCoefFinance & ",2)"
                Call Estimate.setFormat(coefFinanc(1), j, coefFinanc(1), j)
                If coefDecline.Count <> 0 Then
                    Cells(coefDecline(1), j).formula = "=round(" & letterCol(j) & coefFinanc(1) & "*" & numberCoefDecline & ",2)"
                    Call estimateSN.NDSIncluding(coefDecline(1) + 1, j, letterCol(j))
                    Call Estimate.setFormat(coefDecline(1), j, coefDecline(1) + 1, j)
                Else
                    Call estimateSN.NDSIncluding(coefFinanc(1) + 1, j, letterCol(j))
                    Call Estimate.setFormat(coefFinanc(1) + 1, j, coefFinanc(1) + 1, j)
                End If
            ElseIf coefDecline.Count <> 0 Then
                 Cells(coefDecline(1), j).formula = "=round(" & letterCol(j) & totalByEstimate(1) + 2 & "*" & numberCoefDecline & ",2)"
                 Call estimateSN.NDSIncluding(coefDecline(1) + 1, j, letterCol(j))
                 Call Estimate.setFormat(coefDecline(1), j, coefDecline(1) + 1, j)
            End If
         ElseIf totalIncludTaxesAndFees.Count <> 0 Then
            Cells(totalIncludTaxesAndFees(1), j).formula = "=" & letterCol(j) & totalByEstimate(1) & "*1.2"
            Call Estimate.setFormat(totalIncludTaxesAndFees(1), j, totalIncludTaxesAndFees(1), j)
            If coefFinanc.Count <> 0 Then
                Cells(coefFinanc(1), j).formula = "=round(" & letterCol(j) & totalIncludTaxesAndFees(totalIncludTaxesAndFees.Count) & "*" & numberCoefFinance & ",2)"
                Call Estimate.setFormat(coefFinanc(1), j, coefFinanc(1), j)
                If coefDecline.Count <> 0 Then
                    Cells(coefDecline(1), j).formula = "=round(" & letterCol(j) & coefFinanc(1) & "*" & numberCoefDecline & ",2)"
                    Call Estimate.setFormat(coefDecline(1), j, coefDecline(1), j)
                End If
            ElseIf coefDecline.Count <> 0 Then
                Cells(coefDecline(1), j).formula = "=round(" & letterCol(j) & totalIncludTaxesAndFees(1) & "*" & numberCoefDecline & ",2)"
                Call Estimate.setFormat(coefDecline(1), j, coefDecline(1), j)
            End If
    End If
    i = i + 1
Next j


currentRow = 0

For Each currentRow In Range("A1:A" & totalByEstimate(totalByEstimate.Count))
    If currentRow.Value = 1 Then
        firstRow = currentRow.Row
        Exit For
    End If
Next

currentRow = 0

For Each currentRow In Range("A" & firstRow & ":A" & totalByEstimate(totalByEstimate.Count) - 1)
        
    If TypeName(Cells(currentRow.Row, 1).Value) = "Double" Or Cells(currentRow.Row, 1).Value Like "Итого по *разделу:*" Then
        lastRow = Cells(currentRow.Row, 1).Row
    End If
    If Cells(currentRow.Row, 10).Font.Bold = True And Not Cells(currentRow.Row, 1).Value Like "Итого по *разделу:*" Then
            rowPosition.Add currentRow.Row
            Select Case periodCoef
                Case 1
                    Cells(rowPosition(rowPosition.Count), 35).Value = Round(Cells(rowPosition(rowPosition.Count), 9).Value, 2)
                Case 11
                   For i = 13 To 34 Step 2
                       Cells(rowPosition(rowPosition.Count), i).Value = Round(Cells(rowPosition(rowPosition.Count), 9).Value / periodCoef, 2)
                   Next i
                Case 12
                    For i = 13 To 36 Step 2
                       Cells(rowPosition(rowPosition.Count), i).Value = Round(Cells(rowPosition(rowPosition.Count), 9).Value / periodCoef, 2)
                   Next i
            End Select
    End If
        
    For Each currentRow2 In Range("A" & firstRow & ":A" & lastRow)
        If IsEmpty(Cells(currentRow2.Row, 2)) And Cells(currentRow2.Row, 10).Font.Bold = False Then
            Rows(currentRow2.Row).Hidden = True
        End If
    Next
    
    If Cells(currentRow.Row, 3).Value Like "Техническое обслуживание*" Then
       periodCoef = 1
       
    End If
    If Cells(currentRow.Row, 3).Value Like "Технический осмотр*" Then
        If Cells(currentRow.Row, 7).Value = "" Then
            If Cells(currentRow.Row + 1, 7).Value <> "" Then
                If InStr(Cells(currentRow.Row + 1, 7).Value, "11") Then
                    periodCoef = 11
                End If
                If InStr(Cells(currentRow.Row + 1, 7).Value, "12") Then
                    periodCoef = 12
                End If
            End If
        Else
            If Cells(currentRow.Row, 7).Value <> "" Then
                If InStr(Cells(currentRow.Row, 7).Value, "11") Then
                    periodCoef = 11
                End If
                If InStr(Cells(currentRow.Row, 7).Value, "12") Then
                    periodCoef = 12
                End If
            End If
        End If
    End If
    firstRow = lastRow
Next

lastCell = ContractEstimate.seachLastCell()

ActiveSheet.PageSetup.PrintArea = "$A$1:$AM$" & lastCell

For Each currentRow In rowPosition
    Cells(currentRow, 37).formula = "=Sum(M" & currentRow & ":AI" & currentRow & ")"
    Call Estimate.setFormat(currentRow, 37, currentRow, 37)
Next

Call cumulativeList.difference(rowPosition, 9, "I", 37, "AK", 39)

End Sub

Sub outputTotalForSection(numberCol, letterCol)
'вывод формулы итого по разделу
Dim k As Long

k = 1
Do While k <= totalForSection.Count
    Cells(totalForSection(k), numberCol).formula = "=Sum(" & letterCol & beginningOfSection(k) + 1 & ":" & letterCol & (totalForSection(k) - 1) & ")"
    Call Estimate.setFormat(totalForSection(k), numberCol, totalForSection(k), numberCol)
    
    k = k + 1
Loop


End Sub


