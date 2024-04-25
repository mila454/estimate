Attribute VB_Name = "cumulativeList1"

Option Explicit
Dim lastCell As Integer
Dim totalByPosition As New collection
'Dim beginningOfSection As New Collection'
Dim totalForSection As New collection
Dim totalByEstimate As New collection
Dim coefMat As New collection
Dim coefMeh As New collection
Dim coefTransp As New collection
Dim seachRange As Range
Dim seachString As String
Dim i As Integer
Dim j As Integer
Dim rowEM As Integer
Dim rowOT As Integer
Dim rowM As Integer
Dim rowHP As Integer
Dim rowCP As Integer
Dim item As Variant


Sub filTotalForPosition()
'заполнение итого по позиции в текущих ценах'
lastCell = seachLastCell()

Set seachRange = Range("A1:L" & lastCell)
'seachString = "Раздел: *"'
'Set beginningOfSection = Estimate.Seach(seachString, seachRange)'
seachString = "Итого по разделу *"
Set totalForSection = Seach(seachString, seachRange, "row")
seachString = "Всего по позиции"
Set totalByPosition = Seach(seachString, seachRange, "row")
seachString = "ВСЕГО по смете*"
Set totalByEstimate = Seach(seachString, seachRange, "row")
'Call quickSort.quickSort(beginningOfSection, 1, beginningOfSection.Count)'
Call quickSort.quickSort(totalForSection, 1, totalForSection.Count)
Call quickSort.quickSort(totalByPosition, 1, totalByPosition.Count)
Call quickSort.quickSort(totalByEstimate, 1, totalByEstimate.Count)

Set coefMeh = Seach("эксплуатация машин и механизмов", seachRange, 3)
Set coefMat = Seach("материальные ресурсы", seachRange, 3)
Set coefTransp = Seach("перевозка", seachRange, 3)


Call removeItemsFromCollection(coefMeh)
Call removeItemsFromCollection(coefMat)
Call removeItemsFromCollection(coefTransp)

For j = 1 To totalByPosition.Count
    If j = 1 Then
        For i = 1 To totalByPosition(1)
        Call filCurrentPrices(i)
        Next
    Else
        For i = totalByPosition(j - 1) To totalByPosition(j)
        Call filCurrentPrices(i)
        Next
    End If
    Cells(totalByPosition(j), 12).formula = "= L" & rowOT & "+L" & rowEM & "+L" & rowM & "+L" & rowHP & "+L" & rowCP
Next

End Sub
Function seachLastCell()
' поиск последней непустой ячейки в столбцах с 1-го по 12-й
    Dim c(12) As Integer
    For i = 1 To 12
        c(i) = Cells(Rows.Count, i).End(xlUp).Row
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
        Seach.Add foundCell.Row
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
        Case "ОТ"
            rowOT = i
        Case "ЭМ"
            rowEM = i
            Cells(rowEM, 11).Value2 = coefMeh(1)
            Cells(rowEM, 12).formula = "=round(K" & rowEM & "*J" & rowEM & ",2)"
        Case "М"
            rowM = i
            Cells(rowM, 11).Value2 = coefMat(1)
            Cells(rowM, 12).formula = "=round(K" & rowM & "*J" & rowM & ",2)"
        Case "ФОТ"
            rowHP = i + 1
            rowCP = i + 2
End Select

End Sub


