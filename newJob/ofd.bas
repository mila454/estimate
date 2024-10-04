Attribute VB_Name = "ofd"
Option Explicit
Dim column_name As String
Dim column_list(4) As String
Dim i As Integer
Dim item As Variant
Dim match As Boolean
Dim lastCell As Integer



Sub design_list()
'редактирование файла, загруженного из ОФД

column_list(0) = "Название кассы"
column_list(1) = "Дата/время открытия смены"
column_list(2) = "Итоговая сумма расчета"
column_list(3) = "Сумма расчета наличными"
column_list(4) = "Сумма расчета безналичными (эквайринг)"

Range(Cells(1), Cells(1)).EntireRow.Delete

'удаление столбцов, кроме пяти вышеназванных

For i = 1 To 5
    For Each item In column_list
    
        If Cells(1, i).Value <> item Then
            
            match = False
        Else
            match = True
            Exit For
        End If
        
    Next
    
    If match = False Then
        Range(Cells(i), Cells(i)).EntireColumn.Delete
        i = i - 1
    End If
Next

Range("F1:BV1").EntireColumn.Delete

'форматирование столбца Дата
lastCell = seachLastCell()

With Range(Cells(2, 2), Cells(lastCell, 2))
    .NumberFormat = "m/d/yyyy"
    
End With

For i = 2 To lastCell
    Cells(i, 2).Value = DateSerial(year(Cells(i, 2)), Month(Cells(i, 2)), Day(Cells(i, 2)))
Next


'создание сводной таблицы
Range(Cells(1, 1), Cells((lastCell - 1), 6)).Select
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Отчет!R1C1:R129C5", Version:=6).CreatePivotTable TableDestination:= _
        "Лист1!R3C1", TableName:="Сводная таблица1", DefaultVersion:=6
    Sheets("Лист1").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("Сводная таблица1").PivotFields("Название кассы")
        .Orientation = xlRowField
        .position = 1
    End With
    ActiveSheet.PivotTables("Сводная таблица1").AddDataField ActiveSheet. _
        PivotTables("Сводная таблица1").PivotFields("Итоговая сумма расчета"), _
        "Сумма по полю Итоговая сумма расчета", xlSum
    ActiveSheet.PivotTables("Сводная таблица1").AddDataField ActiveSheet. _
        PivotTables("Сводная таблица1").PivotFields("Сумма расчета наличными"), _
        "Сумма по полю Сумма расчета наличными", xlSum
    ActiveSheet.PivotTables("Сводная таблица1").AddDataField ActiveSheet. _
        PivotTables("Сводная таблица1").PivotFields( _
        "Сумма расчета безналичными (эквайринг)"), _
        "Сумма по полю Сумма расчета безналичными (эквайринг)", xlSum
    With ActiveSheet.PivotTables("Сводная таблица1").PivotFields( _
        "Дата/время открытия смены")
        .Orientation = xlPageField
        .position = 1
    End With
End Sub



Function seachLastCell()
' поиск последней непустой ячейки в столбцах с 1-го по 5-й
    Dim c(5) As Integer
    For i = 1 To 5
        c(i) = Cells(Rows.Count, i).End(xlUp).row
    Next
    seachLastCell = WorksheetFunction.Max(c)
End Function


