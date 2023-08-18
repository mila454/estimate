Attribute VB_Name = "EstimateSimple"

Option Explicit
Public currWB As Workbook
Dim smetaName As String
Dim plantingTrees As New Collection, restTrees As New Collection, treesCare1 As New Collection, treesCare2 As New Collection, treesCare3 As New Collection
Dim totalEstimate As New Collection
Public kindOfWorks As New Collection
Dim god1 As String, god2 As String, god3 As String
Public typeEstimate As String
Dim sheetName As String


Sub EstimateSimple()
' Формирование хвостов сметы после выгрузки из программы

Dim seachRange As Range, seachStr As String
Dim lastRow As Integer
Dim i As Integer
Dim currSheet As Variant
Dim accept As New Collection
Dim colLetter As String
Dim colNumber As Integer

Set currWB = ActiveWorkbook
typeEstimate = InputBox("Введите тип сметы: ТСН или СН", , "ТСН")
If typeEstimate = "ТСН" Then
    colLetter = "K"
    colNumber = 11
Else
    colLetter = "J"
    colNumber = 10
End If

god1 = InputBox("Введите первый год ухода", , "2024")
god2 = InputBox("Введите второй год ухода", , "2025")
god3 = InputBox("Введите третий год ухода", , "2026")
sheetName = "Смета*"

For Each currSheet In Worksheets
    If currSheet.Name Like sheetName Then
        currSheet.Activate
    End If
Next
' Удаление и снятие объединения строк после итого по смете
lastRow = ContractEstimate.seachLastCell()
Set seachRange = Range(Cells(1, 1), Cells(lastRow, 9))

seachStr = "Итого по*смете*"
Set totalEstimate = Seach(seachStr, seachRange)
Call quickSort.quickSort(totalEstimate, 1, totalEstimate.Count)

seachStr = "Проверил*"
Set accept = Seach(seachStr, seachRange)

' Сохранение названия сметы
If Sheets("Source").Range("F20") <> "" Then
    smetaName = Sheets("Source").Range("G20")
End If

Range("A" & totalEstimate(1)) = "Итого по локальной смете №1: " & smetaName
Call cancelMerge(Split(Range(colLetter & totalEstimate(1)).Offset(, -1).Address, "$")(1), totalEstimate(1), colLetter, totalEstimate(1), 0)
Cells(totalEstimate(1), colNumber).formula = "=SUM(P36:P" & totalEstimate(1) - 1 & ")"


' Шапка и название сметы
Call header(smetaName)

Range("G36:K36").EntireRow.Hidden = True
Worksheets("Source").Cells(1, 10).Clear
Worksheets("SourceObSm").Cells(1, 10).Clear

' Расчет хвоста
Call estimateTail(seachRange, god1, god2, god3)

' Отображение хвоста
Range("A" & (totalEstimate(1)) & ":A" & (totalEstimate(1) + 2)).EntireRow.Hidden = False

'Call heightAdjustment.heightAdjustment(Range("A" & totalEstimate(1) & ":G" & totalEstimate(1)))
Range("A" & totalEstimate(1) + 1 & ":A" & accept(1) + 1).EntireRow.Delete

If typeEstimate = "ТСН" Then
    Cells(totalEstimate(1) + 2, 9).formula = "=H" & totalEstimate(1) & "+H" & (totalEstimate(1) + 1)
    Cells(totalEstimate(1) + 1, 9).formula = "=round(H" & (totalEstimate(1)) & "*0.2,2)"
    Call setFormat(totalEstimate(1) + 1, 9, totalEstimate(1) + 2, 9)
End If
Call displayTail(totalEstimate(1), Split(Range(colLetter & totalEstimate(1)).Offset(, -1).Address, "$")(1), colLetter, colNumber, god1, god2, god3)



' Проверка хвоста
Cells(totalEstimate(1), 12).formula = "=" & colLetter & kindOfWorks(1) & "+" & colLetter & kindOfWorks(2) & "+" & colLetter & kindOfWorks(3) & "+" & colLetter & kindOfWorks(4)
Cells(totalEstimate(1) + 1, 12).formula = "=" & colLetter & kindOfWorks(1) + 1 & "+" & colLetter & kindOfWorks(2) + 1 & "+" & colLetter & kindOfWorks(3) + 1 & "+" & colLetter & kindOfWorks(4) + 1
Cells(totalEstimate(1) + 2, 12).formula = "=K" & kindOfWorks(1) + 2 & "+" & colLetter & kindOfWorks(2) + 2 & "+" & colLetter & kindOfWorks(3) + 2 & "+" & colLetter & kindOfWorks(4) + 2
Columns("L:L").EntireColumn.AutoFit

'Корректировка ведоости
'Call Estimate.listOfWorks

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

Function formulaTotal(coll, letterColl) As String
'формирование строки с формулой суммы по строкам из коллекции
Dim i As Integer
formulaTotal = "="
For i = 1 To coll.Count
    If i = coll.Count Then
        formulaTotal = formulaTotal & letterColl & coll(i)
    Else
        formulaTotal = formulaTotal & letterColl & coll(i) & "+"
    End If
Next i

End Function

Sub ndsTotal(r, indCol, numCol)
'расчет и вывод НДС и итого с НДС

Cells(r + 1, 1) = "НДС 20%"
Cells(r + 1, numCol).formula = "=round(" & indCol & (r) & "*0.2,2)"

Cells(r + 2, 1) = "Итого с НДС 20%"
Cells(r + 2, numCol).formula = "=" & indCol & r & "+" & indCol & (r + 1)

Range("A" & r & ":" & indCol & (r + 2)).Font.Bold = True
Call setFormat(r, numCol, r + 2, numCol)

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

Sub setFormat(row1Format, col1Format, row2Format, col2Format)
'форматирование диапазона
With Range(Cells(row1Format, col1Format), Cells(row2Format, col2Format))
    .Font.Bold = True
    .NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
End With

End Sub

Sub header(smetaName)
'Формирование шапки
With Range(Cells(3, 2), Cells(7, 11))
    .UnMerge
    .Clear
    
End With

Cells(3, 2) = Chr(34) & "УТВЕРЖДАЮ" & Chr(34)
Cells(5, 2) = "Заказчик:"
Cells(6, 2) = "Заместитель директора ГКУ г. Москвы " & Chr(34) & "Дирекция Мосприроды" & Chr(34) & "_________________________Е.И.Новощинская"

With Range(Cells(6, 2), Cells(6, 5))
   .MergeCells = True
   .WrapText = True
End With

Call heightAdjustment.heightAdjustment(Range("6:6"))

Range("A10, A15") = smetaName
Call heightAdjustment.heightAdjustment(Range("10:10"))
Call heightAdjustment.heightAdjustment(Range("15:15"))

Cells(13, 1) = "ЛОКАЛЬНАЯ СМЕТА № 1"
With Range(Cells(3, 2), Cells(13, 2))
    .Font.Name = "Arial"
    .Font.Size = 13
End With
Range(Cells(7, 7), Cells(7, 11)).Clear

End Sub

Public Sub estimateTail(seachRange, god1, god2, god3)
'формирование коллекции номеров строк итогов по видам работ
Dim seachStr As String
Dim i As Integer

seachStr = "Итого по разделу: Посадка*"
Set plantingTrees = Seach(seachStr, seachRange)

seachStr = "Итого по разделу: Восстановление отпада*"
Set restTrees = Seach(seachStr, seachRange)

seachStr = "Итого по разделу: Уход*" & god1
Set treesCare1 = Seach(seachStr, seachRange)
For i = 1 To treesCare1.Count
    restTrees.Add treesCare1(i)
Next i
    
seachStr = "Итого по разделу: Уход*" & god2
Set treesCare2 = Seach(seachStr, seachRange)

seachStr = "Итого по разделу: Уход*" & god3
Set treesCare3 = Seach(seachStr, seachRange)
Call quickSort.quickSort(plantingTrees, 1, plantingTrees.Count)
Call quickSort.quickSort(restTrees, 1, restTrees.Count)
Call quickSort.quickSort(treesCare2, 1, treesCare2.Count)
Call quickSort.quickSort(treesCare3, 1, treesCare3.Count)

End Sub

Public Sub displayTail(lastRow, letterCol, letterCol2, numberCol, god1, god2, god3)
Set kindOfWorks = New Collection
Call ndsTotal(lastRow, letterCol2, numberCol)
Call setFormat(lastRow + 1, numberCol, lastRow + 2, numberCol)
Cells(lastRow + 4, 1) = "В том числе:"
Cells(lastRow + 6, 1) = "Посадка деревьев (" & god1 & " год)"
Cells(lastRow + 6, numberCol) = formulaTotal(plantingTrees, letterCol)
kindOfWorks.Add (lastRow + 6)
lastRow = lastRow + 6
Call ndsTotal(lastRow, letterCol2, numberCol)
Cells(lastRow + 4, 1) = "Восстановительные и уходные работы (" & god1 & " год)"
Cells(lastRow + 4, numberCol).Value = formulaTotal(restTrees, letterCol)
kindOfWorks.Add (lastRow + 4)
lastRow = lastRow + 4
Call ndsTotal(lastRow, letterCol2, numberCol)
Cells(lastRow + 4, 1) = "Уходные работы (" & god2 & " год)"
Cells(lastRow + 4, numberCol).Value = formulaTotal(treesCare2, letterCol)
kindOfWorks.Add (lastRow + 4)
lastRow = lastRow + 4
Call ndsTotal(lastRow, letterCol2, numberCol)
Cells(lastRow + 4, 1) = "Уходные работы (" & god3 & " год)"
Cells(lastRow + 4, numberCol).Value = formulaTotal(treesCare3, letterCol)
kindOfWorks.Add (lastRow + 4)
lastRow = lastRow + 4
Call ndsTotal(lastRow, letterCol2, numberCol)
Range("A" & lastRow - 17 & ":C" & (lastRow + 2)).Font.Bold = True

End Sub









