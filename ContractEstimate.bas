Attribute VB_Name = "ContractEstimate"
Dim currWB As Workbook
Dim kt As String
Dim kName As String
Public typeEstimate As String




Sub ContractEstimate()
' Формирование сметы к контракту
Dim i As Integer
Dim seachRange As Range, seachStr As String
Dim foundCell As Range, firstFoundCell As Range
Dim torgName As String 'вид торгов
Dim koefType 'вид коэффициента: бюджетного финансирования или по результатам закупки
Dim sheetName As String
Dim currSheet As Variant
Dim colNumber As Integer
Dim colLetter As String

typeEstimate = InputBox("Введите тип сметы: ТСН или СН", , "ТСН")

Select Case typeEstimate
    Case "ТСН"
        colNumber = 11
        colLetter = "K"
    Case "СН"
        colNumber = 10
        colLetter = "J"
End Select
' Ввод вида торгов
Application.SendKeys ("%+") ' Переход на русский язык
kName = InputBox("Введите вид закупки: c учетом коэффициента снижения по результатам", , "открытого конкурса в электронной форме")
' Ввод коэффициента снижения
Application.SendKeys ("%+") ' Переход на английский язык
kt = simpleInput("коэффициент снижения по результатам открытого конкурса в электронной форме")
kName = "c учетом коэффициента снижения по результатам " & kName
Set currWB = ActiveWorkbook
sheetName = "Смета*"
For Each currSheet In Worksheets
    If currSheet.Name Like sheetName Then
        currSheet.Activate
    End If
Next
' Поиск последней непустой ячейки
lastRow = seachLastCell() + 1
' Установление области  и строки поиска
Set seachRange = Range("A1:I" & lastRow)
seachStr = "Итого * финансирования*"
' Поиск первого совпадения
If seachRange.Find(seachStr, LookIn:=xlValues) Is Nothing Then
   seachStr = "Итого с* НДС*"
End If
Set foundCell = seachRange.Find(seachStr, LookIn:=xlValues)
Set firstFoundCell = foundCell
If Not foundCell Is Nothing Then
'Удаление строки в т.ч. НДС, которая относится к коэффициенту снижения
    Rows(foundCell.Row + 1).EntireRow.Delete
    Call TotalWithK(kt, kName, foundCell.Row, foundCell.Row, colLetter, 1, colNumber)
    Rows(foundCell.Row + 3).EntireRow.Insert
End If
' Поиск остальных совпадений
i = 0
Do
    Set foundCell = seachRange.FindNext(After:=foundCell)
    If foundCell.Address = firstFoundCell.Address Then Exit Do
    'Удаление строки в т.ч. НДС, которая относится к коэффициенту снижения
    Rows(foundCell.Row + 1).EntireRow.Delete
    Call TotalWithK(kt, kName, foundCell.Row, foundCell.Row, colLetter, 1, colNumber)
    Rows(foundCell.Row + 3).EntireRow.Insert
    i = i + 1
Loop

' Проверка сходится ли итого
'Call checkTotal(firstFoundCell.Row, t, 9, 0)

' Вставка и заполнение ПССО
Call fillPCCO(firstFoundCell, seachStr)

End Sub

Sub TotalWithK(k, k_Name, r, r2, colLetter, colName, colNumber)
' Заполнение Итого с Коэффициентом, в том числе НДС
Rows((r + 1) & ":" & (r + 2)).Insert
'If Cells(r, colNumber).formula = "" Then
 '   colLetter = "J"
'End If
Cells(r + 1, colName).Value = "Итого с " & k_Name & " K =" & k
Cells(r + 1, colNumber).formula = "=round(" & colLetter & (r2) & "*" & (k) & ",2)"
Call estimateSN.NDSIncluding(r + 2, colNumber, colLetter)

End Sub

Sub checkTotal(r, t, shft, checkStatus)
' Проверка сходится ли итого
If checkStatus = 1 Then
    Cells(r, 13).formula = "=K" & t(0) & "+" & "K" & t(1) & "+" & "K" & t(2) & "+" & "K" & t(3)
    Do While Cells(r, 13).Value <> Cells(r, 11).Value
    
        If Cells(r, 13).Value < Cells(r, 11).Value Then
            Cells(r + shft - 1, 11).Value = Cells(r + shft - 1, 11).Value + 0.01
        Else
            Cells(r + shft - 1, 11).Value = Cells(r + shft - 1, 11).Value - 0.01
        End If
    Loop
Else
    Cells(r + 1, 13).formula = "=K" & t(0) + 1 & "+" & "K" & t(1) + 1 & "+" & "K" & t(2) + 1 & "+" & "K" & t(3) + 1
    Cells(r + 2, 13).formula = "=K" & t(0) + 2 & "+" & "K" & t(1) + 2 & "+" & "K" & t(2) + 2 & "+" & "K" & t(3) + 2
    Do While Cells(r + 1, 13).Value <> Cells(r + 1, 11).Value
        If Cells(r + 1, 13).Value < Cells(r + 1, 11).Value Then
            Cells(r + shft, 11).formula = Cells(r + shft, 11).formula & "+0.01"
        Else
            Cells(r + shft, 11).formula = Cells(r + shft, 11).formula & "-0.01"
        End If
    Loop
    Do While Cells(r + 2, 13).Value <> Cells(r + 2, 13).Value
        If Cells(r + 2, 13).Value < Cells(r + 2, 11).Value Then
            Cells(r + shft + 1, 11).formula = Cells(r + shft + 1, 11).formula & "+0.01"
        Else
            Cells(r + shft + 1, 11).formula = Cells(r + shft + 1, 11).formula & "-0.01"
        End If
    Loop
End If
End Sub

Sub fillPCCO(r, seachString)
' Заполнение ПССО
Sheets.Add Before:=Sheets(1), Type:="C:\Гончарова\эксель\черновик\шаблоны\ПССО.xltx"

totalD = Sheets(5).Range("K" & r.Row) _
        - Sheets(5).Range("K" & r.Row + 1)
Range("A9:C9").Value = Sheets(2).Range("A9:H9").Value
Range("A14").Value = "Снижение стоимости выполнения подрядных работ по результатам электронного конкурса составляет " _
        & Int(totalD) & " руб. " & (totalD * 100 Mod 100) & " коп."

Sheets("ПССО").Range("C16") = "в ценах " & Estimate.seachMonthYear("year", currWB, typeEstimate) & ", руб."
currWB.Sheets(1).Activate
Range("B23") = Sheets(5).Range("K" & r.Row)
If seachString = "Итого * финансирования*" Then
    Cells(24, 1) = Sheets(5).Range("A" & r.Row)
    Cells(24, 2) = Sheets(5).Range("K" & r.Row)
    Call TotalWithK(kt, kName, 24, 24, "B", 1, 2)
    Range("C24:C26") = Range("B24:B26").Value
    With Range("A24:A26")
        .WrapText = True
        .HorizontalAlignment = xlLeft
    End With
    With Range("A24:C26")
        .Borders.LineStyle = xlContinuous
        .Font.Size = 11
    End With
Else
    Call TotalWithK(kt, kName, 23, 23, "B", 1, 2)
    Range("C24") = Range("B24")
    Range("C25") = Range("B25")
    Range("A24:C25").Borders.LineStyle = xlContinuous
End If
Rows("24:24").RowHeight = 42

Call Estimate.setFormat(24, 2, 26, 3)
End Sub

Function simpleInput(ktype) As String
' Простой ввод: коэффициенты, наименование закупки
strInput = "Введите " & ktype
strInput2 = "Ввод: " & ktype
Do Until simpleInput <> ""
    simpleInput = InputBox(strInput, strInput2)
Loop
End Function

Function seachLastCell()
' поиск последней непустой ячейки в столбцах с 1-го по 11-й
    Dim c(11) As Integer
    For i = 1 To 11
        c(i) = Cells(Rows.Count, i).End(xlUp).Row
    Next
    seachLastCell = WorksheetFunction.Max(c)
End Function









