Attribute VB_Name = "callFillNMCK"
Public currWB As Workbook
Public position As New Collection
Public sheetName As String
Public smetaName As String, mon As String, year As String




Sub callFillNMCK()
Dim typeEstimate As String
Dim seachRange As Range, seachStr As String
Dim currSheet As Variant
Dim lastRow As Integer
Dim tempPosition As New Collection

Set currWB = ActiveWorkbook


typeEstimate = InputBox("Введите тип сметы: ТСН или СН", , "ТСН")
If typeEstimate = "ТСН" Then
    colLetter = "K"
    colNumber = 11
Else
    colLetter = "J"
    colNumber = 10
End If

mon = Estimate.seachMonthYear("month", currWB, typeEstimate)
year = Estimate.seachMonthYear("year", currWB, typeEstimate)

' Сохранение названия сметы
If Sheets("Source").Range("F20") <> "" Then
    smetaName = Sheets("Source").Range("G20")
End If

sheetName = "Смета*"
For Each currSheet In Worksheets
    If currSheet.Name Like sheetName Then
        currSheet.Activate
        sheetName = ActiveSheet.Name
    End If
Next

lastRow = ContractEstimate.seachLastCell()
Set seachRange = Range(Cells(1, 1), Cells(lastRow, 9))

seachStr = "Итого по*смете*"
Set tempPosition = Estimate.Seach(seachStr, seachRange)
position.Add tempPosition(1)

Set seachRange = Range(Cells(position(1) + 1, 1), Cells(lastRow, 9))

seachStr = "Посадка*"
Set tempPosition = New Collection
Set tempPosition = Estimate.Seach(seachStr, seachRange)
position.Add tempPosition(1)

seachStr = "Восстановительные*"
Set tempPosition = New Collection
Set tempPosition = Estimate.Seach(seachStr, seachRange)
position.Add tempPosition(1)

seachStr = "Уход*"
Set seachRange = Range(Cells(position(3) + 1, 1), Cells(lastRow, 9))

Set tempPosition = New Collection
Set tempPosition = Estimate.Seach(seachStr, seachRange)
position.Add tempPosition(1)
position.Add tempPosition(2)
Call quickSort.quickSort(position, 1, position.Count)

Call fillNMCK

End Sub

Sub fillNMCK()
'вставка листа НМЦК и заполнение его
'Dim strSheetName As String
Dim letterCol As String

Sheets.Add Before:=Sheets(1), Type:="C:\Гончарова\эксель\черновик\шаблоны\НМЦК.xltx"

Sheets("НМЦК").Activate
Sheets("НМЦК").Name = "РНЦ"

Cells(9, 1) = smetaName
Cells(15, 2) = "Утвержденная сметная стоимость строительства в текущем уровне цен на " & mon & " " & year & " г."
letterCol = "K"
'strSheetName = sheetName
'If typeEstimate = "ТСН" Then
'    strSheetName = "Смета по ТСН-2001(с доп.67"
'    letterCol = "K"
'ElseIf estimateSN.typeEstimate = "СН" Then
'    strSheetName = "Смета СН-2012 по гл. 1-5"
'    letterCol = "J"
'End If

Cells(18, 2).formula = "='" & sheetName & "'!" & letterCol & position(1)
Cells(18, 4).formula = "='" & sheetName & "'!" & letterCol & position(2)
Cells(18, 5) = "='" & sheetName & "'!" & letterCol & position(3)
Cells(18, 6) = "='" & sheetName & "'!" & letterCol & position(4)
Cells(18, 7) = "='" & sheetName & "'!" & letterCol & position(5)

End Sub


