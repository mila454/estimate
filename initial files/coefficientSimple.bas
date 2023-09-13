Attribute VB_Name = "coefficientSimple"
Option Explicit

Dim currWB As Workbook
Dim kt As String
Dim kName As String

Sub coefficientOfBudgetFinancingSimple()
'добавление только в смету коэффициента бюджетного финансирования
Dim seachRange As Range, seachStr As String
Dim foundCell As Range, firstFoundCell As Range
Dim lastRow As Integer
Dim rangeBorders As Range
Dim sheetName As String
Dim currSheet As Variant
Dim typeEstimate As String
Dim colNumber As Integer
Dim colLetter As String

typeEstimate = "ТСН"

Select Case typeEstimate
    Case "ТСН"
        colNumber = 11
        colLetter = "K"
    Case "СН"
        colNumber = 10
        colLetter = "J"
End Select
' Ввод коэффициента снижения
kName = "коэффициентом снижения по результатам закупки"
kt = ContractEstimate.simpleInput(kName)
Set currWB = ActiveWorkbook
sheetName = "Смета*"
For Each currSheet In Worksheets
    If currSheet.Name Like sheetName Then
        currSheet.Activate
    End If
Next
' Поиск последней непустой ячейки
lastRow = ContractEstimate.seachLastCell() + 1

' Установление области  и строки поиска
Set seachRange = Range("A1:I" & lastRow)
' Поиск первого совпадения
seachStr = "Итого с* НДС*"
Set foundCell = seachRange.Find(seachStr, LookIn:=xlValues)
Set firstFoundCell = foundCell
If Not foundCell Is Nothing Then
    Call ContractEstimate.TotalWithK(kt, kName, foundCell.Row, foundCell.Row, colLetter, 1, colNumber)
End If
' Поиск остальных совпадений
Do
    Set foundCell = seachRange.FindNext(After:=foundCell)
    If foundCell.Address = firstFoundCell.Address Then Exit Do
    Call ContractEstimate.TotalWithK(kt, kName, foundCell.Row, foundCell.Row, colLetter, 1, colNumber)
Loop



End Sub




