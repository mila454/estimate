Attribute VB_Name = "BudgetFinancingForSN"

Option Explicit

Dim currWB As Workbook


Sub coefficientOfBudgetFinancingForSN()
'добавление в смету СН без НМЦК коэффициента бюджетного финансирования
Dim seachRange As Range, seachStr As String
Dim foundCell As Range, firstFoundCell As Range
Dim lastRow As Integer
Dim kt As String
Dim kName As String

' Ввод коэффициента снижения
kName = "коэффициентом бюджетного финансирования"
kt = ContractEstimate.simpleInput(kName)
Set currWB = ActiveWorkbook
currWB.Sheets(1).Activate

' Поиск последней непустой ячейки
lastRow = ContractEstimate.seachLastCell() + 1

' Установление области  и строки поиска
Set seachRange = Range("A1:I" & lastRow)
' Поиск первого совпадения
seachStr = "Итого с* НДС*"
Set foundCell = seachRange.Find(seachStr, LookIn:=xlValues)
Set firstFoundCell = foundCell
If Not foundCell Is Nothing Then
    Call ContractEstimate.TotalWithK(kt, kName, foundCell.Row, foundCell.Row, "J", 1, 10)
End If
' Поиск остальных совпадений
Do
    Set foundCell = seachRange.FindNext(After:=foundCell)
    If foundCell.Address = firstFoundCell.Address Then Exit Do
    Call ContractEstimate.TotalWithK(kt, kName, foundCell.Row, foundCell.Row, "J", 1, 10)
Loop

End Sub


