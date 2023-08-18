Attribute VB_Name = "coefficientOfBudgetFinancing"
Option Explicit

Dim currWB As Workbook
Dim kt As String
Dim kName As String

Sub coefficientOfBudgetFinancing()
'добавление в смету и ПНЦ коэффициента бюджетного финансирования
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
kName = "коэффициентом бюджетного финансирования"
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


'sheetName = "РНЦ*"
'For Each currSheet In Worksheets
'    If currSheet.Name Like sheetName Then
'        currSheet.Activate
'    End If
'Next

'If ActiveSheet.Name = "РНЦ" Then
 '   Call ContractEstimate.TotalWithK(kt, kName, 21, 21, "B", 1, 2)
  '  Cells(22, 1).WrapText = True
   ' Rows("22:22").EntireRow.AutoFit
    'Range("B22").Copy
'    Range("C22:G22").Select
'    ActiveSheet.Paste
'    Range("B23").Copy
'    Range("C23:G23").Select
'    ActiveSheet.Paste
'    Range("H22:H22").FillDown
'    Range("H23:H23").FillDown
'    Set rangeBorders = Range("A22:H23")
'    rangeBorders.Borders.LineStyle = xlContinuous
'End If
End Sub



