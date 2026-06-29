Attribute VB_Name = "cumulativeListZvenigorod"
Option Explicit
Dim currentColumm As Integer
Dim prevMonthDate As Date
Dim totalForSection As New collection
Dim beginningOfSection As New collection
Dim i As Integer
Dim j As Integer
Dim strokaFormul As String
Dim letterCol As String






Sub cumulativeListZvenigorod()

' рабочая
' накопительная по Реконструкции без автоматического заполнения столбца Стоимость


currentColumm = ActiveCell.Column

prevMonthDate = DateAdd("m", -1, Date)

'вставка строк и заполнение шапки'
Cells(, currentColumm).EntireColumn.Insert
Cells(, currentColumm).EntireColumn.Insert
Range(Cells(8, currentColumm), Cells(8, currentColumm + 1)).Merge
Range(Cells(8, currentColumm), Cells(8, currentColumm + 1)).Value = Format(prevMonthDate, "MMMM YYYY")
Cells(9, currentColumm).Value = "количество"
Cells(9, currentColumm + 1).Value = "стоимость, руб."

'заполнение формул по позициям'

Set seachRange = Range("A1:B594")
seachString = "Раздел *"
Set beginningOfSection = Seach(seachString, seachRange, "row")
seachString = "Итого по разделу *"
Set totalForSection = Seach(seachString, seachRange, "row")
Call quickSort.quickSort(beginningOfSection, 1, beginningOfSection.Count)
Call quickSort.quickSort(totalForSection, 1, totalForSection.Count)

For i = 1 To beginningOfSection.Count
    For j = beginningOfSection(i) + 1 To totalForSection(i) - 1
        strokaFormul = Cells(j, currentColumm + 2).Formula
        letterCol = Right(Left(Cells(j, currentColumm).Address, 2), 1)
        Cells(j, currentColumm + 2).Formula = strokaFormul & "+" & letterCol & j
        strokaFormul = Cells(j, currentColumm + 3).Formula
        letterCol = Right(Left(Cells(j, currentColumm + 1).Address, 2), 1)
        Cells(j, currentColumm + 3).Formula = strokaFormul & "+" & letterCol & j
        
    Next
Next

i = 1
'заполнение итогов по разделам'

Do While i <= totalForSection.Count
    Cells(totalForSection(i), currentColumm + 1).Formula = "=Sum(" & letterCol & beginningOfSection(i) & ":" & letterCol & (totalForSection(i) - 1) & ")"
    i = i + 1
Loop

'заполнение хвоста'

For Each item In Range(Cells(591, currentColumm + 1), Cells(594, currentColumm + 1))
    Cells(item.row, currentColumm - 1).Copy
    Cells(item.row, currentColumm + 1).PasteSpecial xlFormulas
Next

Application.CutCopyMode = False




'заполнение столбца Стоимость'

'If Not IsEmptyCells(j, currentColumm).Value Then
'    Cells(j, currentColumm - 1).Copy
'    Cells(j, currentColumm + 1).PasteSpecial xlFormulas
'End If

End Sub

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
        Seach.Add foundCell.row
    Else
        Seach.Add foundCell.Offset(0, token).Value
    End If
    
Loop While foundCell.Address <> firstFoundCell.Address

End Function
