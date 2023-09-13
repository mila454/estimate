Attribute VB_Name = "heightAdjustment"
Option Explicit

Sub heightAdjustment(mergedRange)
'автоподбор высоты объединенных €чеек
Dim myCell As Range, myLen As Integer, _
myWidth As Single, k As Single, n As Single
    With mergedRange
        '«адаем объединенной €чейке перенос текста
        .WrapText = True
        '«адаем объединенной €чейке такую высоту строки,
        'чтобы умещалась одна строка текста
        .RowHeight = Cells(mergedRange.Row, mergedRange.Column).Font.Size * 1.3
    End With
myLen = Len(CStr(Cells(mergedRange.Row, mergedRange.Column)))
    For Each myCell In mergedRange
        myWidth = myWidth + myCell.ColumnWidth
    Next
n = 10
k = Cells(mergedRange.Row, mergedRange.Column).Font.Size / n
mergedRange.RowHeight = mergedRange.RowHeight * _
WorksheetFunction.RoundUp(myLen * k / myWidth, 0)
End Sub

