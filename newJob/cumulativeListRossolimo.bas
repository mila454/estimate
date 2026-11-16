Attribute VB_Name = "cumulativeListRossolimo"
Option Explicit
Dim currentColumn As Integer
Dim prevMonthDate As Date
Dim i As Integer
Dim j As Integer
Dim strokaFormul As String
Dim letterCol As String






Sub cumulativeListRossolimo()

' черновик
' накопительная по Россолимо



currentColumn = ActiveCell.Column

prevMonthDate = DateAdd("m", -1, Date)

'вставка строк и заполнение шапки'
Cells(, currentColumn).EntireColumn.Insert
Cells(, currentColumn).EntireColumn.Insert
Range(Cells(8, currentColumn), Cells(8, currentColumn + 1)).Merge
Range(Cells(8, currentColumn), Cells(8, currentColumn + 1)).Value = Format(prevMonthDate, "MMMM YYYY")
Cells(9, currentColumn).Value = "количество"
Cells(9, currentColumn + 1).Value = "стоимость, руб."

'заполнение формул по позициям'

For i = 10 To 130
    If Not IsEmpty(Cells(i, currentColumn).Value) Then
        strokaFormul = Cells(i, currentColumn + 2).Formula
        letterCol = Right(Left(Cells(i, currentColumn).Address, 2), 1)
        Cells(i, currentColumn + 2).Formula = strokaFormul & "+" & letterCol & i
        strokaFormul = Cells(i, currentColumn + 3).Formula
        letterCol = Right(Left(Cells(i, currentColumn + 1).Address, 2), 1)
        Cells(i, currentColumn + 3).Formula = strokaFormul & "+" & letterCol & i
    End If
Next

'заполнение хвоста'

For Each item In Range(Cells(131, currentColumn + 1), Cells(139, currentColumn + 1))
    Cells(item.row, currentColumn - 1).Copy
    Cells(item.row, currentColumn + 1).PasteSpecial xlFormulas
Next

Application.CutCopyMode = False


End Sub


