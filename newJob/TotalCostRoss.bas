Attribute VB_Name = "TotalCostRoss"
Sub TotalCostRoss()

Dim currentColumn As Integer
Dim j As Integer
Dim letterCol As String


currentColumn = ActiveCell.Column

j = 10

'заполнение столбца Стоимость в накопительной'
For j = 10 To 588
    If Not IsEmpty(Cells(j, currentColumn).Value) Then
        letterCol = Split(Cells(1, currentColumn).Address, "$")(1)
        Cells(j, currentColumn + 1).Formula = "=round(E" & j & "*" & letterCol & j & ",2)"
    End If
Next

Application.CutCopyMode = False

End Sub



