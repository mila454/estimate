Attribute VB_Name = "specification"
Option Explicit

Sub specification()

Dim lastCell As Integer
Dim foundCell As Range
Dim firstFoundCell As Range
Dim seachStr As String
Dim seachRange As Range
Dim currRow As Variant
Dim total As New Collection
Dim i As Integer
Dim mergedCells As Range

lastCell = ContractEstimate.seachLastCell()
seachStr = "Страна происхождения*"
Set seachRange = Range("A1:K" & lastCell)
Set foundCell = seachRange.Find(seachStr, LookIn:=xlValues)
Set firstFoundCell = foundCell

Set foundCell = seachRange.Find(seachStr, LookIn:=xlValues)
Set firstFoundCell = foundCell

If firstFoundCell Is Nothing Then
    MsgBox (seachStr & " не найдено")
    Exit Sub
End If

Do
    Set foundCell = seachRange.FindNext(After:=foundCell)
    Rows(foundCell.Row).Interior.color = 65535
Loop While foundCell.Address <> firstFoundCell.Address

For Each currRow In seachRange
'Debug.Print seachRange.Address
    If Rows(currRow.Row).Interior.color = 65535 Then
        Rows(currRow.Row).EntireRow.Delete
    End If
Next

Range("A6:A" & lastCell).RowHeight = 35

seachStr = "*ИТОГО*"

Set seachRange = Range("A1:K" & lastCell)

Set total = Estimate.Seach(seachStr, seachRange)
Call quickSort.quickSort(total, 1, total.Count)


'Application.FindFormat.MergeCells = True
'Set mergedCells = Range("H" & total(1)).MergeArea
'Debug.Print mergedCells.Address

'If Range("H" & total(1)).MergeCells Then
'   mergedCells.UnMerge
'End If

'For i = 2 To 4
'Cells.Replace What:=".", Replacement:=",", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=True, ReplaceFormat:=False


For i = 2 To 4
    Set seachRange = Range("I" & total(i - 1) + 2 & ":K" & total(i) - 1)

    For Each currRow In seachRange
        currRow.Value = Val(currRow.Value)
    Next
    Rows(total(i)).EntireRow.Delete
    Rows(total(i)).EntireRow.Insert
    Cells(total(i), 11).formula = "=sum(K" & total(i - 1) + 2 & ":K" & total(i) - 1 & ")"
    Cells(total(i), 9).formula = "=sum(I" & total(i - 1) + 2 & ":I" & total(i) - 1 & ")"
    Call Estimate.setFormat(total(i), 9, total(i), 11)
Next
Set seachRange = Range("I6:K" & total(1))

For Each currRow In seachRange
    'currRow.Value = Replace(currRow.Value, ".", ",")
    currRow.Value = Val(currRow.Value)
Next

Rows(total(1)).EntireRow.Delete
Rows(total(1)).EntireRow.Insert
Cells(total(1), 11).formula = "=sum(K6:K" & total(1) - 1 & ")"
Call Estimate.setFormat(total(1), 9, total(1), 11)
Cells(total(1), 9).formula = "=sum(I6:I" & total(1) - 1 & ")"

End Sub
