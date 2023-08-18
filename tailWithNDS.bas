Attribute VB_Name = "tailWithNDS"

Sub tailWithNDS()
'добавление в смету НДС Итого с НДС
Dim tailRow As Integer
Dim numberCol As Integer
Dim letterCol As String

Range(Cells(15, 1), Cells(15, 11)).ClearContents
Worksheets("Source").Cells(1, 10).Clear
Worksheets("SourceObSm").Cells(1, 10).Clear
tailRow = ContractEstimate.seachLastCell
Rows(tailRow).Delete
numberCol = 10
letterCol = "J"
tailRow = ContractEstimate.seachLastCell
Call Estimate.cancelMerge("I", tailRow, "J", tailRow, 1)
Call Estimate.ndsTotal(tailRow, letterCol, numberCol)
Call Estimate.setFormat(tailRow + 1, numberCol, tailRow + 2, numberCol)

End Sub
