Attribute VB_Name = "estimateSN"
Option Explicit


Public currWB As Workbook
Dim smetaName As String, smetaName2 As String, mon As String, year As String
Dim plantingTrees As New Collection, restTrees As New Collection, treesCare1 As New Collection, treesCare2 As New Collection, treesCare3 As New Collection
Dim plantingTrees2 As New Collection, restTrees2 As New Collection
Dim totalEstimate As New Collection
Public kindOfWorks As New Collection
Dim numLocEst As New Collection
Dim god1 As String, god2 As String, god3 As String
Public typeEstimate As String


Sub estimateSN()
' ������������ ����� ��, ���� ����� �������� �� ���������
Dim seachRange As Range, seachStr As String
Dim lastRow As Integer, firstRow As Integer
Dim i As Long
Dim tempNumLocEst As Integer
Set plantingTrees = New Collection
Set restTrees = New Collection
Set treesCare1 = New Collection
Set treesCare2 = New Collection
Set treesCare3 = New Collection
Set totalEstimate = New Collection
Set kindOfWorks = New Collection
Set numLocEst = New Collection
Dim title2 As New Collection

god1 = "2024"
god2 = "2025"
god3 = "2026"

Set currWB = ActiveWorkbook
typeEstimate = "��"
mon = "�������"
year = Estimate.seachMonthYear("year", currWB, typeEstimate)
currWB.Sheets(1).Activate
lastRow = ContractEstimate.seachLastCell()

' �������� ����� ����� ����� �� �����
Set seachRange = Range("A1:I" & lastRow)
Range("A" & (lastRow - 1) & ":A" & (lastRow + 1)).EntireRow.Delete
 
' ����������� ������ ������ ����� ��������� ����
Set seachRange = Range("A1:K" & lastRow)
seachStr = "����� �� ��������� �����:*"
Set numLocEst = Estimate.Seach(seachStr, seachRange)

' ������������ �������
lastRow = numLocEst(numLocEst.Count)
Set seachRange = Range("A1:I" & lastRow)

' ���������� �������� �����
If Sheets("Source").Range("F20") <> "" Then
    smetaName = Sheets("Source").Range("G20")
End If

' ����� � �������� �����
Call Estimate.header(smetaName, "����������� �.�.", "����������� ��������� ��� �.������ " & Chr(34) & "�������� ����������" & Chr(34))

Worksheets("Source").Cells(1, 10).Clear
Worksheets("SourceObSm").Cells(1, 10).Clear

Range("A" & lastRow) = "����� �� ��������� ����� �1: " & smetaName
Call Estimate.cancelMerge("I", lastRow, "J", lastRow, 1)
Cells(lastRow, 10).Value = Round(Cells(lastRow, 10).Value * 100, 2) / 100
firstRow = lastRow
' ������ ������ 1-� ��������� �����
Call Estimate.estimateTail(seachRange, god1, god2, god3)
' ����������� ������ 1-� ��������� �����
Cells(lastRow, 10).ColumnWidth = 15
For i = 1 To 21
    Rows(lastRow + 1).Insert
    Rows(lastRow + 1).ClearFormats
Next

Call Estimate.displayTail(lastRow, "J", "I", 10, god1, god2, god3)

'�������� ������ ����� ����� 2 ��������� ������
seachStr = "��������� ����� �*"
Set seachRange = Range("A" & lastRow & ":K" & numLocEst(1))
Set title2 = Estimate.Seach(seachStr, seachRange)


Range("A" & lastRow + 4 & ":A" & title2(1) - 2).EntireRow.Delete
Cells(lastRow + 5, 1).Value = "��������� ����� � 2 "
With Range("A" & lastRow + 10)
    .Value = smetaName
    .RowHeight = 35
End With
smetaName2 = Cells(lastRow + 8, 1).Value

 ' ����������� ������ ������ ����� ��������� ����
lastRow = Sheets(1).Range("A1").SpecialCells(xlCellTypeLastCell).Row
Set seachRange = Range("A1:K" & lastRow)

seachStr = "����� �� ��������� �����*"
Set numLocEst = New Collection
Set numLocEst = Estimate.Seach(seachStr, seachRange)
Call quickSort.quickSort(numLocEst, 1, numLocEst.Count)

seachStr = "����� �� *�����*"
Set totalEstimate = Estimate.Seach(seachStr, seachRange)
Call quickSort.quickSort(totalEstimate, 1, totalEstimate.Count)

lastRow = totalEstimate(totalEstimate.Count)
Range("A" & numLocEst(numLocEst.Count)) = "����� �� ��������� ����� �2: " & smetaName2
Call Estimate.cancelMerge("I", numLocEst(numLocEst.Count), "J", numLocEst(numLocEst.Count), 1)
Call NDSIncluding(numLocEst(numLocEst.Count) + 1, 10, "J")
Call Estimate.setFormat(numLocEst(numLocEst.Count) + 1, 1, numLocEst(numLocEst.Count) + 1, 10)
Rows(lastRow).EntireRow.Delete
Application.DisplayAlerts = False
With Range("A" & lastRow + 1 & ":F" & lastRow + 1)
    .Value = "����� �� ��������� ������ �1,2: " & smetaName
    .RowHeight = 35
    .WrapText = True
    .Merge
    .VerticalAlignment = xlCenter
End With


Call Estimate.cancelMerge("I", lastRow + 1, "J", lastRow + 1, 1)
tempNumLocEst = numLocEst(1) + 2
numLocEst.Add numLocEst(1) + 2, , Before:=1
numLocEst.Remove (2)
Cells(lastRow + 1, 10).formula = Estimate.formulaTotal(numLocEst, "J")
' ������ ������ 2-� ��������� �����
Set seachRange = Range("A" & firstRow & ":K" & lastRow)
seachStr = "����� �� �������: *��� �������*"
Set plantingTrees2 = Estimate.Seach(seachStr, seachRange)
Call Estimate.cancelMerge("I", plantingTrees2(plantingTrees2.Count), "J", plantingTrees2(plantingTrees2.Count), 1)
Call quickSort.quickSort(plantingTrees2, 1, plantingTrees2.Count)
seachStr = "����� �� �������: *��� ��������������*"
Set restTrees2 = Estimate.Seach(seachStr, seachRange)
If restTrees2.Count <> 0 Then
    Call Estimate.cancelMerge("I", restTrees2(restTrees2.Count), "J", restTrees2(restTrees2.Count), 1)
    Call quickSort.quickSort(restTrees2, 1, restTrees2.Count)
End If
' ����������� ������ 2-� ��������� �����
Call NDSIncluding(lastRow + 2, 10, "J")
Cells(lastRow + 4, 1) = "� ��� �����:"
Cells(lastRow + 6, 1) = "������� �������� (" & god1 & " ���)"
Cells(lastRow + 6, 10).formula = "=J" & Estimate.kindOfWorks(1) + 2 & "+J" & plantingTrees2(1)
Call NDSIncluding(lastRow + 7, 10, "J")
lastRow = lastRow + 7
If restTrees2.Count <> 0 Then
    Cells(lastRow + 2, 1) = "����������������� � ������� ������ (" & god1 & " ���)"
    Cells(lastRow + 2, 10).formula = "=J" & Estimate.kindOfWorks(2) + 2 & "+J" & restTrees2(1)
    Call NDSIncluding(lastRow + 3, 10, "J")
    lastRow = lastRow + 4
Else
    Cells(lastRow + 2, 1) = "������� ������ (" & god1 & " ���)"
    Cells(lastRow + 2, 10).formula = "=J" & Estimate.kindOfWorks(2) + 2
    Call NDSIncluding(lastRow + 3, 10, "J")
    lastRow = lastRow + 4
End If
Cells(lastRow + 1, 1) = "������� ������ (" & god2 & " ���)"
Cells(lastRow + 1, 10).formula = "=J" & Estimate.kindOfWorks(3) + 2
Call NDSIncluding(lastRow + 2, 10, "J")
lastRow = lastRow + 3
Cells(lastRow + 1, 1) = "������� ������ (" & god3 & " ���)"
Cells(lastRow + 1, 10).formula = "=J" & Estimate.kindOfWorks(4) + 2
Call NDSIncluding(lastRow + 2, 10, "J")

Range("A" & lastRow - 13 & ":C" & (lastRow + 2)).Font.Bold = True
Call Estimate.setFormat(lastRow - 13, 10, lastRow + 2, 10)

Call fillNMCK


End Sub

Sub NDSIncluding(numberRow, numberCol, letterCol)

Cells(numberRow, 1).Value = "� ��� ����� ��� 20%"
Cells(numberRow, numberCol).formula = "=round(" & letterCol & numberRow - 1 & "*20/120, 2)"

End Sub

Sub fillNMCK()
'������� ����� ���� � ���������� ���
Dim strSheetName As String
Dim letterCol As String

Sheets.Add Before:=Sheets(1), Type:="C:\���������\������\��������\�������\���� ��.xltx"

Sheets("����").Activate
Sheets("����").Name = "���"

Cells(9, 1) = smetaName
Cells(15, 2) = "������������ ������� ��������� ������������� � ������� ������ ��� �� " & mon & " " & year & " �."
If typeEstimate = "���" Then
    strSheetName = "����� �� ���-2001(� ���.67"
    letterCol = "K"
ElseIf typeEstimate = "��" Then
    strSheetName = "����� ��-2012 �� ��. 1-5"
    letterCol = "J"
End If

Cells(18, 2).formula = "='" & strSheetName & "'!" & letterCol & numLocEst(1)
Cells(18, 4).formula = "='" & strSheetName & "'!" & letterCol & Estimate.kindOfWorks(1) + 2
Cells(18, 5) = "='" & strSheetName & "'!" & letterCol & Estimate.kindOfWorks(2) + 2
Cells(18, 6) = "='" & strSheetName & "'!" & letterCol & Estimate.kindOfWorks(3) + 2
Cells(18, 7) = "='" & strSheetName & "'!" & letterCol & Estimate.kindOfWorks(4) + 2
Cells(19, 4) = "='" & strSheetName & "'!" & letterCol & plantingTrees2(1)
If restTrees2.Count <> 0 Then
    Cells(19, 5) = "='" & strSheetName & "'!" & letterCol & restTrees2(1)
End If

End Sub




