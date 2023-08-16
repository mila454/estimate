Attribute VB_Name = "Estimate"

Option Explicit
Public currWB As Workbook
Dim smetaName As String, mon As String, year As String
Dim plantingTrees As New Collection, restTrees As New Collection, treesCare1 As New Collection, treesCare2 As New Collection, treesCare3 As New Collection
Dim totalEstimate As New Collection
Public kindOfWorks As New Collection
Dim god1 As String, god2 As String, god3 As String
Public typeEstimate As String
Dim sheetName As String


Sub Estimate()
Dim seachRange As Range, seachStr As String
Dim lastRow As Integer
Dim i As Integer
Dim currSheet As Variant
Dim signer As String
Dim position As String

' ������������ �����, ���� � ��������� ����� �������� �� ���������

Set currWB = ActiveWorkbook
typeEstimate = InputBox("������� ��� �����: ��� ��� ��", , "���")
signer = "�.�. �����������"
position = "����������� ��������� ��� �.������ " & Chr(34) & "�������� ����������" & Chr(34)

god1 = InputBox("������� ������ ��� �����", , "2023")
god2 = InputBox("������� ������ ��� �����", , "2024")
god3 = InputBox("������� ������ ��� �����", , "2025")

sheetName = "�����*"
For Each currSheet In Worksheets
    If currSheet.Name Like sheetName Then
        currSheet.Activate
        sheetName = ActiveSheet.Name
    End If
Next

mon = seachMonthYear("month", currWB, typeEstimate)
year = seachMonthYear("year", currWB, typeEstimate)
currWB.Sheets(1).Activate

' �������� � ������ ����������� ����� ����� ����� �� �����
lastRow = ContractEstimate.seachLastCell()
Set seachRange = Range(Cells(1, 1), Cells(lastRow, 9))

seachStr = "����� ��*�����*"
Set totalEstimate = Seach(seachStr, seachRange)

Range(Cells(totalEstimate(totalEstimate.Count) + 1, 1), Cells(lastRow + 1, 1)).EntireRow.Delete

' ����� ������������ ����� �������� ����� ��������� �������� ������
lastRow = ContractEstimate.seachLastCell()
Set seachRange = Range(Cells(1, 1), Cells(lastRow, 9))

' ��������� � �������� ������� �����
Range("A" & (lastRow) & ":A" & (lastRow + 2)).EntireRow.Hidden = False
Range("A" & (lastRow + 1) & ":A" & (lastRow + 3)).Delete

' ���������� �������� �����
If Sheets("Source").Range("F20") <> "" Then
    smetaName = Sheets("Source").Range("G20")
End If

Range("A" & lastRow) = "����� �� ��������� ����� �1: " & smetaName
Call cancelMerge("J", lastRow, "K", lastRow, 0)
Cells(lastRow, 11).formula = "=SUM(P36:P" & lastRow & ")"
Call cancelMerge("H", lastRow, "I", lastRow, 0)
Cells(lastRow, 9).formula = "=SUM(O36:O" & lastRow & ")"
Call setFormat(lastRow, 9, lastRow, 9)

' ����� � �������� �����
Call header(smetaName, signer, position)

Range("G36:K36").EntireRow.Hidden = True
Worksheets("Source").Cells(1, 10).Clear
Worksheets("SourceObSm").Cells(1, 10).Clear

' ������ ������
Call estimateTail(seachRange, god1, god2, god3)


' ����������� ������
Range("G" & (lastRow + 1) & ":G" & (lastRow + 2)).Clear
Range("I" & lastRow & ", K" & lastRow).ColumnWidth = 15
Cells(lastRow + 1, 9).formula = "=round(I" & (lastRow) & "*0.2,2)"
Cells(lastRow + 2, 9).formula = "=I" & lastRow & "+I" & (lastRow + 1)
Call setFormat(lastRow + 1, 9, lastRow + 2, 9)

Call displayTail(lastRow, "K", "J", 11, god1, god2, god3)



' �������� ������
'Cells(rowTotal, 13).Formula = "=K" & t(0) & "+" & "K" & t(1) & "+" & "K" & t(2) & "+" & "K" & t(3)
'Call checkTotal(rowTotal, t, 7, 1)
'Call checkTotal(rowTotal, t, 7, 0)

' ������� ������ �� ���������� � ���������� ��
Call fillSummaryEstimate("����")

Call fillSummaryEstimate("���")

' ������� ����� � ���� � ���������� ���
Call fillNMCK

' ���������� ���������
Call listOfWorks

End Sub

Function Seach(seachStr, seachRange) As Collection
'����� �� ������ � ���������� ������ ���� � ���������
Dim foundCell As Range
Dim firstFoundCell As Range

Set Seach = New Collection

Set foundCell = seachRange.Find(seachStr, LookIn:=xlValues, MatchCase:=True)
Set firstFoundCell = foundCell

If firstFoundCell Is Nothing Then
    MsgBox (seachStr & " �� �������")
    Exit Function
End If

Do
    Set foundCell = seachRange.FindNext(After:=foundCell)
    Seach.Add foundCell.Row
Loop While foundCell.Address <> firstFoundCell.Address

End Function

Function formulaTotal(coll, letterColl) As String
'������������ ������ � �������� ����� �� ������� �� ���������
Dim i As Integer
formulaTotal = "="
For i = 1 To coll.Count
    If i = coll.Count Then
        formulaTotal = formulaTotal & letterColl & coll(i)
    Else
        formulaTotal = formulaTotal & letterColl & coll(i) & "+"
    End If
Next i

End Function

Sub ndsTotal(r, indCol, numCol)
'������ � ����� ��� � ����� � ���

Cells(r + 1, 1) = "��� 20%"
Cells(r + 1, numCol).formula = "=round(" & indCol & (r) & "*0.2,2)"

Cells(r + 2, 1) = "����� � ��� 20%"
Cells(r + 2, numCol).formula = "=" & indCol & r & "+" & indCol & (r + 1)

Range("A" & r & ":" & indCol & (r + 2)).Font.Bold = True
Call setFormat(r, numCol, r + 2, numCol)

End Sub

Sub cancelMerge(col1, row1, col2, row2, transferStatus)
' ������ ����������� ����� � ������� ������
Range(col1 & row1 & ":" & col2 & row2).UnMerge

If transferStatus = 1 Then
    Range(col1 & row1).Copy
    Range(col2 & row2).PasteSpecial (xlPasteValuesAndNumberFormats)
        Range(col1 & row1).Clear
Else
    Range(col1 & row1 & ":" & col2 & row2).Clear
End If

End Sub

Sub fillSummaryEstimate(nameSh)
'������� ������ �������� �������� ������� � ���������� ��
Sheets.Add Before:=Sheets(1), Type:="C:\���������\������\��������\�������\���.xltx"

nameSh = "��� " & nameSh

Sheets(1).Name = nameSh

currWB.Sheets(nameSh).Activate

Range("A15") = smetaName

If nameSh = "��� ���" Then
    Range("F34").Value = currWB.Sheets(sheetName).Range("J21").Value
    Range("A17") = "� ����� �� " & mon & " " & year & " �."
ElseIf nameSh = "��� ����" Then
    Range("F34").Value = currWB.Sheets(sheetName).Range("I21").Value
    Range("A17") = "� ����� �� 01.01.2000 �."
End If

End Sub


Function seachMonthYear(monthOrYear, currWB, typeEstimate)
' ���������� ������ ��� ���� �������������
Dim sRange As Range, fCell As Range, fFCell As Range
Dim seachSH As Worksheet
Dim sStr As String
Dim LR As Integer
Dim numCol As Integer
Dim coefCell As New Collection

Set seachSH = currWB.Worksheets("Source")
currWB.Sheets("Source").Activate

LR = Cells(seachSH.Rows.Count, 1).End(xlUp).Row
Set sRange = seachSH.Range("B1:C" & LR)

If typeEstimate = "���" Then
    sStr = "������������ � ���-2001 ���"
    numCol = 4
ElseIf typeEstimate = "��" Then
    sStr = "������� ���*"
    numCol = 3
End If

Set coefCell = Seach(sStr, sRange)

If monthOrYear = "month" Then
    Select Case Cells(coefCell(1), 5).Value
        Case 1
            seachMonthYear = "������"
        Case 2
            seachMonthYear = "�������"
        Case 3
            seachMonthYear = "����"
        Case 4
            seachMonthYear = "������"
        Case 5
            seachMonthYear = "���"
        Case 6
            seachMonthYear = "����"
        Case 7
            seachMonthYear = "����"
        Case 8
            seachMonthYear = "������"
        Case 9
            seachMonthYear = "��������"
        Case 10
            seachMonthYear = "�������"
        Case 11
            seachMonthYear = "������"
        Case 12
            seachMonthYear = "�������"
    End Select
ElseIf monthOrYear = "year" Then
    seachMonthYear = Cells(coefCell(1), numCol).Value
End If

End Function

Sub fillNMCK()
'������� ����� ���� � ���������� ���
Dim strSheetName As String
Dim letterCol As String

Sheets.Add Before:=Sheets(1), Type:="C:\���������\������\��������\�������\����.xltx"

Sheets("����").Activate
Sheets("����").Name = "���"

Cells(9, 1) = smetaName
Cells(15, 2) = "������������ ������� ��������� ������������� � ������� ������ ��� �� " & mon & " " & year & " �."
letterCol = "K"
strSheetName = sheetName
'If typeEstimate = "���" Then
'    strSheetName = "����� �� ���-2001(� ���.67"
'    letterCol = "K"
'ElseIf estimateSN.typeEstimate = "��" Then
'    strSheetName = "����� ��-2012 �� ��. 1-5"
'    letterCol = "J"
'End If

Cells(18, 2).formula = "='" & strSheetName & "'!" & letterCol & totalEstimate(totalEstimate.Count)
Cells(18, 4).formula = "='" & strSheetName & "'!" & letterCol & kindOfWorks(1)
Cells(18, 5) = "='" & strSheetName & "'!" & letterCol & kindOfWorks(2)
Cells(18, 6) = "='" & strSheetName & "'!" & letterCol & kindOfWorks(3)
Cells(18, 7) = "='" & strSheetName & "'!" & letterCol & kindOfWorks(4)

End Sub

Sub listOfWorks()
'������������� ��������� �����
Sheets("��������� ���������").Name = "��������� �����"
Sheets("��������� �����").Activate

Cells(11, 1) = "��������� �����"
Cells(12, 1) = smetaName
Range(Cells(14, 2), Cells(16, 2)).Clear
Rows(19).Delete

End Sub

Sub setFormat(row1Format, col1Format, row2Format, col2Format)
'�������������� ���������
With Range(Cells(row1Format, col1Format), Cells(row2Format, col2Format))
    .Font.Bold = True
    .NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
End With

End Sub

Sub header(smetaName, signer, position)
'������������ �����
With Range(Cells(3, 2), Cells(6, 11))
    .UnMerge
    .Clear
    
End With

Cells(3, 2) = Chr(34) & "���������" & Chr(34)
Cells(5, 2) = "��������:"
Cells(6, 2) = position & "_________________________ " & signer

With Range(Cells(6, 2), Cells(6, 5))
   .MergeCells = True
   .WrapText = True
End With

Rows("6:6").RowHeight = 35
Range("A10, A15") = smetaName
'Range("10:10, 15:15").RowHeight = 35
Call heightAdjustment.heightAdjustment(Range("A10:K10"))
Call heightAdjustment.heightAdjustment(Range("A15:K15"))

Cells(13, 1) = "��������� ����� � 1"
With Range(Cells(3, 2), Cells(13, 2))
    .Font.Name = "Arial"
    .Font.Size = 13
End With
Range(Cells(7, 7), Cells(7, 11)).Clear

End Sub

Public Sub estimateTail(seachRange, god1, god2, god3)
'������������ ��������� ������� ����� ������ �� ����� �����
Dim seachStr As String
Dim i As Integer

seachStr = "����� �� �������: �������*"
Set plantingTrees = Seach(seachStr, seachRange)

seachStr = "����� �� �������: �������������� ������*"
Set restTrees = Seach(seachStr, seachRange)

seachStr = "����� �� �������: ����*" & god1
Set treesCare1 = Seach(seachStr, seachRange)
For i = 1 To treesCare1.Count
    restTrees.Add treesCare1(i)
Next i
    
seachStr = "����� �� �������: ����*" & god2
Set treesCare2 = Seach(seachStr, seachRange)

seachStr = "����� �� �������: ����*" & god3
Set treesCare3 = Seach(seachStr, seachRange)
Call quickSort.quickSort(plantingTrees, 1, plantingTrees.Count)
Call quickSort.quickSort(restTrees, 1, restTrees.Count)
Call quickSort.quickSort(treesCare2, 1, treesCare2.Count)
Call quickSort.quickSort(treesCare3, 1, treesCare3.Count)

End Sub

Public Sub displayTail(lastRow, letterCol, letterCol2, numberCol, god1, god2, god3)
Set kindOfWorks = New Collection
Call ndsTotal(lastRow, letterCol, numberCol)
Call setFormat(lastRow + 1, numberCol, lastRow + 2, numberCol)
Cells(lastRow + 4, 1) = "� ��� �����:"
Cells(lastRow + 6, 1) = "������� �������� (" & god1 & " ���)"
Cells(lastRow + 6, numberCol) = formulaTotal(plantingTrees, letterCol2)
kindOfWorks.Add (lastRow + 6)
lastRow = lastRow + 6
Call ndsTotal(lastRow, letterCol, numberCol)
Cells(lastRow + 4, 1) = "����������������� � ������� ������ (" & god1 & " ���)"
Cells(lastRow + 4, numberCol).Value = formulaTotal(restTrees, letterCol2)
kindOfWorks.Add (lastRow + 4)
lastRow = lastRow + 4
Call ndsTotal(lastRow, letterCol, numberCol)
Cells(lastRow + 4, 1) = "������� ������ (" & god2 & " ���)"
Cells(lastRow + 4, numberCol).Value = formulaTotal(treesCare2, letterCol2)
kindOfWorks.Add (lastRow + 4)
lastRow = lastRow + 4
Call ndsTotal(lastRow, letterCol, numberCol)
Cells(lastRow + 4, 1) = "������� ������ (" & god3 & " ���)"
Cells(lastRow + 4, numberCol).Value = formulaTotal(treesCare3, letterCol2)
kindOfWorks.Add (lastRow + 4)
lastRow = lastRow + 4
Call ndsTotal(lastRow, letterCol, numberCol)
Range("A" & lastRow - 17 & ":C" & (lastRow + 2)).Font.Bold = True

End Sub








