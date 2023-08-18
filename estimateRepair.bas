Attribute VB_Name = "estimateRepair"
Option Explicit


Public currWB As Workbook
Dim smetaName As String, smetaName2 As String, mon As String, year As String
Dim totalEstimate As New Collection
Dim numLocEst As New Collection
Public typeEstimate As String
Public number As Integer



Sub estimateRepair()
'������� �������
' ������������ ����� ��, ���� ����� �������� �� ���������
Dim seachRange As Range, seachStr As String
Dim lastRow As Integer, firstRow As Integer
Dim i As Long
Dim tempNumLocEst As Integer
Dim signer As String
Dim position As String
Dim title2 As New Collection

Set currWB = ActiveWorkbook
typeEstimate = InputBox("������� ��� �����: ��� ��� ��", , "��")
number = InputBox("������� ���������� ����:", , "1")
signer = "�.�.�������"
position = "�������� ��� �.������ " & Chr(34) & "�������� ����������" & Chr(34)

mon = "�������"
year = Estimate.seachMonthYear("year", currWB, typeEstimate)
currWB.Sheets(1).Activate
lastRow = ContractEstimate.seachLastCell()

Set seachRange = Range("A1:I" & lastRow)
 
' ����������� ������ ������ ����� ��������� ����
Set seachRange = Range("A1:K" & lastRow)
seachStr = "����� �� ��������� �����*"
Set numLocEst = Estimate.Seach(seachStr, seachRange)
Call quickSort.quickSort(numLocEst, 1, numLocEst.Count)

' ������������ �������
lastRow = numLocEst(1)
Set seachRange = Range("A1:I" & lastRow)

' ���������� �������� �����
If Sheets("Source").Range("F12") <> "" Then
    smetaName = Sheets("Source").Range("G12")
    Sheets("Source").Range("F20").formula = "=G12"
End If

' ����� � �������� �����
Call Estimate.header(smetaName, signer, position)

Worksheets("Source").Cells(1, 10).Clear
Worksheets("SourceObSm").Cells(1, 10).Clear

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
Dim Shift As Integer

Sheets.Add Before:=Sheets(1), Type:="C:\���������\������\��������\�������\���� ������.xltx"

Sheets("����").Activate
Sheets("����").Name = "���"

Cells(8, 1) = smetaName
Call heightAdjustment.heightAdjustment(Range("A8:E8"))

Cells(15, 2) = "������������ ������� ��������� ������������� � ������� ������ ��� �� " & mon & " " & year & " �."
If typeEstimate = "���" Then
    strSheetName = "����� �� ���-2001(� ���.67"
    letterCol = "K"
    
ElseIf typeEstimate = "��" Then
    strSheetName = "����� ��-2012 �� ��. 1-5"
    letterCol = "I"
End If

If number = 1 Then
    Shift = 0
Else
    Shift = 1
    Cells(20, 1).EntireRow.Insert CopyOrigin:=xlFormatFromRightOrBelow
    Cells(20, 1).Value = "�����������-��������� ������ (��������� ����� � 1)"
    Cells(20, 2).formula = "='" & strSheetName & "'!" & letterCol & numLocEst(1)
    Cells(20 + Shift, 1).Value = "�����������-��������� ������ (��������� ����� � 2)"
    Cells(22, 2).formula = "=B20-B21"
    Range("C21:E21").Copy
    Range("C20").PasteSpecial
    Application.CutCopyMode = False
End If

Cells(20 + Shift, 2).formula = "='" & strSheetName & "'!" & letterCol & numLocEst(1 + Shift)


End Sub





