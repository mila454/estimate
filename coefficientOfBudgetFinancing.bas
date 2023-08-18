Attribute VB_Name = "coefficientOfBudgetFinancing"
Option Explicit

Dim currWB As Workbook
Dim kt As String
Dim kName As String

Sub coefficientOfBudgetFinancing()
'���������� � ����� � ��� ������������ ���������� ��������������
Dim seachRange As Range, seachStr As String
Dim foundCell As Range, firstFoundCell As Range
Dim lastRow As Integer
Dim rangeBorders As Range
Dim sheetName As String
Dim currSheet As Variant
Dim typeEstimate As String
Dim colNumber As Integer
Dim colLetter As String

typeEstimate = "���"

Select Case typeEstimate
    Case "���"
        colNumber = 11
        colLetter = "K"
    Case "��"
        colNumber = 10
        colLetter = "J"
End Select
' ���� ������������ ��������
kName = "������������� ���������� ��������������"
kt = ContractEstimate.simpleInput(kName)
Set currWB = ActiveWorkbook
sheetName = "�����*"
For Each currSheet In Worksheets
    If currSheet.Name Like sheetName Then
        currSheet.Activate
    End If
Next
' ����� ��������� �������� ������
lastRow = ContractEstimate.seachLastCell() + 1

' ������������ �������  � ������ ������
Set seachRange = Range("A1:I" & lastRow)
' ����� ������� ����������
seachStr = "����� �* ���*"
Set foundCell = seachRange.Find(seachStr, LookIn:=xlValues)
Set firstFoundCell = foundCell
If Not foundCell Is Nothing Then
    Call ContractEstimate.TotalWithK(kt, kName, foundCell.Row, foundCell.Row, colLetter, 1, colNumber)
End If
' ����� ��������� ����������
Do
    Set foundCell = seachRange.FindNext(After:=foundCell)
    If foundCell.Address = firstFoundCell.Address Then Exit Do
    Call ContractEstimate.TotalWithK(kt, kName, foundCell.Row, foundCell.Row, colLetter, 1, colNumber)
Loop


'sheetName = "���*"
'For Each currSheet In Worksheets
'    If currSheet.Name Like sheetName Then
'        currSheet.Activate
'    End If
'Next

'If ActiveSheet.Name = "���" Then
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



