Attribute VB_Name = "calculationOfEquipment2"
Option Explicit
Dim lastCell As Integer
Dim seachRange As Range
Dim seachString As String
Dim �quipment As New collection
Dim i As Integer
Dim �quipmentSummary As New collection
Dim j As Integer
Dim stringOfFormula As String


Sub calculationOfEquipment2()
'������ ������������ �������������
'�������

j = 1
i = 1


lastCell = seachLastCell()
Set seachRange = Range("A1:N" & lastCell)

For i = 1 To lastCell
    If Cells(i, 1).Value Like "*�" Then
       �quipment.Add i
    End If
Next

seachString = "������������"
Set �quipmentSummary = Seach(seachString, seachRange)
Call quickSort.quickSort(�quipment, 1, �quipment.Count)

For j = 1 To �quipment.Count
    If j = 1 Then
        stringOfFormula = "N" & �quipment(j)
    Else
        stringOfFormula = stringOfFormula & "+N" & �quipment(j)
    End If
Next

Cells(�quipmentSummary(�quipmentSummary.Count - 1), 14).Formula = "=" & stringOfFormula

Set �quipment = New collection
Set �quipmentSummary = New collection


End Sub

Function seachLastCell()
' ����� ��������� �������� ������ � �������� � 1-�� �� 12-�
    Dim c(12) As Integer
    For i = 1 To 12
        c(i) = Cells(Rows.Count, i).End(xlUp).row
    Next
    seachLastCell = WorksheetFunction.Max(c)
End Function
Function Seach(seachStr, seachRange) As collection
'����� �� ������ � ���������� ������ ���� � ���������
Dim foundCell As Range
Dim firstFoundCell As Range

Set Seach = New collection

Set foundCell = seachRange.Find(seachStr, LookIn:=xlValues, MatchCase:=True)
Set firstFoundCell = foundCell

If firstFoundCell Is Nothing Then
    MsgBox (seachStr & " �� �������")
    Exit Function
End If

Do
    Set foundCell = seachRange.FindNext(After:=foundCell)
    Seach.Add foundCell.row
Loop While foundCell.Address <> firstFoundCell.Address

End Function

