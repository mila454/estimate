Attribute VB_Name = "calculationOfEquipment"
Option Explicit
Dim lastCell As Integer
Dim seachRange As Range
Dim seachString As String
Dim �quipment As New collection
Dim i As Integer
Dim �quipmentSummary As New collection
Dim j As Integer
Dim stringOfFormula As String


Sub calculationOfEquipment()
'������ ������������ ���������

j = 1

lastCell = seachLastCell()

Set seachRange = Range("A1:C" & lastCell)
seachString = "������������:*"
Set �quipment = Estimate.Seach(seachString, seachRange)
seachString = "������������"
Set �quipmentSummary = Estimate.Seach(seachString, seachRange)
Call quickSort.quickSort(�quipment, 1, �quipment.Count)

For j = 1 To �quipment.Count
    If j = 1 Then
        If IsEmpty(Cells(�quipment(j) + 2, 12)) Then
            stringOfFormula = "L" & �quipment(j) + 3
        Else
            stringOfFormula = "L" & �quipment(j) + 2
        End If
    Else
        If IsEmpty(Cells(�quipment(j) + 2, 12)) Then
            stringOfFormula = stringOfFormula & "+L" & �quipment(j) + 3
        Else
            stringOfFormula = stringOfFormula & "+L" & �quipment(j) + 2
        End If
    End If
Next

Cells(�quipmentSummary(�quipmentSummary.Count - 1), 12).Formula = "=" & stringOfFormula



End Sub

Function seachLastCell()
' ����� ��������� �������� ������ � �������� � 1-�� �� 12-�
    Dim c(12) As Integer
    For i = 1 To 12
        c(i) = Cells(Rows.Count, i).End(xlUp).row
    Next
    seachLastCell = WorksheetFunction.Max(c)
End Function
Function Seach(seachStr, seachRange, token) As collection
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
    If token = "row" Then
        Seach.Add foundCell.row
    Else
        Seach.Add foundCell.Offset(0, token).Value
    End If
    
Loop While foundCell.Address <> firstFoundCell.Address

End Function

