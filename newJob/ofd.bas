Attribute VB_Name = "ofd"
Option Explicit
Dim column_name As String
Dim column_list(4) As String
Dim i As Integer
Dim item As Variant
Dim match As Boolean
Dim lastCell As Integer



Sub design_list()
'�������������� �����, ������������ �� ���

column_list(0) = "�������� �����"
column_list(1) = "����/����� �������� �����"
column_list(2) = "�������� ����� �������"
column_list(3) = "����� ������� ���������"
column_list(4) = "����� ������� ������������ (���������)"

Range(Cells(1), Cells(1)).EntireRow.Delete

'�������� ��������, ����� ���� �������������

For i = 1 To 5
    For Each item In column_list
    
        If Cells(1, i).Value <> item Then
            
            match = False
        Else
            match = True
            Exit For
        End If
        
    Next
    
    If match = False Then
        Range(Cells(i), Cells(i)).EntireColumn.Delete
        i = i - 1
    End If
Next

Range("F1:BV1").EntireColumn.Delete

'�������������� ������� ����
lastCell = seachLastCell()

With Range(Cells(2, 2), Cells(lastCell, 2))
    .NumberFormat = "m/d/yyyy"
    
End With

For i = 2 To lastCell
    Cells(i, 2).Value = DateSerial(year(Cells(i, 2)), Month(Cells(i, 2)), Day(Cells(i, 2)))
Next


'�������� ������� �������
Range(Cells(1, 1), Cells((lastCell - 1), 6)).Select
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "�����!R1C1:R129C5", Version:=6).CreatePivotTable TableDestination:= _
        "����1!R3C1", TableName:="������� �������1", DefaultVersion:=6
    Sheets("����1").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("������� �������1").PivotFields("�������� �����")
        .Orientation = xlRowField
        .position = 1
    End With
    ActiveSheet.PivotTables("������� �������1").AddDataField ActiveSheet. _
        PivotTables("������� �������1").PivotFields("�������� ����� �������"), _
        "����� �� ���� �������� ����� �������", xlSum
    ActiveSheet.PivotTables("������� �������1").AddDataField ActiveSheet. _
        PivotTables("������� �������1").PivotFields("����� ������� ���������"), _
        "����� �� ���� ����� ������� ���������", xlSum
    ActiveSheet.PivotTables("������� �������1").AddDataField ActiveSheet. _
        PivotTables("������� �������1").PivotFields( _
        "����� ������� ������������ (���������)"), _
        "����� �� ���� ����� ������� ������������ (���������)", xlSum
    With ActiveSheet.PivotTables("������� �������1").PivotFields( _
        "����/����� �������� �����")
        .Orientation = xlPageField
        .position = 1
    End With
End Sub



Function seachLastCell()
' ����� ��������� �������� ������ � �������� � 1-�� �� 5-�
    Dim c(5) As Integer
    For i = 1 To 5
        c(i) = Cells(Rows.Count, i).End(xlUp).row
    Next
    seachLastCell = WorksheetFunction.Max(c)
End Function


