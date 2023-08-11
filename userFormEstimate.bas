Attribute VB_Name = "userFormEstimate"
Option Explicit

Dim lastRow As Integer
Dim seachRange As Range
Dim seachStr As String
Dim letterCol As String
Dim numberCol As Integer
Dim currYear As Integer
Dim answer As Variant '����� �� ������ �� ��� � ��� �����
Dim signer As String '��� �������������
Dim position As String '��������� �������������
Dim typeEstimate As String '��� �����: ��� ��� ��
Dim totalEstimate As New Collection '����� ������ ����� �� �����
Dim rowForCoefficient As New Collection '����� ������ ��� ������ ������������
Dim namePosition As New Dictionary '����:����� ��������� ����� � ��������: ������������ ��������� �����
Dim nameLocation As Variant '����� ���� ������������ ������������ �����
Dim smetaName As Variant '������������ �����
Dim numberEstimates As Integer '���������� ����
Dim coefficientName As String '������������ �����������
Dim coefficient As Variant '�������� ������������

Sub userFormEstimate()

prepareEstimate.Show

End Sub

Sub nds()
Dim i As Variant
Dim item As Variant

Call activateSheet("����� *")
lastRow = seachLastCell() + 1
Set seachRange = Range(Cells(1, 1), Cells(lastRow, 11))
Call determinationEstimateType
Call header
Call clearTail

If totalEstimate.Count > 1 Then
    For i = 1 To totalEstimate.Count - 1
        Range("A" & totalEstimate(i) + 1 & ":A" & totalEstimate(i) + 3).EntireRow.Insert
        Range("A" & totalEstimate(i) & ":H" & totalEstimate(i)).Value = "����� �� ��������� ����� �" & i & ": " & smetaName(i - 1)
        Call heightAdjustment(Range("A" & totalEstimate(i) & ":H" & totalEstimate(i)))
        If answer = 6 And Cells(totalEstimate(i), 1).Value Like "*���������*�����������*���������*" Then
            Call NDSIncluding(totalEstimate(i) + 1)
        Else
            Call ndsTotal(totalEstimate(i), numberCol, letterCol)
        End If
        totalEstimate.Add totalEstimate(i + 1) + i * 3, , , i
        totalEstimate.Remove (i + 2)
    Next
End If

Range("A" & totalEstimate(totalEstimate.Count) + 1 & ":A" & totalEstimate(totalEstimate.Count) + 3).EntireRow.Insert
Range("A" & totalEstimate(totalEstimate.Count) & ":H" & totalEstimate(totalEstimate.Count)).Value = "����� �� �����:" & smetaName(0)
If answer = 6 Then
    Call NDSIncluding(totalEstimate(totalEstimate.Count) + 1)
Else
    Call ndsTotal(totalEstimate(totalEstimate.Count), numberCol, letterCol)
End If
Call heightAdjustment(Range("A" & totalEstimate(totalEstimate.Count) & ":H" & totalEstimate(totalEstimate.Count)))
For Each item In totalEstimate
    Call cancelMerge(item)
Next

If typeEstimate = "���" Then
    Call ndsTotal(totalEstimate(1), 9, "I")
    Call setFormat(totalEstimate(1), 9, totalEstimate(1) + 2, 9)
End If

Call finishOrGoToMainMenu

End Sub

Function seachLastCell()
' ����� ��������� �������� ������ � �������� � 1-�� �� 11-�
    Dim c(11) As Integer
    Dim i As Variant
    
    For i = 1 To 11
        c(i) = Cells(Rows.Count, i).End(xlUp).Row
    Next
    seachLastCell = WorksheetFunction.Max(c)
End Function

Sub activateSheet(sheetName)
' ������������� �����
Dim currSheet As Variant

For Each currSheet In Worksheets
    If currSheet.Name Like sheetName Then
        currSheet.Activate
        sheetName = ActiveSheet.Name
    End If
Next

End Sub

Function Seach(seachStr, seachRange) As Collection
'����� �� ������ � ���������� ������ ���� � ���������
Dim foundCell As Range
Dim firstFoundCell As Range

Set Seach = New Collection

Set foundCell = seachRange.Find(seachStr, LookIn:=xlValues)
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

Sub quickSort(coll As Collection, first As Long, last As Long)
'������� ���������� ��������� ���������
Dim centreVal As Variant, temp As Variant
Dim low As Long
Dim high As Long

If last = 0 Then Exit Sub

low = first
high = last
centreVal = coll((first + last) \ 2)

Do While low <= high
    Do While coll(low) < centreVal And low < last
        low = low + 1
    Loop
    Do While centreVal < coll(high) And high > first
        high = high - 1
    Loop
    If low <= high Then
    ' �������� ��������
        temp = coll(low)
        coll.Add coll(high), After:=low
        coll.Remove low
        coll.Add temp, Before:=high
        coll.Remove high + 1
        ' ������� � ��������� ��������
        low = low + 1
        high = high - 1
    End If
    Loop
    If first < high Then quickSort coll, first, high
    If low < last Then quickSort coll, low, last
End Sub

Sub clearTail()
'�������, �������� ����������� ����� ����� �� ��������� �����... � ��������-��������
Dim i As Variant

Call activateSheet("����� *")
Set seachRange = Range(Cells(1, 1), Cells(lastRow, 9))
seachStr = "����� ��*�����*"
Set totalEstimate = Seach(seachStr, seachRange)
Call quickSort(totalEstimate, 1, totalEstimate.Count)

For i = 1 To totalEstimate.Count
    If i = totalEstimate.Count Then
        If numberEstimates = 1 Then
            Range("A" & totalEstimate(i) & ":A" & lastRow).EntireRow.Delete
            totalEstimate.Remove (totalEstimate.Count)
        Else
            Range("A" & totalEstimate(i) + 1 & ":A" & lastRow).EntireRow.Hidden = False
            Range("A" & totalEstimate(i) + 1 & ":A" & lastRow).EntireRow.Delete
        End If
    Else
        If numberEstimates > 1 Then
            If i < numberEstimates Then
                Range("A" & totalEstimate(i) + 1 & ":A" & nameLocation(i)).EntireRow.Hidden = False
                Range("A" & totalEstimate(i) + 1 & ":A" & nameLocation(i)).EntireRow.Delete
                totalEstimate.Add totalEstimate(i + 1) - ((nameLocation(i)) - (totalEstimate(i))), , , i
                totalEstimate.Add totalEstimate(i + 3) - ((nameLocation(i)) - (totalEstimate(i))), , , i + 1
                totalEstimate.Remove (totalEstimate.Count)
               
            Else
                Range("A" & totalEstimate(i) + 1 & ":A" & totalEstimate(i + 1)).EntireRow.Hidden = False
                Range("A" & totalEstimate(i) + 1 & ":A" & totalEstimate(i + 1) - 1).EntireRow.Delete
                totalEstimate.Add totalEstimate(i + 1) - ((totalEstimate(i + 1) - 1) - (totalEstimate(i))), , , i
                totalEstimate.Add totalEstimate(i + 3) - ((totalEstimate(i + 1) - 1) - (totalEstimate(i))), , , i + 1
                totalEstimate.Remove (totalEstimate.Count)
                totalEstimate.Remove (totalEstimate.Count)
                totalEstimate.Remove (totalEstimate.Count)
            End If
        Else
            Range("A" & totalEstimate(i) + 1 & ":A" & totalEstimate(i + 1)).EntireRow.Hidden = False
            Range("A" & totalEstimate(i) + 1 & ":A" & totalEstimate(i + 1) - 1).EntireRow.Delete
            totalEstimate.Add totalEstimate(i + 1) - ((totalEstimate(i + 1) - 1) - (totalEstimate(i))), , , i
            totalEstimate.Remove (totalEstimate.Count)
        End If
    End If
Next

End Sub

Sub ndsTotal(r, numberCol, letterCol)
'������ � ����� ��� � ����� � ���

Cells(r + 1, 1).Value = "��� 20%"
Cells(r + 1, numberCol).formula = "=round(" & letterCol & r & "*0.2,2)"
Cells(r + 2, 1) = "����� � ��� 20%"
Cells(r + 2, numberCol).formula = "=" & letterCol & r & "+" & letterCol & (r + 1)
Range("A" & r + 1 & ":A" & r + 2).WrapText = False
Call setFormat(r, numberCol, r + 2, numberCol)

End Sub

Sub setFormat(row1Format, col1Format, row2Format, col2Format)
'�������������� ���������
With Range(Cells(row1Format, col1Format), Cells(row2Format, col2Format))
    .Font.Bold = True
    .NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
End With

End Sub

Sub determinationEstimateType()
'����������� ���� �����: ��� ��� ��
Dim currSheet As Variant

sheetName = "�����*"
For Each currSheet In Worksheets
    If currSheet.Name Like sheetName Then
        If InStr(currSheet.Name, "���") = 0 Then
            typeEstimate = "��"
            numberCol = 10
            letterCol = "J"
        Else
            typeEstimate = "���"
            numberCol = 11
            letterCol = "K"
        End If
    End If
Next

End Sub

Sub header()
'������������ �����, ������������ �����
Dim item As Variant
Dim i As Variant

Call activateSheet("����� *")

For Each item In Range("A1:K" & lastRow)
    If item Like "*�����������*" Then
        Rows("" & item.Row & ":" & item.Row + 5).Delete
        
    End If
Next
Rows("3:8").Insert
signer = "�.�. �����������"
position = "����������� ��������� ��� �.������ " & Chr(34) & "�������� ����������" & Chr(34)
Cells(3, 3) = Chr(34) & "���������" & Chr(34)
Cells(3, 3).HorizontalAlignment = xlLeft
Cells(5, 2) = "��������:"
Cells(6, 2) = position
Cells(7, 2) = "_________________________ " & signer
currYear = Format(Date, "yyyy")
Cells(8, 2) = Chr(34) & "_____" & Chr(34) & "___________________ " & currYear & " �."
Cells(3, 3).Font.Bold = True
With Range("B6:E6, B7:D7, B8:D8")
    .MergeCells = True
    .WrapText = True
    .HorizontalAlignment = xlLeft
End With
Call heightAdjustment(Range("B6:D6"))
Call countEstimate

For i = 0 To numberEstimates - 1
    Cells(nameLocation(i) + 3, 1).Value = Cells(nameLocation(i), 1).Value & i + 1
    Cells(nameLocation(i) + 5, 1).Value = smetaName(i)
    Range(Cells(nameLocation(i), 1), Cells(nameLocation(i) + 1, numberCol + 1)).Clear
Next

Call heightAdjustment(Range("A10:K10"))
Call heightAdjustment(Range("A15:K15"))


With Range(Cells(3, 1), Cells(7, 2))
    .Font.Name = "Times New Roman"
    .Font.Size = 13
End With


End Sub

Sub heightAdjustment(mergedRange)
'���������� ������ ������������ �����
Dim myCell As Range, myLen As Integer, _
myWidth As Single, k As Single, n As Single

With mergedRange
    '������ ������������ ������ ������� ������
    .WrapText = True
    '������ ������������ ������ ����� ������ ������, ����� ��������� ���� ������ ������
    .RowHeight = Cells(mergedRange.Row, mergedRange.Column).Font.Size * 1.3
End With
myLen = Len(CStr(Cells(mergedRange.Row, mergedRange.Column)))
For Each myCell In mergedRange
    myWidth = myWidth + myCell.ColumnWidth
Next
n = 10
k = Cells(mergedRange.Row, mergedRange.Column).Font.Size / n
mergedRange.RowHeight = mergedRange.RowHeight * WorksheetFunction.RoundUp(myLen * k / myWidth, 0)

End Sub

Sub countEstimate()
'����������� ���������� ���� � �� �����������
Dim item As Variant
Dim quiestion As String

Set namePosition = New Dictionary

For Each item In Range("A1:K" & lastRow)
    If item Like "*��������� ����� �*" Then
        namePosition.Add item.Row, Cells((item.Row + 3), item.Column).Value
        
    End If
Next

item = ""

nameLocation = namePosition.Keys
smetaName = namePosition.Items
numberEstimates = UBound(smetaName) + 1
For Each item In smetaName
    If item Like "*���������*�����������*���������" Then
        quiestion = "� �����: " & item & " ��� � ��� �����?"
        answer = MsgBox(quiestion, vbYesNo)
    End If
Next
End Sub

Sub cancelMerge(numberRow)
'������ ����������� ����� � ������� �������

If typeEstimate = "���" Then
    Range("H" & numberRow & ":I" & numberRow).UnMerge
    Cells(numberRow, 9).formula = Cells(numberRow, 8).formula
    Cells(numberRow, 8).Clear
    Call setFormat(numberRow, 9, numberRow, 9)

End If

Range(Split(Range(letterCol & numberRow).Offset(, -1).Address, "$")(1) & numberRow & ":" & letterCol & numberRow).UnMerge
Cells(numberRow, numberCol).formula = Cells(numberRow, numberCol - 1).formula
Cells(numberRow, numberCol - 1).Clear
Call setFormat(numberRow, numberCol, numberRow, numberCol)

Columns(letterCol & ":" & letterCol).EntireColumn.AutoFit


End Sub

Sub NDSIncluding(numberRow)
'���: � ��� �����

Cells(numberRow, 1).Value = "� ��� ����� ��� 20%"
Cells(numberRow, numberCol).formula = "=round(" & letterCol & numberRow - 1 & "*20/120, 2)"
Range("A" & numberRow & ":A" & numberRow).WrapText = False
Call setFormat(numberRow, numberCol, numberRow, numberCol)

End Sub

Sub finishOrGoToMainMenu()
'��������� ���������� ����� ��� ������� � �������� ����
Dim passing As Integer

passing = MsgBox("������� OK ��� ����������� ���������� ����� ��� Cancel ��� ����������", vbOKCancel)

If passing = 1 Then
    Call userFormEstimate
Else
    Exit Sub
End If

End Sub

Sub coefBudgetFinancing()

coefficientName = "������������� ���������� ��������������"
Call completeAddCoef

End Sub

Sub addCoef(numberRow)
'���������� ������� ��� ������������
Rows((numberRow + 1) & ":" & (numberRow + 2)).Insert

Cells(numberRow + 1, 1).Value = "����� � " & coefficientName & " K =" & coefficient
Cells(numberRow + 1, numberCol).formula = "=round(" & letterCol & numberRow & "*" & coefficient & ",2)"
Call NDSIncluding(numberRow + 2)

End Sub

Sub completeAddCoef()
'���������� ������������
Dim i As Variant


If typeEstimate = "" Then
    Call determinationEstimateType
End If

coefficient = InputBox("��������� �� ���������� � ������� �������� ������������")

Call activateSheet("*�����*")
If lastRow = 0 Then lastRow = seachLastCell() + 1

Set seachRange = Range(Cells(1, 1), Cells(lastRow, 11))
seachStr = "?????* ���*"
Set rowForCoefficient = Seach(seachStr, seachRange)
Call quickSort(rowForCoefficient, 1, rowForCoefficient.Count)

For i = 0 To rowForCoefficient.Count - 1
    Call addCoef(rowForCoefficient(i + 1) + i * 2)
    
Next

Call activateSheet("���")

End Sub

