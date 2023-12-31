Attribute VB_Name = "userFormEstimate"
Option Explicit

Public simpleFrameList(3)
Public complexFrameList(5)
Public executionFrameList(2)
Public typesOfWorksFrameList(5)

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
Dim nameLocation() As String '����� ���� ������������ ������������ �����
Dim smetaName() As String '������������ �����
Dim numberEstimates As Integer '���������� ����
Dim coefficientName As String '������������ �����������
Dim coefficient As Variant '�������� ������������

Sub userFormEstimate()

prepareEstimate.Show

End Sub

Sub nds()
Dim i As Variant
Dim item As Variant


If numberEstimates > 1 Then
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
Range("A" & totalEstimate(totalEstimate.Count) & ":G" & totalEstimate(totalEstimate.Count)).Value = "����� �� �����: " & smetaName(0)
If answer = 6 Then
    Call NDSIncluding(totalEstimate(totalEstimate.Count) + 1)
Else
    Call ndsTotal(totalEstimate(totalEstimate.Count), numberCol, letterCol)
End If
Call heightAdjustment(Range("A" & totalEstimate(totalEstimate.Count) & ":G" & totalEstimate(totalEstimate.Count)))
For Each item In totalEstimate
    Call cancelMerge(item)
Next

If typeEstimate = "���" Then
    Call ndsTotal(totalEstimate(1), 9, "I")
    Call setFormat(totalEstimate(1), 9, totalEstimate(1) + 2, 9)
End If

For i = LBound(simpleFrameList) To UBound(simpleFrameList)
        If simpleFrameList(i) = "financeCheckBox" Then
            Call coefBudgetFinancing
        End If
Next

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
'�������� ����� � ����� ����
Dim i As Variant
Dim tempOffset As Integer
Dim tempRow As Integer

Call activateSheet("����� *")
Set seachRange = Range(Cells(1, 1), Cells(lastRow, 9))
seachStr = "����� ��*�����*"
Set totalEstimate = Seach(seachStr, seachRange)
Call quickSort(totalEstimate, 1, totalEstimate.Count)

Dim rangeForClearing As Range

For i = 1 To totalEstimate.Count
    If i = totalEstimate.Count Then
        Set rangeForClearing = Range("A" & totalEstimate(i) + 1 & ":" & letterCol & lastRow)
    Else
        If i = numberEstimates Then
            tempRow = totalEstimate(i + 1) - tempOffset - 1
        Else
            tempRow = nameLocation(i) - 1
        End If
        Set rangeForClearing = Range("A" & totalEstimate(i) + 1 & ":" & letterCol & tempRow)
        tempOffset = tempOffset + (tempRow - (totalEstimate(i)))
        totalEstimate.Add totalEstimate(i + 1) - (tempOffset), , , i
        totalEstimate.Remove (i + 2)
            
    End If
    rangeForClearing.EntireRow.Delete
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
'������������ �����
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

Sub countEstimateAndSeachTitle()
'����������� ���������� ���� � �� �����������
Dim i As Variant
Dim item As Variant
Dim quiestion As String
Dim tempLocation As String
Dim tempName As String

For Each item In Range("A1:K" & lastRow)
    If item Like "*��������� ����� �*" Then
         tempLocation = tempLocation & item.Row & " "
    End If
Next
tempLocation = Trim(tempLocation)
nameLocation = Split(tempLocation, " ")

Sheets("SourceObSm").Activate
Cells(1, 10).Clear

Sheets("Source").Activate
Cells(1, 10).Clear
lastRow = seachLastCell() + 1

For i = 1 To lastRow
    If Cells(i, 6).HasFormula = False And Cells(i, 6).Value = "����� ��������� �����" Then
        tempName = tempName & Cells(i, 7).Value & ";"
    End If
Next
tempName = Left(tempName, Len(tempName) - 1)
smetaName = Split(tempName, ";")

For Each item In smetaName
    If item Like "*���������*�����������*���������" Then
        quiestion = "� �����: " & item & " ��� � ��� �����?"
        answer = MsgBox(quiestion, vbYesNo)
    End If
Next

numberEstimates = UBound(nameLocation) + 1

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
Dim j As Variant

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

Set seachRange = Range(Cells(1, 1), Cells(lastRow, 11))
seachStr = "��������� � ������ ���*"

Set rowForCoefficient = Seach(seachStr, seachRange)
Call quickSort(rowForCoefficient, 1, rowForCoefficient.Count)

Rows((rowForCoefficient(1) + 1) & ":" & (rowForCoefficient(1) + 2)).Insert

For i = 2 To 8
    numberCol = i
    letterCol = Split(Cells(rowForCoefficient(1), i).Address, "$")(1)
    Cells(rowForCoefficient(1) + 1, 1).Value = "����� � " & coefficientName & " K =" & coefficient
    Cells(rowForCoefficient(1) + 1, numberCol).formula = "=round(" & letterCol & rowForCoefficient(1) & "*" & coefficient & ",2)"
    Call NDSIncluding(rowForCoefficient(1) + 2)
Next

With Range("A" & rowForCoefficient(1) + 1 & ":H" & rowForCoefficient(1) + 2)
    .Borders.LineStyle = xlContinuous
    .Font.Bold = True
End With
With Range("A" & rowForCoefficient(1) + 1)
            .WrapText = True
            .HorizontalAlignment = xlLeft
            .EntireRow.AutoFit
End With

End Sub

Sub usn()

End Sub

Sub startEstimate()

Dim i As Variant

Call activateSheet("����� *")

lastRow = seachLastCell() + 1
Set seachRange = Range(Cells(1, 1), Cells(lastRow, 11))

Call determinationEstimateType

For i = LBound(simpleFrameList) To UBound(simpleFrameList)
    If simpleFrameList(i) = "NDSOptionButton" Then
        Call header
        Call countEstimateAndSeachTitle
        Call insertEstimateTitle
        lastRow = seachLastCell() + 1
        Set seachRange = Range(Cells(1, 1), Cells(lastRow, 11))
        Call clearTail
        Call nds
    ElseIf simpleFrameList(i) = "USNOptionButton" Then
        Call header
        Call countEstimateAndSeachTitle
        Call insertEstimateTitle
        Call clearTail
        Call usn
    ElseIf simpleFrameList(i) = "financeCheckBox" Then
        Call coefBudgetFinancing
    End If
Next


End Sub

Sub insertEstimateTitle()
'������� ��������� �����
Dim i As Variant

Call activateSheet("����� *")

For i = 1 To numberEstimates
    Cells(nameLocation(i - 1), 1).Value = Cells(nameLocation(i - 1), 1).Value & i
    Cells(nameLocation(i - 1) + 5, 1).Value = smetaName(i - 1)
    Cells(nameLocation(i - 1) + 3, 1).EntireRow.Clear
    Call heightAdjustment(Range("A" & nameLocation(i - 1) + 5 & ":K" & nameLocation(i - 1) + 5))
Next




End Sub



