Attribute VB_Name = "form/userFormEstimate"
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
Dim answer As Variant 'îòâåò íà âîïðîñ îá ÍÄÑ â òîì ÷èñëå
Dim signer As String 'ÔÈÎ óòâåðæäàþùåãî
Dim position As String 'äîëæíîñòü óòâåðæäàþùåãî
Dim typeEstimate As String 'òèï ñìåòû: ÒÑÍ èëè ÑÍ
Dim totalEstimate As New Collection 'íîìåð ñòðîêè èòîãî ïî ñìåòå
Dim rowForCoefficient As New Collection 'íîìåð ñòðîêè äëÿ âûâîäà êîýôôèöèåíòà
Dim nameLocation() As String 'íîìåð ðÿäà ðàñïîëîæåíèÿ íàèìåíîâàíèÿ ñìåòû
Dim smetaName() As String 'íàèìåíîâàíèå ñìåòû
Dim numberEstimates As Integer 'êîëè÷åñòâî ñìåò
Dim coefficientName As String 'íàèìåíîâàíèå êîýôôèöèåòà
Dim coefficient As Variant 'çíà÷åíèå êîýôôèöèåíòà

Sub userFormEstimate()

prepareEstimate.Show

End Sub

Sub nds()
Dim i As Variant
Dim item As Variant

Call activateSheet("Ñìåòà *")
lastRow = seachLastCell() + 1
Set seachRange = Range(Cells(1, 1), Cells(lastRow, 11))
Call determinationEstimateType
Call header
Call clearTail

If numberEstimates > 1 Then
    For i = 1 To totalEstimate.Count - 1
        Range("A" & totalEstimate(i) + 1 & ":A" & totalEstimate(i) + 3).EntireRow.Insert
        Range("A" & totalEstimate(i) & ":H" & totalEstimate(i)).Value = "Èòîãî ïî ëîêàëüíîé ñìåòå ¹" & i & ": " & smetaName(i - 1)
        Call heightAdjustment(Range("A" & totalEstimate(i) & ":H" & totalEstimate(i)))
        If answer = 6 And Cells(totalEstimate(i), 1).Value Like "*Ñòîèìîñòü*ïîñàäî÷íîãî*ìàòåðèàëà*" Then
            Call NDSIncluding(totalEstimate(i) + 1)
        Else
            Call ndsTotal(totalEstimate(i), numberCol, letterCol)
        End If
        totalEstimate.Add totalEstimate(i + 1) + i * 3, , , i
        totalEstimate.Remove (i + 2)
    Next
End If

Range("A" & totalEstimate(totalEstimate.Count) + 1 & ":A" & totalEstimate(totalEstimate.Count) + 3).EntireRow.Insert
Range("A" & totalEstimate(totalEstimate.Count) & ":G" & totalEstimate(totalEstimate.Count)).Value = "Èòîãî ïî ñìåòå: " & smetaName(0)
If answer = 6 Then
    Call NDSIncluding(totalEstimate(totalEstimate.Count) + 1)
Else
    Call ndsTotal(totalEstimate(totalEstimate.Count), numberCol, letterCol)
End If
Call heightAdjustment(Range("A" & totalEstimate(totalEstimate.Count) & ":G" & totalEstimate(totalEstimate.Count)))
For Each item In totalEstimate
    Call cancelMerge(item)
Next

If typeEstimate = "ÒÑÍ" Then
    Call ndsTotal(totalEstimate(1), 9, "I")
    Call setFormat(totalEstimate(1), 9, totalEstimate(1) + 2, 9)
End If

Call finishOrGoToMainMenu

End Sub

Function seachLastCell()
' ïîèñê ïîñëåäíåé íåïóñòîé ÿ÷åéêè â ñòîëáöàõ ñ 1-ãî ïî 11-é
    Dim c(11) As Integer
    Dim i As Variant
    
    For i = 1 To 11
        c(i) = Cells(Rows.Count, i).End(xlUp).Row
    Next
    seachLastCell = WorksheetFunction.Max(c)
End Function

Sub activateSheet(sheetName)
' àêòèâèðîâàíèå ëèñòà
Dim currSheet As Variant

For Each currSheet In Worksheets
    If currSheet.Name Like sheetName Then
        currSheet.Activate
        sheetName = ActiveSheet.Name
    End If
Next

End Sub

Function Seach(seachStr, seachRange) As Collection
'ïîèñê ïî ñòðîêå è ñîõðàíåíèå íîìåðà ðÿäà â êîëëåêöèþ
Dim foundCell As Range
Dim firstFoundCell As Range

Set Seach = New Collection

Set foundCell = seachRange.Find(seachStr, LookIn:=xlValues, MatchCase:=True)
Set firstFoundCell = foundCell

If firstFoundCell Is Nothing Then
    MsgBox (seachStr & " íå íàéäåíî")
    Exit Function
End If

Do
    Set foundCell = seachRange.FindNext(After:=foundCell)
    Seach.Add foundCell.Row
Loop While foundCell.Address <> firstFoundCell.Address

End Function

Sub quickSort(coll As Collection, first As Long, last As Long)
'áûñòðàÿ ñîðòèðîâêà ýëåìåíòîâ êîëëåêöèè
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
    ' Ïîìåíÿòü çíà÷åíèÿ
        temp = coll(low)
        coll.Add coll(high), After:=low
        coll.Remove low
        coll.Add temp, Before:=high
        coll.Remove high + 1
        ' Ïåðåéòè ê ñëåäóþùèì ïîçèöèÿì
        low = low + 1
        high = high - 1
    End If
    Loop
    If first < high Then quickSort coll, first, high
    If low < last Then quickSort coll, low, last
End Sub

Sub clearTail()
'î÷èñòêà, óäàëåíèå îáúåäèíåíèÿ ìåæäó Èòîãî ïî ëîêàëüíîé ñìåòå... è Ñîñòàâèë-Ïðîâåðèë
Dim i As Variant

Call activateSheet("Ñìåòà *")
Set seachRange = Range(Cells(1, 1), Cells(lastRow, 9))
seachStr = "Èòîãî ïî*ñìåòå*"
Set totalEstimate = Seach(seachStr, seachRange)
Call quickSort(totalEstimate, 1, totalEstimate.Count)

For i = 1 To totalEstimate.Count
    If i = totalEstimate.Count Then
        If numberEstimates = 1 Then
            Range("A" & totalEstimate(i) + 1 & ":A" & lastRow).EntireRow.Delete
            
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
'ðàñ÷åò è âûâîä ÍÄÑ è èòîãî ñ ÍÄÑ

Cells(r + 1, 1).Value = "ÍÄÑ 20%"
Cells(r + 1, numberCol).formula = "=round(" & letterCol & r & "*0.2,2)"
Cells(r + 2, 1) = "Èòîãî ñ ÍÄÑ 20%"
Cells(r + 2, numberCol).formula = "=" & letterCol & r & "+" & letterCol & (r + 1)
Range("A" & r + 1 & ":A" & r + 2).WrapText = False
Call setFormat(r, numberCol, r + 2, numberCol)

End Sub

Sub setFormat(row1Format, col1Format, row2Format, col2Format)
'ôîðìàòèðîâàíèå äèàïàçîíà
With Range(Cells(row1Format, col1Format), Cells(row2Format, col2Format))
    .Font.Bold = True
    .NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
End With

End Sub

Sub determinationEstimateType()
'îïðåäåëåíèå òèïà ñìåòû: ÒÑÍ èëè ÑÍ
Dim currSheet As Variant

sheetName = "Ñìåòà*"
For Each currSheet In Worksheets
    If currSheet.Name Like sheetName Then
        If InStr(currSheet.Name, "ÒÑÍ") = 0 Then
            typeEstimate = "ÑÍ"
            numberCol = 10
            letterCol = "J"
        Else
            typeEstimate = "ÒÑÍ"
            numberCol = 11
            letterCol = "K"
        End If
    End If
Next

End Sub

Sub header()
'Ôîðìèðîâàíèå øàïêè, íàèìåíîâàíèÿ ñìåòû
Dim item As Variant
Dim i As Variant

Call activateSheet("Ñìåòà *")

For Each item In Range("A1:K" & lastRow)
    If item Like "*ÑÎÃËÀÑÎÂÀÍÎ*" Then
        Rows("" & item.Row & ":" & item.Row + 5).Delete
        
    End If
Next
Rows("3:8").Insert
signer = "Å.È. Íîâîùèíñêàÿ"
position = "Çàìåñòèòåëü äèðåêòîðà ÃÊÓ ã.Ìîñêâû " & Chr(34) & "Äèðåêöèÿ Ìîñïðèðîäû" & Chr(34)
Cells(3, 3) = Chr(34) & "ÓÒÂÅÐÆÄÀÞ" & Chr(34)
Cells(3, 3).HorizontalAlignment = xlLeft
Cells(5, 2) = "Çàêàç÷èê:"
Cells(6, 2) = position
Cells(7, 2) = "_________________________ " & signer
currYear = Format(Date, "yyyy")
Cells(8, 2) = Chr(34) & "_____" & Chr(34) & "___________________ " & currYear & " ã."
Cells(3, 3).Font.Bold = True
With Range("B6:E6, B7:D7, B8:D8")
    .MergeCells = True
    .WrapText = True
    .HorizontalAlignment = xlLeft
End With
Call heightAdjustment(Range("B6:D6"))
Call countEstimate

For i = 1 To numberEstimates
    Cells(nameLocation(i - 1) + 3, 1).Value = Cells(nameLocation(i - 1), 1).Value & i + 1
    Cells(nameLocation(i - 1) + 5, 1).Value = smetaName(i - 1)
    Range(Cells(nameLocation(i - 1), 1), Cells(nameLocation(i - 1) + 1, numberCol + 1)).Clear
Next

Call heightAdjustment(Range("A10:K10"))
Call heightAdjustment(Range("A15:K15"))


With Range(Cells(3, 1), Cells(7, 2))
    .Font.Name = "Times New Roman"
    .Font.Size = 13
End With


End Sub

Sub heightAdjustment(mergedRange)
'àâòîïîäáîð âûñîòû îáúåäèíåííûõ ÿ÷ååê
Dim myCell As Range, myLen As Integer, _
myWidth As Single, k As Single, n As Single

With mergedRange
    'Çàäàåì îáúåäèíåííîé ÿ÷åéêå ïåðåíîñ òåêñòà
    .WrapText = True
    'Çàäàåì îáúåäèíåííîé ÿ÷åéêå òàêóþ âûñîòó ñòðîêè, ÷òîáû óìåùàëàñü îäíà ñòðîêà òåêñòà
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
'îïðåäåëåíèå êîëè÷åñòâà ñìåò è èõ íàìåíîâàíèé
Dim i As Variant
Dim item As Variant
Dim quiestion As String
Dim tempLocation As String
Dim tempName As String

For Each item In Range("A1:K" & lastRow)
    If item Like "*ËÎÊÀËÜÍÀß ÑÌÅÒÀ ¹*" Then
         tempLocation = tempLocation & item.Row & " "
    End If
Next
tempLocation = Trim(tempLocation)
nameLocation = Split(tempLocation, " ")

Sheets("Source").Activate
For i = 1 To lastRow
    If Cells(i, 6).HasFormula = False And Cells(i, 6).Value = "Íîâàÿ ëîêàëüíàÿ ñìåòà" Then
        tempName = tempName & Cells(i, 7).Value & ";"
    End If
Next
tempName = Trim(tempName)
smetaName = Split(tempName, ";")

For Each item In smetaName
    If item Like "*Ñòîèìîñòü*ïîñàäî÷íîãî*ìàòåðèàëà" Then
        quiestion = "Â ñìåòå: " & item & " ÍÄÑ â òîì ÷èñëå?"
        answer = MsgBox(quiestion, vbYesNo)
    End If
Next

numberEstimates = UBound(nameLocation) + 1

End Sub

Sub cancelMerge(numberRow)
'îòìåíà îáúåäèíåíèÿ ÿ÷ååê è ïåðåíîñ ôîðìóëû

If typeEstimate = "ÒÑÍ" Then
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
'ÍÄÑ: â òîì ÷èñëå

Cells(numberRow, 1).Value = "Â òîì ÷èñëå ÍÄÑ 20%"
Cells(numberRow, numberCol).formula = "=round(" & letterCol & numberRow - 1 & "*20/120, 2)"
Range("A" & numberRow & ":A" & numberRow).WrapText = False
Call setFormat(numberRow, numberCol, numberRow, numberCol)

End Sub

Sub finishOrGoToMainMenu()
'çàêîí÷èòü îôîðìëåíèå ñìåòû èëè ïåðåéòè â îñíîâíîå ìåíþ
Dim passing As Integer

passing = MsgBox("Íàæìèòå OK äëÿ ïðîäîëæåíèÿ îôîðìëåíèÿ ñìåòû èëè Cancel äëÿ çàâåðøåíèÿ", vbOKCancel)

If passing = 1 Then
    Call userFormEstimate
Else
    Exit Sub
End If

End Sub

Sub coefBudgetFinancing()

coefficientName = "êîýôôèöèåíòîì áþäæåòíîãî ôèíàíñèðîâàíèÿ"
Call completeAddCoef

End Sub

Sub addCoef(numberRow)
'äîáàâëåíèå ôîðìóëû äëÿ êîýôôèöèåíòà
Rows((numberRow + 1) & ":" & (numberRow + 2)).Insert

Cells(numberRow + 1, 1).Value = "Èòîãî ñ " & coefficientName & " K =" & coefficient
Cells(numberRow + 1, numberCol).formula = "=round(" & letterCol & numberRow & "*" & coefficient & ",2)"
Call NDSIncluding(numberRow + 2)

End Sub

Sub completeAddCoef()
'äîáàâëåíèå êîýôôèöèåíòà
Dim i As Variant


If typeEstimate = "" Then
    Call determinationEstimateType
End If

coefficient = InputBox("Ïåðåéäèòå íà àíãëèéñêèé è ââåäèòå çíà÷åíèå êîýôôèöèåíòà")

Call activateSheet("*Ñìåòà*")
If lastRow = 0 Then lastRow = seachLastCell() + 1

Set seachRange = Range(Cells(1, 1), Cells(lastRow, 11))
seachStr = "?????* ÍÄÑ*"
Set rowForCoefficient = Seach(seachStr, seachRange)
Call quickSort(rowForCoefficient, 1, rowForCoefficient.Count)

For i = 0 To rowForCoefficient.Count - 1
    Call addCoef(rowForCoefficient(i + 1) + i * 2)
    
Next

Call activateSheet("ÏÍÖ")

End Sub

