Attribute VB_Name = "KS_2_3"
Option Explicit
Public currWB As Workbook
Dim smetaNameRow As New Collection
Dim totalEstimate As New Collection
Dim totalBySection As New Collection



Sub KS_2_3()

'оформление КС-2 после выгрузки

Dim seachRange As Range, seachStr As String
Dim lastRow As Integer
Dim signiture As New Collection
Dim currRow As Variant
Dim construction As New Collection
Dim customer As New Collection
Dim accept As New Collection
Dim signer As String
Dim OKPO As String
Dim position As String
Dim kFin As String
Dim kDecline As String
Dim kName As String
Dim i As Integer
Dim temp As Integer
Dim smetaName As New Collection
Dim totalByAct As New Collection

Set currWB = ActiveWorkbook
Set smetaName = New Collection
Sheets("АктКС-2поТСН-2001(с доп.67").Activate

signer = "Е.И. Новощинская"
OKPO = "17785844"
position = "Заместитель директора ГКУ г.Москвы " & Chr(34) & "Дирекция Мосприроды" & Chr(34)
kName = "учетом коэффициента снижения по результатам открытого конкурса в электронной форме"

Sheets("Source").Cells(1, 10).Clear
Sheets("SourceObSm").Cells(1, 10).Clear

lastRow = ContractEstimate.seachLastCell()

Set seachRange = Range(Cells(1, 1), Cells(lastRow, 12))
seachStr = "Итого по разделу: *"
Set totalBySection = Seach(seachStr, seachRange)
Call quickSort.quickSort(totalBySection, 1, totalBySection.Count)

seachStr = "*Локальная смета:*"
Set smetaNameRow = Seach(seachStr, seachRange)
Call quickSort.quickSort(smetaNameRow, 1, smetaNameRow.Count)

For Each currRow In smetaNameRow
    smetaName.Add Split(Cells(currRow, 1).Value, ":")(1)
Next

For i = 1 To smetaNameRow.Count
    Cells(smetaNameRow(i), 1).Value = smetaName(i)
Next

seachStr = "Стройка*"
Set construction = Seach(seachStr, seachRange)
Cells(construction(1), 3).Value = smetaName(1)
Cells(construction(1) + 2, 3).Value = smetaName(1)
Call heightAdjustment.heightAdjustment(Range("C" & construction(1) & ":H" & construction(1)))
Call heightAdjustment.heightAdjustment(Range("C" & construction(1) + 2 & ":H" & construction(1) + 2))


seachStr = "Заказчик*"
Set customer = Seach(seachStr, seachRange)
Cells(customer(1), 3).Value = "ГКУ г.Москвы " & Chr(34) & "Дирекция Мосприроды" & Chr(34) & ", 117420, г.Москва, ул.Профсоюзная, д.41, тел. 8(495) 531-20-08"
Sheets("Source").Cells(15, 37).Value = OKPO
Call heightAdjustment.heightAdjustment(Range("C" & customer(1) & ":H" & customer(1)))

For Each currRow In totalBySection
    If Cells(currRow, 11).Value = 0 Then
        Rows(currRow - 3 & ":" & currRow).Interior.color = 65535
    End If
Next
currRow = 0

seachStr = "Итого по*смете*"
Set totalEstimate = Seach(seachStr, seachRange)
Call quickSort.quickSort(totalEstimate, 1, totalEstimate.Count)

If totalEstimate.Count > 1 Then
    For i = 1 To 2
        If Cells(totalEstimate(i), 11).Value = 0 Then
            Rows(totalEstimate(i)).Interior.color = 65535
            Rows(totalEstimate(i) + 3).Interior.color = 65535
            Rows(totalEstimate(i) + 4).Interior.color = 65535
            Rows(smetaNameRow(i)).Interior.color = 65535
        Else
            Cells(totalEstimate(i), 1).Value = "Итого по " & smetaName(i)
            If Cells(totalEstimate(i) + 3, 4).Value Like "*НДС 20%*" Then
                Rows(totalEstimate(i) + 3).Interior.color = 65535
                Rows(totalEstimate(i) + 4).Interior.color = 65535
            End If
        End If
    Next
Else
    Rows(totalEstimate(1)).EntireRow.Delete
End If

For Each currRow In seachRange
    If Rows(currRow.Row).Interior.color = 65535 Then
        Rows(currRow.Row).EntireRow.Delete
    End If
Next

seachStr = "Принял  *"
Set accept = Seach(seachStr, seachRange)
Cells(accept(1), 4).Value = "Заместитель директора ГКУ г.Москвы " & Chr(34) & "Дирекция Мосприроды" & Chr(34)
Cells(accept(1), 12).Value = signer


seachStr = "Сдал*"
Set signiture = Seach(seachStr, seachRange)



seachStr = "Итого по акту:*"
Set totalByAct = Seach(seachStr, seachRange)



Cells(totalByAct(1), 1).Value = "Итого по акту: " & smetaName(1)
Call cancelMerge("K", totalByAct(1), "L", totalByAct(1), 0)
Cells(totalByAct(1), 12).formula = "=SUM(P36:P" & totalByAct(1) & ")"


Rows(totalByAct(1) + 3).EntireRow.Clear

Range("A" & totalByAct(1) + 1 & ":A" & signiture(1) - 2).EntireRow.Delete
Rows(totalByAct(1) + 1).Insert
Rows(totalByAct(1) + 1).ClearFormats
Rows(totalByAct(1) + 1).Insert
Rows(totalByAct(1) + 1).ClearFormats
Call Estimate.ndsTotal(totalByAct(1), "L", 12)
temp = totalByAct(1)
Set totalByAct = New Collection
totalByAct.Add temp + 2

kFin = simpleInput("коэффициент бюджетного финансирования")
If kFin <> "" Then
    Call ContractEstimate.TotalWithK(kFin, "коэффициента бюджетного финансирования", totalByAct(1), totalByAct(1), "L", 1, 12)
    totalByAct.Add totalByAct(1) + 1
End If


kDecline = simpleInput("коэффициент снижения по итогам торгов")
If kDecline <> "" Then
    Call ContractEstimate.TotalWithK(kDecline, kName, totalByAct(totalByAct.Count) + 1, totalByAct(totalByAct.Count), "L", 1, 12)
    totalByAct.Add totalByAct(totalByAct.Count) + 2
End If


'заполнение КС-3

Sheets("Макет форма-3").Activate
Cells(12, 3).Value = smetaName(1)
Cells(8, 3) = Sheets("АктКС-2поТСН-2001(с доп.67").Cells(11, 3)
Call heightAdjustment.heightAdjustment(Range("C8:H8"))
Call heightAdjustment.heightAdjustment(Range("C12:H12"))
Cells(8, 11) = OKPO
Cells(14, 11).formula = "='АктКС-2поТСН-2001(с доп.67'!J20"
Cells(15, 11).formula = "='АктКС-2поТСН-2001(с доп.67'!J21"
Cells(21, 3).formula = "='АктКС-2поТСН-2001(с доп.67'!G27"
Cells(21, 9).formula = "='АктКС-2поТСН-2001(с доп.67'!I27"
Cells(21, 11).formula = "='АктКС-2поТСН-2001(с доп.67'!J27"
Cells(35, 2).Value = "В том числе:" & smetaName(1)
Range("B35:L35").UnMerge
Range("B35:E35").Merge
Call heightAdjustment.heightAdjustment(Range("B35:E35"))
Range("K35:L35").Merge
Cells(38, 11).formula = "='АктКС-2поТСН-2001(с доп.67'!L" & totalByAct(totalByAct.Count)
Cells(37, 11).formula = "=round(K38*20/120,2)"
Cells(36, 11).formula = "=K38-K37"

Cells(35, 11).formula = "=K36"
With Range("K35:L38")
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With
Range("G35:H35").Merge
Range("I35:J35").Merge
Range("F35:J35").Borders.LineStyle = xlContinuous

Call Estimate.setFormat(35, 11, 38, 11)

Cells(42, 3).Value = position
Call heightAdjustment.heightAdjustment(Range("C42:E42"))
Cells(42, 10).Value = signer

End Sub

Function simpleInput(ktype) As String
' Простой ввод: коэффициенты, наименование закупки
Dim strInput As String

strInput = "Введите " & ktype
simpleInput = InputBox(strInput)
If simpleInput = "" Then
    Exit Function
End If

End Function



