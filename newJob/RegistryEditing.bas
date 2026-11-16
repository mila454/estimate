Attribute VB_Name = "RegistryEditing"
Option Explicit

Sub design_list()
    Dim column_list As Variant
    Dim i As Long, j As Long, lastCol As Long
    Dim colName As Variant
    Dim foundCol As Long

    ' --- ЗАДАЕМ ЖЕСТКИЙ ПОРЯДОК СТОЛБЦОВ ---
    ' Порядок в этом массиве = порядок столбцов после работы макроса
    column_list = Array("Дата вложения", "Номер", "Сумма", "Сумма НДС")
    

    ' 1. Сначала удаляем ВСЁ, что не входит в наш список (и дубли)
    lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
    
    For i = lastCol To 1 Step -1
        Dim isNeeded As Boolean
        isNeeded = False
        
        For Each colName In column_list
            If StrComp(NormHeader(Cells(1, i).Value), CStr(colName), vbTextCompare) = 0 Then
                isNeeded = True
                Exit For
            End If
        Next colName
        
        ' Если столбец не нужен, или это дубль (проверяем левее) - удаляем
        If Not isNeeded Then
            Columns(i).Delete
        Else
            Dim isDup As Boolean
            isDup = False
            For j = 1 To i - 1
                If StrComp(NormHeader(Cells(1, j).Value), NormHeader(Cells(1, i).Value), vbTextCompare) = 0 Then
                    isDup = True
                    Exit For
                End If
            Next j
            If isDup Then Columns(i).Delete
        End If
    Next i

    ' 2. Теперь расставляем оставшиеся столбцы в нужном порядке (слева направо)
    ' Идем по нашему списку желаемого порядка
    For i = LBound(column_list) To UBound(column_list)
        foundCol = 0
        
        ' Ищем, в каком столбце сейчас находится нужный заголовок
        For j = 1 To Columns.Count
            If StrComp(NormHeader(Cells(1, j).Value), CStr(column_list(i)), vbTextCompare) = 0 Then
                foundCol = j
                Exit For
            End If
        Next j
        
        ' Если нашли и он стоит не на своем месте (не i+1), переставляем
        If foundCol > 0 And foundCol <> (i + 1) Then
            Columns(foundCol).Cut
            Columns(i + 1).Insert shift:=xlToRight
        End If
    Next i

End Sub


Function NormHeader(v As Variant) As String
    Dim s As String
    s = CStr(v)
    s = Replace(s, Chr(160), " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormHeader = Trim(s)
End Function
