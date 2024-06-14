Attribute VB_Name = "TransferNumberToText"
Sub Перевести_выделенное_число_в_текст()
    Dim SumBase As Double, SumText As String
    With Selection
        SumText = .Text
        SumText = Replace(SumText, " ", "", 1, , vbTextCompare) ' Удаляем в числе пробелы
        SumText = Replace(SumText, "'", "", 1, , vbTextCompare) ' Удаляем в числе знаки '
        SumText = Replace(SumText, ",", ".", 1, , vbTextCompare) ' Меняем , на .
        SumText = Replace(SumText, Chr(160), "", 1, , vbBinaryCompare) ' Удаляем в числе неразрывные пробелы
        SumBase = Val(SumText)
        .Collapse Direction:=wdCollapseEnd
        .TypeText Text:=" " & Число_в_текст(SumBase, "руб")
    End With
End Sub
 
Public Function Число_в_текст(ByVal SumBase As Double, ByVal Valuta As String) As String
'Переводит цифровое значение в текстовое предложение.
'Параметр Valuta:
' "руб" - рубли,
' "дол" - доллары,
' "евр" - евро,
' "грив"- гривны,
' "" - без наименования,
' прочие текстовые наименования валют используются без склонения.
    Dim Edinicy(0 To 19) As String
    Dim Desyatki(0 To 9) As String
    Dim Sotni(0 To 9) As String
    Dim mlrd(0 To 9) As String
    Dim mln(0 To 9) As String
    Dim tys(0 To 9) As String
    Dim SumInt, x, shag, vl As Integer
    Dim txt, Sclon_Tys As String
    Dim Naim_Valuta_1 As String, Naim_Valuta_2 As String, Naim_Valuta_5 As String
    Dim Naim_Sotye_1 As String, Naim_Sotye_2 As String, Naim_Sotye_5 As String
    Dim Sotye As Integer, StrSotye As String
    Dim PereKluch  As String
    Edinicy(0) = ""
    Edinicy(1) = "один "
    Edinicy(2) = "два "
    Edinicy(3) = "три "
    Edinicy(4) = "четыре "
    Edinicy(5) = "пять "
    Edinicy(6) = "шесть "
    Edinicy(7) = "семь "
    Edinicy(8) = "восемь "
    Edinicy(9) = "девять "
    Edinicy(11) = "одиннадцать "
    Edinicy(12) = "двенадцать "
    Edinicy(13) = "тринадцать "
    Edinicy(14) = "четырнадцать "
    Edinicy(15) = "пятнадцать "
    Edinicy(16) = "шестнадцать "
    Edinicy(17) = "семнадцать "
    Edinicy(18) = "восемнадцать "
    Edinicy(19) = "девятнадцать "
    '---------------------------------------------
    Desyatki(0) = ""
    Desyatki(1) = "десять "
    Desyatki(2) = "двадцать "
    Desyatki(3) = "тридцать "
    Desyatki(4) = "сорок "
    Desyatki(5) = "пятьдесят "
    Desyatki(6) = "шестьдесят "
    Desyatki(7) = "семьдесят "
    Desyatki(8) = "восемьдесят "
    Desyatki(9) = "девяносто "
    '---------------------------------------------
    Sotni(0) = ""
    Sotni(1) = "сто "
    Sotni(2) = "двести "
    Sotni(3) = "триста "
    Sotni(4) = "четыреста "
    Sotni(5) = "пятьсот "
    Sotni(6) = "шестьсот "
    Sotni(7) = "семьсот "
    Sotni(8) = "восемьсот "
    Sotni(9) = "девятьсот "
    '---------------------------------------------
    mlrd(0) = "миллиардов "
    mlrd(1) = "миллиард "
    mlrd(2) = "миллиарда "
    mlrd(3) = "миллиарда "
    mlrd(4) = "миллиарда "
    mlrd(5) = "миллиардов "
    mlrd(6) = "миллиардов "
    mlrd(7) = "миллиардов "
    mlrd(8) = "миллиардов "
    mlrd(9) = "миллиардов "
    '---------------------------------------------
    mln(0) = "миллионов "
    mln(1) = "миллион "
    mln(2) = "миллиона "
    mln(3) = "миллиона "
    mln(4) = "миллиона "
    mln(5) = "миллионов "
    mln(6) = "миллионов "
    mln(7) = "миллионов "
    mln(8) = "миллионов "
    mln(9) = "миллионов "
    '---------------------------------------------
    tys(0) = "тысяч "
    tys(1) = "тысяча "
    tys(2) = "тысячи "
    tys(3) = "тысячи "
    tys(4) = "тысячи "
    tys(5) = "тысяч "
    tys(6) = "тысяч "
    tys(7) = "тысяч "
    tys(8) = "тысяч "
    tys(9) = "тысяч "
    '---------------------------------------------
    On Local Error Resume Next
    shag = 0
    SumInt = Int(SumBase)
    For x = Len(SumInt) To 1 Step -1
        shag = shag + 1
        Select Case x
            Case 12 ' - сотни миллиардов
                vl = Mid(SumInt, shag, 1)
                txt = txt & Sotni(vl)
            Case 11 ' - десятки  миллиардов
                vl = Mid(SumInt, shag, 1)
                If vl = "1" And Mid(SumInt, shag + 1, 1) <> 0 Then GoTo LblNextX Else txt = txt & Desyatki(vl)  ' - если конец триады от 11 до 19 то перескакиваем на единицы, иначе - формируем десятки
            Case 10 ' - единицы  миллиардов
                vl = Mid(SumInt, shag, 1)
                If shag > 1 Then
                    If Mid(SumInt, shag - 1, 1) = 1 Then txt = txt & Edinicy(Mid(SumInt, shag - 1, 2)) & "миллиардов " Else txt = txt & Edinicy(vl) & mlrd(vl) 'числа в диапозоне от 11 до 19 склоняются на "миллиардов" независимо от последнего числа триады
                Else
                    txt = txt & Edinicy(vl) & mlrd(vl)
                End If
 
                '-КОНЕЦ БЛОКА_______________________
            Case 9 ' - сотни миллионов
                vl = Mid(SumInt, shag, 1)
                txt = txt & Sotni(vl)
            Case 8 ' - десятки  миллионов
                vl = Mid(SumInt, shag, 1)
                If vl = "1" And Mid(SumInt, shag + 1, 1) <> 0 Then GoTo LblNextX Else txt = txt & Desyatki(vl)  ' - если конец триады от 11 до 19 то перескакиваем на единицы, иначе - формируем десятки
            Case 7 ' - единицы  миллионов
                vl = Mid(SumInt, shag, 1)
                If shag > 2 Then
                    If (Mid(SumInt, shag - 2, 1) = 0 And Mid(SumInt, shag - 1, 1) = 0 And vl = "0") Then GoTo LblNextX
                End If
                If shag > 1 Then
                    If Mid(SumInt, shag - 1, 1) = 1 Then txt = txt & Edinicy(Mid(SumInt, shag - 1, 2)) & "миллионов " Else: txt = txt & Edinicy(vl) & mln(vl)  'числа в диапозоне от 11 до 19 склоняются на "миллиардов" независимо от последнего числа триады
                Else
                    txt = txt & Edinicy(vl) & mln(vl)
                End If
                '-КОНЕЦ БЛОКА_______________________
            Case 6 ' - сотни тысяч
                vl = Mid(SumInt, shag, 1)
                txt = txt & Sotni(vl)
            Case 5 ' - десятки  тысяч
                vl = Mid(SumInt, shag, 1)
                If vl = 1 And Mid(SumInt, shag + 1, 1) <> 0 Then GoTo LblNextX Else txt = txt & Desyatki(vl)  ' - если конец триады от 11 до 19 то перескакиваем на единицы, иначе - формируем десятки
                Case 4 ' - единицы  тысяч
                vl = Mid(SumInt, shag, 1)
                If shag > 2 Then
                    If (Mid(SumInt, shag - 2, 1) = 0 And Mid(SumInt, shag - 1, 1) = 0 And vl = "0") Then GoTo LblNextX
                End If
                Sclon_Tys = Edinicy(vl) & tys(vl) ' - вводим переменную Sclon_Tys из-за иного склонения  тысяч в русском языке
                If vl = 1 Then Sclon_Tys = "одна " & tys(vl) ' - для тысяч склонение "один" и "два" неприменимо ( поэтому вводим переменную  Sclon_Tys )
                If vl = 2 Then Sclon_Tys = "две " & tys(vl) ' - для тысяч склонение "один" и "два" неприменимо ( поэтому вводим переменную  Sclon_Tys )
                If shag > 1 Then
                    If Mid(SumInt, shag - 1, 1) = 1 Then Sclon_Tys = Edinicy(Mid(SumInt, shag - 1, 2)) & "тысяч "
                End If
                txt = txt & Sclon_Tys
                '-КОНЕЦ БЛОКА_______________________
            Case 3 ' - сотни
                vl = Mid(SumInt, shag, 1)
                txt = txt & Sotni(vl)
            Case 2 ' - десятки
                vl = Mid(SumInt, shag, 1)
                If vl = "1" And Mid(SumInt, shag + 1, 1) <> 0 Then GoTo LblNextX Else txt = txt & Desyatki(vl)  ' - если конец триады от 11 до 19 то перескакиваем на единицы, иначе - формируем десятки
            Case 1 ' - единицы
                vl = Mid(SumInt, shag, 1)
                If shag > 2 Then
                    If (Mid(SumInt, shag - 2, 1) = 0 And Mid(SumInt, shag - 1, 1) = 0 And vl = "0") Then GoTo LblNextX
                End If
                If shag > 1 Then
                    If Mid(SumInt, shag - 1, 1) = 1 Then txt = txt & Edinicy(Mid(SumInt, shag - 1, 2)) Else: txt = txt & Edinicy(vl)
                Else
                    txt = txt & Edinicy(vl)
                End If
                '-КОНЕЦ БЛОКА_______________________
        End Select
LblNextX:
    Next x
    If InStr(1, LCase(Valuta), "руб") > 0 Then Valuta = "рубли"
    If InStr(1, LCase(Valuta), "дол") > 0 Then Valuta = "доллары"
    If InStr(1, LCase(Valuta), "евр") > 0 Then Valuta = "евро"
    If InStr(1, LCase(Valuta), "грив") > 0 Then Valuta = "гривны"
    Select Case Valuta
        Case "рубли"
            Naim_Valuta_1 = "рубль"
            Naim_Valuta_2 = "рубля"
            Naim_Valuta_5 = "рублей"
            Naim_Sotye_1 = "копейка"
            Naim_Sotye_2 = "копейки"
            Naim_Sotye_5 = "копеек"
        Case "доллары"
            Naim_Valuta_1 = "доллар"
            Naim_Valuta_2 = "доллара"
            Naim_Valuta_5 = "долларов"
            Naim_Sotye_1 = "цент"
            Naim_Sotye_2 = "цента"
            Naim_Sotye_5 = "центов"
        Case "евро"
            Naim_Valuta_1 = "евро"
            Naim_Valuta_2 = "евро"
            Naim_Valuta_5 = "евро"
            Naim_Sotye_1 = "цент"
            Naim_Sotye_2 = "цента"
            Naim_Sotye_5 = "центов"
        Case "гривны"
            Naim_Valuta_1 = "гривна"
            Naim_Valuta_2 = "гривны"
            Naim_Valuta_5 = "гривен"
            Naim_Sotye_1 = "копейка"
            Naim_Sotye_2 = "копейки"
            Naim_Sotye_5 = "копеек"
        Case ""
            Naim_Valuta_1 = ""
            Naim_Valuta_2 = ""
            Naim_Valuta_5 = ""
            Naim_Sotye_1 = ""
            Naim_Sotye_2 = ""
            Naim_Sotye_5 = ""
        Case Else
            Naim_Valuta_1 = Valuta
            Naim_Valuta_2 = Valuta
            Naim_Valuta_5 = Valuta
            Naim_Sotye_1 = "сотая"
            Naim_Sotye_2 = "сотых"
            Naim_Sotye_5 = "сотых"
    End Select
    If shag = 1 Then shag = 2
    If vl = 0 Or vl > 4 Or (Mid(SumInt, shag - 1, 2) > 10 And Mid(SumInt, shag - 1, 2) < 20) Then
        txt = txt + Naim_Valuta_5
    Else
        If vl = 1 Then txt = txt + Naim_Valuta_1 Else txt = txt + Naim_Valuta_2
    End If
    Sotye = CInt((SumBase - SumInt) * 100)
    StrSotye = Format(Sotye, "00")
    txt = txt & " " & StrSotye & " "
    Select Case Left(StrSotye, 1)
        Case "0", "2", "3", "4", "5", "6", "7", "8", "9"
            PereKluch = Right(StrSotye, 1)
        Case Else
            PereKluch = StrSotye
    End Select
    Select Case PereKluch
        Case "1"
            txt = txt & Naim_Sotye_1
        Case "2", "3", "4"
            txt = txt & Naim_Sotye_2
        Case Else
            txt = txt & Naim_Sotye_5
    End Select
    Число_в_текст = UCase(Left(txt, 1)) & Right(txt, Len(txt) - 1)
End Function


