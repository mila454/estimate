Attribute VB_Name = "TransferNumberToText"
Sub ���������_����������_�����_�_�����()
    Dim SumBase As Double, SumText As String
    With Selection
        SumText = .Text
        SumText = Replace(SumText, " ", "", 1, , vbTextCompare) ' ������� � ����� �������
        SumText = Replace(SumText, "'", "", 1, , vbTextCompare) ' ������� � ����� ����� '
        SumText = Replace(SumText, ",", ".", 1, , vbTextCompare) ' ������ , �� .
        SumText = Replace(SumText, Chr(160), "", 1, , vbBinaryCompare) ' ������� � ����� ����������� �������
        SumBase = Val(SumText)
        .Collapse Direction:=wdCollapseEnd
        .TypeText Text:=" " & �����_�_�����(SumBase, "���")
    End With
End Sub
 
Public Function �����_�_�����(ByVal SumBase As Double, ByVal Valuta As String) As String
'��������� �������� �������� � ��������� �����������.
'�������� Valuta:
' "���" - �����,
' "���" - �������,
' "���" - ����,
' "����"- ������,
' "" - ��� ������������,
' ������ ��������� ������������ ����� ������������ ��� ���������.
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
    Edinicy(1) = "���� "
    Edinicy(2) = "��� "
    Edinicy(3) = "��� "
    Edinicy(4) = "������ "
    Edinicy(5) = "���� "
    Edinicy(6) = "����� "
    Edinicy(7) = "���� "
    Edinicy(8) = "������ "
    Edinicy(9) = "������ "
    Edinicy(11) = "����������� "
    Edinicy(12) = "���������� "
    Edinicy(13) = "���������� "
    Edinicy(14) = "������������ "
    Edinicy(15) = "���������� "
    Edinicy(16) = "����������� "
    Edinicy(17) = "���������� "
    Edinicy(18) = "������������ "
    Edinicy(19) = "������������ "
    '---------------------------------------------
    Desyatki(0) = ""
    Desyatki(1) = "������ "
    Desyatki(2) = "�������� "
    Desyatki(3) = "�������� "
    Desyatki(4) = "����� "
    Desyatki(5) = "��������� "
    Desyatki(6) = "���������� "
    Desyatki(7) = "��������� "
    Desyatki(8) = "����������� "
    Desyatki(9) = "��������� "
    '---------------------------------------------
    Sotni(0) = ""
    Sotni(1) = "��� "
    Sotni(2) = "������ "
    Sotni(3) = "������ "
    Sotni(4) = "��������� "
    Sotni(5) = "������� "
    Sotni(6) = "�������� "
    Sotni(7) = "������� "
    Sotni(8) = "��������� "
    Sotni(9) = "��������� "
    '---------------------------------------------
    mlrd(0) = "���������� "
    mlrd(1) = "�������� "
    mlrd(2) = "��������� "
    mlrd(3) = "��������� "
    mlrd(4) = "��������� "
    mlrd(5) = "���������� "
    mlrd(6) = "���������� "
    mlrd(7) = "���������� "
    mlrd(8) = "���������� "
    mlrd(9) = "���������� "
    '---------------------------------------------
    mln(0) = "��������� "
    mln(1) = "������� "
    mln(2) = "�������� "
    mln(3) = "�������� "
    mln(4) = "�������� "
    mln(5) = "��������� "
    mln(6) = "��������� "
    mln(7) = "��������� "
    mln(8) = "��������� "
    mln(9) = "��������� "
    '---------------------------------------------
    tys(0) = "����� "
    tys(1) = "������ "
    tys(2) = "������ "
    tys(3) = "������ "
    tys(4) = "������ "
    tys(5) = "����� "
    tys(6) = "����� "
    tys(7) = "����� "
    tys(8) = "����� "
    tys(9) = "����� "
    '---------------------------------------------
    On Local Error Resume Next
    shag = 0
    SumInt = Int(SumBase)
    For x = Len(SumInt) To 1 Step -1
        shag = shag + 1
        Select Case x
            Case 12 ' - ����� ����������
                vl = Mid(SumInt, shag, 1)
                txt = txt & Sotni(vl)
            Case 11 ' - �������  ����������
                vl = Mid(SumInt, shag, 1)
                If vl = "1" And Mid(SumInt, shag + 1, 1) <> 0 Then GoTo LblNextX Else txt = txt & Desyatki(vl)  ' - ���� ����� ������ �� 11 �� 19 �� ������������� �� �������, ����� - ��������� �������
            Case 10 ' - �������  ����������
                vl = Mid(SumInt, shag, 1)
                If shag > 1 Then
                    If Mid(SumInt, shag - 1, 1) = 1 Then txt = txt & Edinicy(Mid(SumInt, shag - 1, 2)) & "���������� " Else txt = txt & Edinicy(vl) & mlrd(vl) '����� � ��������� �� 11 �� 19 ���������� �� "����������" ���������� �� ���������� ����� ������
                Else
                    txt = txt & Edinicy(vl) & mlrd(vl)
                End If
 
                '-����� �����_______________________
            Case 9 ' - ����� ���������
                vl = Mid(SumInt, shag, 1)
                txt = txt & Sotni(vl)
            Case 8 ' - �������  ���������
                vl = Mid(SumInt, shag, 1)
                If vl = "1" And Mid(SumInt, shag + 1, 1) <> 0 Then GoTo LblNextX Else txt = txt & Desyatki(vl)  ' - ���� ����� ������ �� 11 �� 19 �� ������������� �� �������, ����� - ��������� �������
            Case 7 ' - �������  ���������
                vl = Mid(SumInt, shag, 1)
                If shag > 2 Then
                    If (Mid(SumInt, shag - 2, 1) = 0 And Mid(SumInt, shag - 1, 1) = 0 And vl = "0") Then GoTo LblNextX
                End If
                If shag > 1 Then
                    If Mid(SumInt, shag - 1, 1) = 1 Then txt = txt & Edinicy(Mid(SumInt, shag - 1, 2)) & "��������� " Else: txt = txt & Edinicy(vl) & mln(vl)  '����� � ��������� �� 11 �� 19 ���������� �� "����������" ���������� �� ���������� ����� ������
                Else
                    txt = txt & Edinicy(vl) & mln(vl)
                End If
                '-����� �����_______________________
            Case 6 ' - ����� �����
                vl = Mid(SumInt, shag, 1)
                txt = txt & Sotni(vl)
            Case 5 ' - �������  �����
                vl = Mid(SumInt, shag, 1)
                If vl = 1 And Mid(SumInt, shag + 1, 1) <> 0 Then GoTo LblNextX Else txt = txt & Desyatki(vl)  ' - ���� ����� ������ �� 11 �� 19 �� ������������� �� �������, ����� - ��������� �������
                Case 4 ' - �������  �����
                vl = Mid(SumInt, shag, 1)
                If shag > 2 Then
                    If (Mid(SumInt, shag - 2, 1) = 0 And Mid(SumInt, shag - 1, 1) = 0 And vl = "0") Then GoTo LblNextX
                End If
                Sclon_Tys = Edinicy(vl) & tys(vl) ' - ������ ���������� Sclon_Tys ��-�� ����� ���������  ����� � ������� �����
                If vl = 1 Then Sclon_Tys = "���� " & tys(vl) ' - ��� ����� ��������� "����" � "���" ����������� ( ������� ������ ����������  Sclon_Tys )
                If vl = 2 Then Sclon_Tys = "��� " & tys(vl) ' - ��� ����� ��������� "����" � "���" ����������� ( ������� ������ ����������  Sclon_Tys )
                If shag > 1 Then
                    If Mid(SumInt, shag - 1, 1) = 1 Then Sclon_Tys = Edinicy(Mid(SumInt, shag - 1, 2)) & "����� "
                End If
                txt = txt & Sclon_Tys
                '-����� �����_______________________
            Case 3 ' - �����
                vl = Mid(SumInt, shag, 1)
                txt = txt & Sotni(vl)
            Case 2 ' - �������
                vl = Mid(SumInt, shag, 1)
                If vl = "1" And Mid(SumInt, shag + 1, 1) <> 0 Then GoTo LblNextX Else txt = txt & Desyatki(vl)  ' - ���� ����� ������ �� 11 �� 19 �� ������������� �� �������, ����� - ��������� �������
            Case 1 ' - �������
                vl = Mid(SumInt, shag, 1)
                If shag > 2 Then
                    If (Mid(SumInt, shag - 2, 1) = 0 And Mid(SumInt, shag - 1, 1) = 0 And vl = "0") Then GoTo LblNextX
                End If
                If shag > 1 Then
                    If Mid(SumInt, shag - 1, 1) = 1 Then txt = txt & Edinicy(Mid(SumInt, shag - 1, 2)) Else: txt = txt & Edinicy(vl)
                Else
                    txt = txt & Edinicy(vl)
                End If
                '-����� �����_______________________
        End Select
LblNextX:
    Next x
    If InStr(1, LCase(Valuta), "���") > 0 Then Valuta = "�����"
    If InStr(1, LCase(Valuta), "���") > 0 Then Valuta = "�������"
    If InStr(1, LCase(Valuta), "���") > 0 Then Valuta = "����"
    If InStr(1, LCase(Valuta), "����") > 0 Then Valuta = "������"
    Select Case Valuta
        Case "�����"
            Naim_Valuta_1 = "�����"
            Naim_Valuta_2 = "�����"
            Naim_Valuta_5 = "������"
            Naim_Sotye_1 = "�������"
            Naim_Sotye_2 = "�������"
            Naim_Sotye_5 = "������"
        Case "�������"
            Naim_Valuta_1 = "������"
            Naim_Valuta_2 = "�������"
            Naim_Valuta_5 = "��������"
            Naim_Sotye_1 = "����"
            Naim_Sotye_2 = "�����"
            Naim_Sotye_5 = "������"
        Case "����"
            Naim_Valuta_1 = "����"
            Naim_Valuta_2 = "����"
            Naim_Valuta_5 = "����"
            Naim_Sotye_1 = "����"
            Naim_Sotye_2 = "�����"
            Naim_Sotye_5 = "������"
        Case "������"
            Naim_Valuta_1 = "������"
            Naim_Valuta_2 = "������"
            Naim_Valuta_5 = "������"
            Naim_Sotye_1 = "�������"
            Naim_Sotye_2 = "�������"
            Naim_Sotye_5 = "������"
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
            Naim_Sotye_1 = "�����"
            Naim_Sotye_2 = "�����"
            Naim_Sotye_5 = "�����"
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
    �����_�_����� = UCase(Left(txt, 1)) & Right(txt, Len(txt) - 1)
End Function


