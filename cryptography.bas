Attribute VB_Name = "cryptography"
Option Explicit
Dim alphabet
Dim position As Integer

Sub cesar()

Dim sypher As String
Dim alphabet As String
Dim str1 As String, str2 As String, str3 As String, str4 As String, str5 As String
Dim str6 As String, str7 As String, str8 As String, str9 As String, str10 As String
Dim i As Integer
Dim position As Integer
Dim letter As String
Dim Shift As Integer

Shift = 2

str1 = "חתסת¸טשכןüתלד¸üכאסלטכזט‎חתהלדי"
str2 = "טבת¸םהכלתüÿךחדלÿüכÿכןÿזץןךתחד¸דףתד"
str3 = "כד‎חת¸דגתנדדüץûךתחחט‎טûתחותוט‏"
str4 = "טüטÿחתגüתחדÿ‏טזתככÿהמטזגתלÿזחÿחד"
str5 = "ÿיÿךÿבדüתר"
str6 = "ץנלכ‏צצזעמןיזצטלךןי‏ÿשחאןידגסü"
str7 = "קזףןללÿקדכז‎ףםדמדףלגזךכ‏צזעמאז"
str8 = "ודכדמ‏טיüץדאלדןילאלûנלןטמשנשח"
str9 = "ןזךאליןךדמנזכ‏ךלדחיüÿזךלחט‏מנ"
str10 = "זכדבליתÿדחכ‏טלכדפ"
sypher = str7 & str8 & str9 & str10
alphabet = "אבגדהו¸זחטיךכלםמןנסעףפץצקרשתûü‎‏ÿ"


For i = 1 To Len(sypher)
    letter = Mid(sypher, i, 1)
    position = InStr(alphabet, letter)
    If position + Shift > 33 Then
        position = position + Shift - 33
    Else
        position = position + Shift
    End If
    letter = Mid(alphabet, position, 1)
    Debug.Print letter
Next


End Sub


Sub polibiusSquare()

Dim alphabet As Variant
Dim sypher As Variant
Dim x As Integer
Dim y As Integer
Dim res As String
Dim i As Integer

alphabet = Range("G85:L90")
sypher = "331115111343534315244232116243331542133211246123336133426131141163116133112463612465615211316133266311326311211124553111262143136161142443526242522542446351133651111365336152254254144343213561334252614143244315334333613132134323112624433343214561514233366124311461536361515551343542612361244351556243336164"
res = ""

For i = 1 To Len(sypher) Step 2
    x = Mid(sypher, i, 1)
    y = Mid(sypher, i + 1, 1)
    res = res + alphabet(x, y)
Next
Debug.Print res



End Sub


Sub vigenereCipherEncrypt()

Dim str1 As String
Dim str2 As String
Dim str3 As String
Dim str4 As String
Dim str5 As String
Dim str6 As String
Dim str7 As String
Dim str8 As String
Dim str0 As String

Dim res As String
Dim codeWord As String
Dim tempRes As String

Dim letter As String
Dim i As Variant
Dim j As Variant
Dim letterCode As String
Dim Shift As Integer


'alphabet = Array("א", "ב", "ג", "ד", "ה", "ו", "¸", "ז", "ח", "ט", "י", "ך", "כ", "ל", "ם", "מ", "ן", "נ", "ס", "ע", "ף", "פ", "ץ", "צ", "ק", "ר", "ש", "ת", "û", "ü", "‎", "‏", "ÿ")
'shift = Array(24, 5, 17, 5, 16)
str1 = "טרנפÿצ‎ףכתחה¸ףץ‎וענû"
str2 = "יבע¸‎םףפתס‎אפץסדףפגד"
str3 = "בחצךץ‎¸טנמיÿך‎םדפפסי"
str4 = "‎¸מובאקסץאטז‎רצÿמושר"
str5 = "רוךףתהגזהûתףף¸חתתצכא"
str6 = "טבץפ‏אÿמûערüצ¸הסץפüש"
str7 = "שבופדה¸חלד‎יפפעךךףשץ"
str7 = "חזיקüסעךתûזגף¸ףאנםנü"
str8 = "‎דפףצח"
'str0 = str1 + str2 + str3 + str4 + str5 + str6 + str7 + str8
str0 = "לךחןהחןקתעףעךענטגמפליגדבתÿםהמסבפב¸ןדיהעיוףףהעתץ‏צנסגק¸עויזפהמףבייתק‏פאעגהגûםטץםףסףאןןÿרדזךנÿץמ¸¸אגככעüגבםצמהץג¸ûהנגכמה‎ףיד¸¸¸¸ברכÿעתתר‏כןהלת¸ב‏אסטעבננסגלףזחיפפטבאתמךל‎צÿכדמדברדץ¸זנסüוןמטת‏צדבק¸בפתנהקיגüפןףלבקחף¸והןדף¸דפ¸נגשעלתםתץכן‎דבהטכ‎כל¸‏המזנ"

j = 1

alphabet = "אבגדהו¸זחטיךכלםמןנסעףפץצקרשתûü‎‏ÿ"
codeWord = "קונון"

For i = 1 To Len(str0)
    letter = Mid(str0, i, 1)
    position = InStr(alphabet, letter)
    If j > Len(codeWord) Then j = 1
    letterCode = Mid(codeWord, j, 1)
    Shift = InStr(alphabet, letterCode) - 1
    position = ((position + Shift) Mod 33)
    If position = 0 Then position = 33
    letterCode = Mid(codeWord, j, 1)
    tempRes = Mid(alphabet, position, 1)
    res = res + tempRes
    j = j + 1
Next
Debug.Print res

End Sub

Sub vigenereCipherDecrypt()

Dim str1 As String
Dim str2 As String
Dim str3 As String
Dim str4 As String
Dim str5 As String
Dim str6 As String
Dim str7 As String
Dim str8 As String
Dim str0 As String

Dim res As String
Dim codeWord As String
Dim tempRes As String

Dim letter As String
Dim i As Variant
Dim j As Variant
Dim letterCode As String
Dim Shift As Integer


'alphabet = Array("א", "ב", "ג", "ד", "ה", "ו", "¸", "ז", "ח", "ט", "י", "ך", "כ", "ל", "ם", "מ", "ן", "נ", "ס", "ע", "ף", "פ", "ץ", "צ", "ק", "ר", "ש", "ת", "û", "ü", "‎", "‏", "ÿ")
'shift = Array(24, 5, 17, 5, 16)
str1 = "הןרפףÿפחÿגךקûקאאזÿשü"
str2 = "יבע¸‎םףפתס‎אפץסדףפגד"
str3 = "בחצךץ‎¸טנמיÿך‎םדפפסי"
str4 = "‎¸מובאקסץאטז‎רצÿמושר"
str5 = "רוךףתהגזהûתףף¸חתתצכא"
str6 = "טבץפ‏אÿמûערüצ¸הסץפüש"
str7 = "שבופדה¸חלד‎יפפעךךףשץ"
str7 = "חזיקüסעךתûזגף¸ףאנםנü"
str8 = "‎דפףצח"
'str0 = str1 + str2 + str3 + str4 + str5 + str6 + str7 + str8
str0 = "מוüונםלםüץחךו¸םמüחנךוץאןמאמחאדצוכואמוץכאץקאזודאמכוךכוהוץרדוהםנאאצץםז¸חזûחףכבנבונקמאזûחעצםםהתאמחזאגםץםüמםמאדךנםןוץםמםהשאבחמאץדצאלכאברבפתחאןאץםבר¸םמאס"

alphabet = "אבגדהו¸זחטיךכלםמןנסעףפץצקרשתûü‎‏ÿ"
codeWord = "קקווננווןן"
j = 1

For i = 1 To Len(str0)
    letter = Mid(str0, i, 1)
    position = InStr(alphabet, letter)
    If j > Len(codeWord) Then j = 1
    letterCode = Mid(codeWord, j, 1)
    Shift = InStr(alphabet, letterCode)
    position = (Abs(position - Shift + 33)) Mod 33
    If position = 0 Then position = 33
    letterCode = Mid(codeWord, j, 1)
    tempRes = Mid(alphabet, position, 1)
    res = res + tempRes
    j = j + 1
Next
Debug.Print res
End Sub

Sub encodeMessage()

Dim string1 As String
Dim string2 As String
Dim coordinate As String
Dim notebook As String
Dim letter As String
Dim number1 As String
Dim number2 As String
Dim number As String
Dim i As Variant
Dim res As String
Dim binary As Variant

binary = Array("000", "001", "010", "011", "100", "101", "110", "111")


res = ""
coordinate = InputBox("Êממנהטםאעû סממבשוםטÿ")

For i = 1 To Len(coordinate)
   number = Val(Mid(coordinate, i, 1))
   string1 = string1 + binary(number)
Next

notebook = InputBox("Áכמךםמע")

For i = 1 To Len(notebook)
    letter = Mid(notebook, i, 1)
    If InStr("אוטמ‎‏ÿ", letter) Then
        string2 = string2 + "1"
    Else
        string2 = string2 + "0"
    End If
Next

For i = 1 To Len(string1)
    number1 = Mid(string1, i, 1)
    number2 = Mid(string2, i, 1)
    If number1 <> number2 Then
        res = res + "1"
    Else
        res = res + "0"
    End If
Next

Debug.Print res

End Sub

Sub encodeAngles()

Dim angles As String
Dim binAngles As String
Dim newspaper As String
Dim binNewspaper As String
Dim number As String
Dim number1 As String
Dim number2 As String
Dim letter As String
Dim res As String
Dim text As String
Dim i As Variant
Dim binary As Variant
Dim answer As Variant
Dim n As Integer
Dim x As Integer
Dim y As Integer

angles = "eenennnessnewwsweeswssseeeswnnswnnnewwnwwwnwnnnessswnnnewwsessswwwsessneeenwwwnweenwwwnwnnneeenewwnweenennnweenennnesssennsw _
"
For i = 1 To Len(angles) Step 2
    letter = Mid(angles, i, 2)
    Select Case letter
        Case "nn"
            binAngles = binAngles + "00"
        Case "ee"
            binAngles = binAngles + "01"
        Case "ss"
            binAngles = binAngles + "10"
        Case "ww"
            binAngles = binAngles + "11"
        Case "se"
            binAngles = binAngles + "00"
        Case "sw"
            binAngles = binAngles + "01"
        Case "nw"
            binAngles = binAngles + "10"
        Case "ne"
            binAngles = binAngles + "11"
    End Select
Next

binary = Array("001", "010", "011", "100", "101", "110", "111")

newspaper = "ןמהחטלםÿÿןמסאהךאמגמשויםוןנטץמעטחבאכמגאםםûץעוןכûלךכטלאעמל‏זאם"
For i = 1 To Len(newspaper)
    letter = Mid(newspaper, i, 1)
    If InStr("אוטמ‎‏ÿ", letter) Then
        binNewspaper = binNewspaper + "1"
    Else
        binNewspaper = binNewspaper + "0"
    End If
Next

For i = 1 To Len(angles)
    number1 = Mid(binAngles, i, 1)
    number2 = Mid(binNewspaper, i, 1)
    If number1 <> number2 Then
        res = res + "1"
    Else
        res = res + "0"
    End If
Next

For i = 1 To Len(res) Step 3
    number = Mid(res, i, 3)
    n = WorksheetFunction.Match(number, binary, 0)
    answer = answer & n
Next

alphabet = Range("K208:R213")

For i = 1 To Len(answer) Step 2
    x = Mid(answer, i, 1)
    y = Mid(answer, i + 1, 1)
    text = text + alphabet(x, y)
Next

Debug.Print text

End Sub
