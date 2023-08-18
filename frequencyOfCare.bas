Attribute VB_Name = "frequencyOfCare"
Dim frequency As Integer

Sub frequencyOfCare()
'кратность ухода

Dim care1 As New Collection
Dim care2 As New Collection
Dim care3 As New Collection
Dim god1 As String, god2 As String, god3 As String
Dim lastRow As Integer
Dim seachRange As Range, seachStr As String
Dim sheetName As String

sheetName = "Source"
For Each currSheet In Worksheets
    If currSheet.Name Like sheetName Then
        currSheet.Activate
    End If
Next

god1 = "2-й этап"
god2 = "3-й этап"
god3 = "4-й этап"
frequency = 2

lastRow = ContractEstimate.seachLastCell()
Set seachRange = Range(Cells(1, 7), Cells(lastRow, 7))

seachStr = "”ход*" & god1
Set care1 = Seach(seachStr, seachRange)
Call quickSort.quickSort(care1, 1, care1.Count)

seachStr = "”ход*" & god2
Set care2 = Seach(seachStr, seachRange)
Call quickSort.quickSort(care2, 1, care2.Count)

seachStr = "”ход*" & god3
Set care3 = Seach(seachStr, seachRange)
Call quickSort.quickSort(care3, 1, care3.Count)

Call changeFrequency(care3)

End Sub

Sub changeFrequency(care)
Dim i As Variant
Dim j As Variant

For i = 1 To care.Count Step 3
    Cells(care(i) + 4, 17).formula = "=(Round((Round((((ET" & care(i) + 4 & "*" & frequency & "))*AV" & care(i) + 4 & "*I" & care(i) + 4 & "),2)*BB" & care(i) + 4 & "),2)+Round((Round(((AE" & care(i) + 4 & "-((EU" & care(i) + 4 & "*" & frequency & ")))*AV" & care(i) + 4 & "*I" & care(i) + 4 & "),2)*BS" & care(i) + 4 & "),2))"
    Cells(care(i) + 5, 17).formula = "=(Round((Round((((ET" & care(i) + 5 & "*" & frequency & "))*AV" & care(i) + 5 & "*I" & care(i) + 5 & "),2)*BB" & care(i) + 5 & "),2)+Round((Round(((AE" & care(i) + 5 & "-((EU" & care(i) + 5 & "*" & frequency & ")))*AV" & care(i) + 5 & "*I" & care(i) + 5 & "),2)*BS" & care(i) + 5 & "),2))"
    Cells(care(i) + 4, 29).formula = "=Round(((ES" & care(i) + 4 & "*" & frequency & ")),6)"
    Cells(care(i) + 5, 29).formula = "=Round(((ES" & care(i) + 5 & "*" & frequency & ")),6)"
    Cells(care(i) + 4, 30).formula = "=Round(((((ET" & care(i) + 4 & "*" & frequency & "))-((EU" & care(i) + 4 & "*" & frequency & ")))+AE" & frequency & "),6)"
    Cells(care(i) + 5, 30).formula = "=Round(((((ET" & care(i) + 5 & "*" & frequency & "))-((EU" & care(i) + 5 & "*" & frequency & ")))+AE" & frequency & "),6)"
    Cells(care(i) + 4, 31).formula = "=Round(((EU" & care(i) + 4 & "*" & frequency & ")),6)"
    Cells(care(i) + 5, 31).formula = "=Round(((EU" & care(i) + 5 & "*" & frequency & ")),6)"
    Cells(care(i) + 4, 32).formula = "=Round(((EV" & care(i) + 4 & "*" & frequency & ")),6)"
    Cells(care(i) + 5, 32).formula = "=Round(((EV" & care(i) + 5 & "*" & frequency & ")),6)"
    Cells(care(i) + 4, 34).formula = "=((EW" & care(i) + 4 & "*" & frequency & "))"
    Cells(care(i) + 5, 34).formula = "=((EW" & care(i) + 5 & "*" & frequency & "))"
    Cells(care(i) + 4, 35).formula = "=((EX" & care(i) + 4 & "*" & frequency & "))"
    Cells(care(i) + 5, 35).formula = "=((EX" & care(i) + 5 & "*" & frequency & "))"
    Cells(care(i) + 4, 96).formula = "=(Round((Round((((ET" & care(i) + 4 & "*" & frequency & "))*AV" & care(i) + 4 & "*1" & "),2)*BB" & care(i) + 4 & "),2)+Round((Round(((AE" & care(i) + 4 & "-((EU" & care(i) + 4 & "*" & frequency & ")))*AV" & care(i) + 4 & "*1),2)*BS" & care(i) + 4 & "),2))"
    Cells(care(i) + 5, 96).formula = "=(Round((Round((((ET" & care(i) + 5 & "*" & frequency & "))*AV" & care(i) + 5 & "*1" & "),2)*BB" & care(i) + 5 & "),2)+Round((Round(((AE" & care(i) + 5 & "-((EU" & care(i) + 5 & "*" & frequency & ")))*AV" & care(i) + 5 & "*1),2)*BS" & care(i) + 5 & "),2))"
    For j = 108 To 111
        Cells(care(i) + 4, j).Value = ")*" & frequency
        Cells(care(i) + 5, j).Value = ")*" & frequency
    Next
    For j = 113 To 114
        Cells(care(i) + 4, j).Value = ")*" & frequency
        Cells(care(i) + 5, j).Value = ")*" & frequency
    Next
Next


End Sub
