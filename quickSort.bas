Attribute VB_Name = "quickSort"
Option Explicit
Sub testSort()

Dim coll As New Collection
coll.Add 135
coll.Add 181
coll.Add 235
coll.Add 35
coll.Add 336

Call quickSort(coll, 1, coll.Count)

Dim item As Variant
For Each item In coll
    Debug.Print item
Next


End Sub


Sub quickSort(coll As Collection, first As Long, last As Long)

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
    ' Поменять значения
        temp = coll(low)
        coll.Add coll(high), After:=low
        coll.Remove low
        coll.Add temp, Before:=high
        coll.Remove high + 1
        ' Перейти к следующим позициям
        low = low + 1
        high = high - 1
    End If
    Loop
    If first < high Then quickSort coll, first, high
    If low < last Then quickSort coll, low, last
End Sub
        
