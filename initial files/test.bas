Attribute VB_Name = "test"
Sub test()
Dim nameLocation As String
Dim nameLocation1() As String
Dim smetaName As String
Dim smetaName1() As String

Dim i As Variant

'Debug.Print Cells(20, 6).HasFormula
Sheets("Source").Activate
For i = 1 To 538
    If Cells(i, 6).HasFormula = False And Cells(i, 6).Value = "Новая локальная смета" Then
        smetaName = smetaName & ";" & Cells(i, 7).Value
    End If
Next
smetaName1 = Split(smetaName, ";")
Debug.Print smetaName
For i = LBound(smetaName1) To UBound(smetaName1)
    Debug.Print smetaName1(i)
Next
Sheets("Смета СН-2012 по гл. 1-5").Activate


For Each item In Range("A1:K319")
    If item Like "*ЛОКАЛЬНАЯ СМЕТА №*" Then
        nameLocation = nameLocation & " " & item.Row
        
    End If
Next
nameLocation1 = Split(nameLocation, " ")
'Cells(248, 11).formula = a
'Cells(248, 11).formula = Cells(248, 9).formula
'Debug.Print Range("F20").Font.color
'Debug.Print TypeName(Format(Date, "yyyy"))

'Debug.Print Cells(20, 6).Font.color
 Debug.Print nameLocation
 For i = LBound(nameLocation1) To UBound(nameLocation1)
    Debug.Print nameLocation1(i)
Next


End Sub
