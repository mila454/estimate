VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} prepareEstimate 
   Caption         =   "Оформление сметы"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8280
   OleObjectBlob   =   "prepareEstimate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "prepareEstimate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub exitCommandButton_Click()
Unload prepareEstimate
End Sub

Private Sub SNOptionButton_Click()
typeEstimate = "СН"
End Sub

Private Sub UserForm_Initialize()
prepareEstimate.PlantingTextBox.Value = "Посадка *2024"
prepareEstimate.restorationTextBox.Value = "Воссановление *2024"
prepareEstimate.care1TextBox.Value = "Уход *2024"
prepareEstimate.care2TextBox.Value = "Уход *2025"
prepareEstimate.care3TextBox.Value = "Уход *2026"
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode = vbReyEscape Then
    Unload Me
End If

End Sub
