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
Option Explicit
Public typeEstimate As String


Private Sub cancelCommandButton_Click()
Dim element As Variant

For Each element In prepareEstimate.Controls
    If TypeOf element Is MSForms.CheckBox Then
        element.Value = False
    End If
    If TypeOf element Is MSForms.TextBox Then
        element.Value = ""
    End If
Next


End Sub

Private Sub estimateCommandButton_Click()
prepareEstimate.Hide
If NDSCheckBox.Value = True Then
    Call userFormEstimate.nds(typeEstimate)
End If

End Sub

Private Sub exitCommandButton_Click()
Unload prepareEstimate
End Sub


Private Sub SNOptionButton_Click()

typeEstimate = "СН"

End Sub

Private Sub TSNOptionButton_Click()

typeEstimate = "ТСН"

End Sub

Private Sub UserForm_Initialize()

If SNOptionButton.Value = True Then
    typeEstimate = "СН"
Else
    typeEstimate = "ТСН"
End If

End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode = vbReyEscape Then
    Unload Me
End If

End Sub
