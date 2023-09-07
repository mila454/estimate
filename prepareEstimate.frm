VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} prepareEstimate 
   Caption         =   "ќформление сметы"
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



Private Sub care1TextBox_Change()

End Sub

Private Sub care2TextBox_Change()

End Sub

Private Sub clearCommandButton_Click()
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
Dim ctr As Control
Dim i As Variant
i = 0
prepareEstimate.Hide

For Each ctr In prepareEstimate.simpleFrame.Controls
    If ctr.Value = True Then
        simpleFrameList(i) = ctr.Name
    End If
    i = i + 1
Next
i = 0
For Each ctr In prepareEstimate.complexFrame.Controls
    If ctr.Value = True Then
        complexFrameList(i) = ctr.Name
    End If
    i = i + 1
Next
i = 0
For Each ctr In prepareEstimate.executionFrame.Controls
    If ctr.Value = True Then
        executionFrameList(i) = ctr.Name
    End If
    i = i + 1
Next
i = 0
For Each ctr In prepareEstimate.typesOfWorksFrame.Controls
    If TypeOf ctr Is MSForms.Label Then
        GoTo continue
    ElseIf TypeOf ctr Is MSForms.TextBox And ctr.text <> "" Then
        typesOfWorksFrameList(i) = ctr.text
    End If
    i = i + 1
continue: Next
'If NDSOptionButton.Value = True Then
'    Call userFormEstimate.nds
'End If

'If financeCheckBox.Value = True Then
'    Call coefBudgetFinancing
'End If

End Sub

Private Sub exitCommandButton_Click()
Unload prepareEstimate
End Sub


Private Sub OptionButton1_Click()

End Sub

Private Sub simplePrepareFrame_Click()

End Sub

Private Sub plantingLabel_Click()

End Sub

Private Sub UserForm_Initialize()



End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode = vbReyEscape Then
    Unload Me
End If

End Sub
