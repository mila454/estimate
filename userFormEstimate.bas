Attribute VB_Name = "userFormEstimate"
Public typeEstimate As String

Sub userFormEstimate()
prepareEstimate.PlantingTextBox.Value = "������� *2024"
prepareEstimate.restorationTextBox.Value = "������������� *2024"
prepareEstimate.care1TextBox.Value = "���� *2024"
prepareEstimate.care2TextBox.Value = "���� *2025"
prepareEstimate.care3TextBox.Value = "���� *2026"

prepareEstimate.Show

End Sub
