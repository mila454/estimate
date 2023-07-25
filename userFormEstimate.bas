Attribute VB_Name = "userFormEstimate"
Public typeEstimate As String

Sub userFormEstimate()
prepareEstimate.PlantingTextBox.Value = "Посадка *2024"
prepareEstimate.restorationTextBox.Value = "Воссановление *2024"
prepareEstimate.care1TextBox.Value = "Уход *2024"
prepareEstimate.care2TextBox.Value = "Уход *2025"
prepareEstimate.care3TextBox.Value = "Уход *2026"

prepareEstimate.Show

End Sub
