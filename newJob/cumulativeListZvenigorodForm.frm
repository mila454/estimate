VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cumulativeListZvenigorodForm 
   Caption         =   "Накопительная"
   ClientHeight    =   4095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8565.001
   OleObjectBlob   =   "cumulativeListZvenigorodForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cumulativeListZvenigorodForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCumulativeRoss_Click()
cumulativeListRossolimo.cumulativeListRossolimo
Unload Me
End Sub

Private Sub btnCumulativeZve_Click()
cumulativeListZvenigorod.cumulativeListZvenigorod
Unload Me
End Sub

Private Sub btnTotalCostRoss_Click()
TotalCostRoss.TotalCostRoss
Unload Me
End Sub

Private Sub btnTotalCostZve_Click()
TotalCostZve.TotalCostZve
Unload Me
End Sub


