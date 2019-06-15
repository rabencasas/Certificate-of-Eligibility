VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAssistanceType 
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8400
   OleObjectBlob   =   "frmAssistanceType.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAssistanceType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Sheet4.Range("H16").Value = ListBox1.Value
    Me.Hide
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Sheet4.Range("H16").Value = ListBox1.Value
    Me.Hide
End Sub

Private Sub UserForm_Activate()
    ListBox1.List = Sheet4.Range("Q33:Q38").Value
End Sub
