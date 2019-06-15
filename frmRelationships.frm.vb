VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRelationships 
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8850
   OleObjectBlob   =   "frmRelationships.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRelationships"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Sheet4.Range("H22").Value = ListBox1.Value
    Me.Hide
End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Sheet4.Range("H22").Value = ListBox1.Value
    Me.Hide
End Sub

Private Sub UserForm_Activate()
    ListBox1.List = Sheet4.Range("M33:M49").Value
End Sub

Private Sub UserForm_Click()
    
End Sub
