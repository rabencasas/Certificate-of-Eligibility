VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPrintCopy 
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5580
   OleObjectBlob   =   "frmPrintCopy.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPrintCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    CopyData
    GenerateId (True)
    Sheet1.PrintOut from:=1, to:=1, copies:=TextBox1.Value
    
    ' Show successfull print message to user
    Sheet4.Range("E24").Value = "Certifiation to " & UCase(Sheet4.Range("H10").Value) & " is successfully printed."
    
    Me.Hide
End Sub

Private Sub UserForm_Activate()
    TextBox1.SetFocus

End Sub

Private Sub UserForm_Click()

End Sub
