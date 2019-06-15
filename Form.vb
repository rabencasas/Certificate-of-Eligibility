VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub cmdDateToday_Click()
    Sheet4.Range("H8").Value = Format(DateTime.Now, "mm/dd/yyyy")
    Sheet4.Range("H8").Activate
End Sub

Private Sub cmdOpenDir_Click()
    Call Shell("explorer.exe" & " " & Sheet5.Range("H7").Value, vbNormalFocus)
End Sub

Private Sub cmdPreview_Click()
    OpenSheet
    
    CopyData
    GenerateId (False)
    Sheet1.PrintPreview
    
    CloseSheet
End Sub

Private Sub cmdPrint_Click()
    OpenSheet
    
    frmPrintCopy.Show
    
    CloseSheet
End Sub

Private Sub cmdSave_Click()
    OpenSheet
    
    SaveCertification
    
    CloseSheet
End Sub

Private Sub cmdSetAsClient_Click()
    Sheet4.Range("H20").Value = "CLIENT"
    Sheet4.Range("H22").Value = ""
End Sub

Private Sub CommandButton1_Click()
    OpenSheet
    frmAddress.Show
    CloseSheet
    Sheet4.Range("H14").Activate
End Sub

Private Sub CommandButton2_Click()
    frmAssistanceType.Show
End Sub

Private Sub CommandButton3_Click()
    frmRelationships.Show
    Sheet4.Range("H22").Activate
End Sub

Private Sub CommandButton4_Click()
    OpenSheet
    frmSetClerk.Show
    CloseSheet
End Sub

Private Sub CommandButton5_Click()
    OpenSheet
    frmSetDate.Show
    CloseSheet
End Sub

Private Sub CommandButton6_Click()
    OpenSheet
    frmAssistanceType.Show
    CloseSheet
     Sheet4.Range("H18").Activate
End Sub

Private Sub Worksheet_Activate()
    OpenSheet
    Sheet4.Range("H8").Value = "asdfasdf"
    CloseSheet
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    
End Sub
