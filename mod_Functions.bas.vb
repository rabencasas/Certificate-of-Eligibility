Attribute VB_Name = "mod_Functions"
Public Function GenerateId(Yes As Boolean)
    If Yes Then
        Sheet1.Range("A51").Value = "[" & Format(DateTime.Now, "yy.mm.dd") & "][" & Format(DateTime.Now, "hh.mm.ss") & "]"
    Else
        Sheet1.Range("A51").Value = "ID not visible for preview"
    End If
End Function

Public Function CopyData()
    'name
    Sheet1.Range("F7").Value = UCase(Sheet4.Range("H10").Value)
    'age
    Sheet1.Range("L8").Value = Sheet4.Range("H12").Value
    'address
    Sheet1.Range("B9").Value = UCase((Sheet4.Range("H14").Value))
    
    'assistance type
    Sheet1.txtMedical.Value = ""
    Sheet1.txtBurial.Value = ""
    Sheet1.txtTransportation.Value = ""
    Sheet1.txtFood.Value = ""
    Sheet1.txtEducational.Value = ""
    
    Select Case Sheet4.Range("H16").Value
    Case "MEDICAL ASSISTANCE"
        Sheet1.txtMedical.Value = "X"
    Case "BURIAL ASSISTANCE"
        Sheet1.txtBurial.Value = "X"
    Case "TRANSPORTATION ASSISTANCE"
        Sheet1.txtTransportation.Value = "X"
    Case "FOOD ASSISTANCE"
        Sheet1.txtFood.Value = "X"
    Case "EDUCATIONAL ASSISTANCE"
        Sheet1.txtEducational.Value = "X"
    Case Else
        Sheet1.txtMedical.Value = ""
        Sheet1.txtBurial.Value = ""
        Sheet1.txtTransportation.Value = ""
        Sheet1.txtFood.Value = ""
        Sheet1.txtEducational.Value = ""
    End Select
    
    'amount in words
    Sheet1.Range("D14").Value = UCase(SpellNumber(Sheet4.Range("H18").Value))
    'amount in figure
    Sheet1.Range("I15").Value = Format(Sheet4.Range("H18").Value, "Standard")
    'name2
    Sheet1.Range("J18").Value = UCase(Sheet4.Range("H10").Value)
    'relationship
    Sheet1.Range("J21").Value = UCase(Sheet4.Range("H22").Value)
    'beneficiary
    Sheet1.Range("J24").Value = Sheet4.Range("H20").Value
    'date issued: complete
    Sheet1.Range("C30").Value = Format(Sheet4.Range("H8").Value, "mmmm dd, yyyy")
    
    'date issued: day
    Sheet1.Range("B33").Value = Day(CDate(Sheet4.Range("H8").Value))
    
    If Right(Day(CDate(Sheet4.Range("H8").Value)), 1) = 1 Or Right(Day(CDate(Sheet4.Range("H8").Value)), 1) = 2 Or Right(Day(CDate(Sheet4.Range("H8").Value)), 1) = 3 Then
        If Right(Day(CDate(Sheet4.Range("H8").Value)), 1) = 1 And Day(CDate(Sheet4.Range("H8").Value)) <> 11 Then
            Sheet1.Range("B33").Value = Sheet1.Range("B33").Value & "st"
        ElseIf Right(Day(CDate(Sheet4.Range("H8").Value)), 1) = 2 And Day(CDate(Sheet4.Range("H8").Value)) <> 12 Then
            Sheet1.Range("B33").Value = Sheet1.Range("B33").Value & "nd"
        ElseIf Right(Day(CDate(Sheet4.Range("H8").Value)), 1) = 3 And Day(CDate(Sheet4.Range("H8").Value)) <> 13 Then
            Sheet1.Range("B33").Value = Sheet1.Range("B33").Value & "rd"
        Else
            Sheet1.Range("B33").Value = Sheet1.Range("B33").Value & "th"
        End If
    Else
        Sheet1.Range("B33").Value = Sheet1.Range("B33").Value & "th"
    End If
    
    'date issued: month & year
    Sheet1.Range("D33").Value = Format(Sheet4.Range("H8").Value, "mmmm, yyyy")
    
    'clerk
    Sheet1.Range("A47").Value = "Prepared by: " & Sheet4.Range("H5").Value
End Function

Public Function OpenSheet()
    Sheet4.Unprotect ("admin.pass")
    Sheet1.Unprotect ("admin.pass")
End Function

Public Function CloseSheet()
    Sheet4.Protect ("admin.pass")
    Sheet1.Protect ("admin.pass")
End Function

Public Function SaveCertification()
    If Sheet5.Range("H7").Value <> "" Then
        If Sheet1.Range("A51").Value <> "ID not visible for preview" Then
            Dim file As String
            Dim textfile As Integer
            
            file = Sheet5.Range("H7").Value & "\" & Sheet1.Range("A51").Value & " - " & UCase(Sheet4.Range("H10").Value) & ".certificate"
            
            textfile = FreeFile
            
            Open file For Output As textfile
            
            Print #textfile, Sheet1.Range("A51").Value & " - " & UCase(Sheet4.Range("H10").Value) & vbNewLine
            
            Print #textfile, "Name:" & vbTab & vbTab & vbTab & UCase(Sheet4.Range("H10").Value)
            Print #textfile, "Age:" & vbTab & vbTab & vbTab & Sheet4.Range("H12").Value
            Print #textfile, "Address:" & vbTab & vbTab & UCase(Sheet4.Range("H14").Value)
            Print #textfile, "Assistance type:" & vbTab & UCase(Sheet4.Range("H16").Value)
            Print #textfile, "Amount:" & vbTab & vbTab & vbTab & Format(Sheet4.Range("H18").Value, "Standard")
            Print #textfile, "Beneficiary:" & vbTab & vbTab & Sheet4.Range("H20").Value
            Print #textfile, "Relationship to Beneficiary:" & vbTab & UCase(Sheet4.Range("H22").Value)
            Print #textfile, "Date Issued:" & vbTab & vbTab & Format(Sheet4.Range("H8").Value, "mmmm dd, yyyy") & "__" & WeekdayName(Weekday(Format(Sheet4.Range("H8").Value, "mmmm dd, yyyy")))
            Print #textfile, "Clerk:" & vbTab & vbTab & vbTab & Sheet4.Range("H5").Value
            
            Close textfile
            
            ' Show successfull log message to user
            Sheet4.Range("E24").Value = "Certifiation to " & UCase(Sheet4.Range("H10").Value) & " is successfully saved."
        Else
            MsgBox "Cannot continue to save. The certification has not been printed yet.", vbCritical
        End If
        
    Else
        MsgBox "Please set the log folder path in order to save."
    End If
End Function
