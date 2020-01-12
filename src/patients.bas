
Private Const wsName = "Patients"

Private Sub searchPatient()

    Unprotect
    
    Dim ws As Worksheet
    Set ws = Worksheets(wsName)
    
    criteriaEntry = ws.Range("PatientsCriteria").Value
    
    Call unfilter
    
    If criteriaEntry <> "" Then
        If IsNumeric(criteriaEntry) Then
            criteria = "=" & criteriaEntry
            ws.Range("PatientsRecords").AutoFilter Field:=1, Criteria1:=criteria, Operator:=xlAnd
        Else
            criteria = "=*" & criteriaEntry & "*"
            ws.Range("PatientsRecords").AutoFilter Field:=2, Criteria1:=criteria, Operator:=xlAnd
        End If
    End If
    
    Call scrollToTop
    
    Call protectSheet

End Sub

Private Sub clearSearch()

    Unprotect
    
    Dim ws As Worksheet
    Set ws = Worksheets(wsName)
    
    Range("PatientsCriteria").ClearContents
    
    Call protectSheet

End Sub

Private Sub addPatient()
    
    ' Layout dependent
    headerRow = 6
    maxRow = 10000000
    idColumn = "A"
    selectColumn = "B"
    clearFieldsFromColumn = "B"
    clearFieldsToColumn = "K"
    
    Dim ws As Worksheet
    Set ws = Worksheets(wsName)
    
    Unprotect
    
    ' Finds the first empty row and the last populated one
    emptyRow = -1
    copyRow = -1
    
    For currentRow = headerRow To maxRow
        currentId = Range(idColumn & currentRow).Value
        If currentId = "" Then
            emptyRow = currentRow
            copyRow = emptyRow - 1
            Exit For
        End If
    Next currentRow
    
    ' Copys the last record to the first empty rowc to maintain the formating and data validation
    Range(copyRow & ":" & copyRow).Copy
    Range(emptyRow & ":" & emptyRow).Select
    ws.Paste
    Application.CutCopyMode = False
    
    ' Put next id
    Range(idColumn & emptyRow).Value = Range(idColumn & copyRow) + 1
    
    ' Clear fields
    Range(clearFieldsFromColumn & emptyRow & ":" & clearFieldsToColumn & emptyRow).ClearContents
    
    ' Select the next field
    Range(selectColumn & emptyRow).Select
    
    Call protectSheet
    

End Sub

Private Sub selectSearch()

    Range("PatientsCriteria").Select

End Sub

Private Sub scrollToTop()
    
    ActiveWindow.ScrollRow = 1
    
End Sub

Private Sub unfilter()

    Dim ws As Worksheet
    Set ws = Worksheets(wsName)
    
    On Error Resume Next
        ws.showAllData
    On Error GoTo 0
    
End Sub

Private Sub protectSheet()

    Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, _
    AllowFiltering:=True, AllowDeletingRows:=True, AllowFormattingRows:=True, _
    AllowFormattingColumns:=True, AllowFormattingCells:=True
    
End Sub

Private Sub SearchButton_Click()

    Call searchPatient
    Call selectSearch
    
End Sub

Private Sub ClearButton_Click()

    Call clearSearch
    Call searchPatient
    Call selectSearch

End Sub

Private Sub AddPatientButton_Click()

    Call clearSearch
    Call addPatient
    
End Sub
