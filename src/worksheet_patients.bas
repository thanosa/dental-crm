' Layout dependent
Private Const headerRow = 6
Private Const maxRow = 10000000
Private Const idColumn = "A"
Private Const idColumnIdx = 1
Private Const selectColumn = "B"
Private Const selectColumnIdx = 2
Private Const practiceColumn = "D"
Private Const practiceColumnIdx = 4

Private Const clearFieldsFromColumn = "B"
Private Const clearFieldsToColumn = "L"
    

Private Sub searchPatient()

    Unprotect
    
    Dim ws As Worksheet
    Set ws = Worksheets(patientsWsName)
    
    Call unfilter
    
    practiceCriteriaEntry = ws.Range("PatientsPractice").Value
    If practiceCriteriaEntry <> "" Then
        practiceCriteria = "=" & practiceCriteriaEntry
        ws.Range("PatientsRecords").AutoFilter Field:=practiceColumnIdx, Criteria1:=practiceCriteria, Operator:=xlAnd
    End If
    
    patientCriteriaEntry = ws.Range("PatientsCriteria").Value
    If patientCriteriaEntry <> "" Then
        If IsNumeric(patientCriteriaEntry) Then
            patientCriteria = "=" & patientCriteriaEntry
            ws.Range("PatientsRecords").AutoFilter Field:=idColumnIdx, Criteria1:=patientCriteria, Operator:=xlAnd
        Else
            patientCriteria = "=*" & patientCriteriaEntry & "*"
            ws.Range("PatientsRecords").AutoFilter Field:=selectColumnIdx, Criteria1:=patientCriteria, Operator:=xlAnd
        End If
    End If
    
    Call scrollToTop
    
    Call protectSheet

End Sub


Private Sub filterPractice()

    Unprotect
    
    Dim ws As Worksheet
    Set ws = Worksheets(patientsWsName)
    
    Call unfilter
    
    practiceCriteriaEntry = ws.Range("PatientsPractice").Value
    If practiceCriteriaEntry <> "" Then
        practiceCriteria = "=" & practiceCriteriaEntry
        ws.Range("PatientsRecords").AutoFilter Field:=practiceColumnIdx, Criteria1:=practiceCriteria, Operator:=xlAnd
    End If
    
    Call protectSheet

End Sub


Private Sub clearSearch()

    Unprotect
    
    Dim ws As Worksheet
    Set ws = Worksheets(patientsWsName)
    
    Range("PatientsCriteria").ClearContents
    
    Call protectSheet

End Sub

Private Sub addPatient()
    
    Dim ws As Worksheet
    Set ws = Worksheets(patientsWsName)
    
    ' Check that the practice has been selected
    If ws.Range("PatientsPractice").Value = "" Then
        Call MsgBox("Please fill Practice before adding a patient")
        ws.Range("PatientsPractice").Select
        Exit Sub
    End If
    
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
    
    ' Set the branch
    branchName = ws.Range("PatientsPractice").Value
    If branchName <> "" Then
        Range(practiceColumn & emptyRow).Value = branchName
    End If
    
    ' Select the next field
    Range(selectColumn & emptyRow).Select
    
    Call protectSheet
    

End Sub

Private Sub selectSearch()

    Range("PatientsCriteria").Select

End Sub

Private Sub unfilter()

    Dim ws As Worksheet
    Set ws = Worksheets(patientsWsName)
    
    On Error Resume Next
        ws.showAllData
    On Error GoTo 0
    
End Sub

Private Sub protectSheet()

    Protect drawingObjects:=True, Contents:=True, Scenarios:=True, _
    AllowFiltering:=True, AllowDeletingRows:=True, AllowFormattingRows:=True, _
    AllowFormattingColumns:=True, AllowFormattingCells:=True
    
End Sub

Private Sub Worksheet_Activate()

    Call performancePre
    Call unprotectAllWs
    
    Call refreshPivotTableData
    
    Call protectAllWs(True)
    Call performancePost
    
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
    Call filterPractice
    
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

    If Target.Address = Range("PatientsPractice").Address Then
        Call clearSearch
        Call filterPractice
    End If
    
End Sub






