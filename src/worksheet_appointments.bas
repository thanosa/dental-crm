' Layout dependent
Private Const headerRow = 7
Private Const maxRow = 1000000000
Private Const patientIdColumn = "A"
Private Const patientIdColumnIdx = 1
Private Const branchIdColumn = "C"
Private Const branchIdColumnIdx = "3"
Private Const dateColumn = "D"
Private Const complaintColumn = "G"
Private Const transactionNotesColumn = "L"
Private Const costColumn = "N"
Private Const receiptColumn = "O"


Private Sub searchAppointment()

    Unprotect
    
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(appointmentsWsName)
    
    Dim patientsWs As Worksheet
    Set patientsWs = ActiveWorkbook.Worksheets(patientsWsName)
    
    patientIdEntry = ws.Range("AppointmentsCriteria").Value
    
    Call unfilter
    
    branch = patientsWs.Range("PatientsPractice")
    If branch <> "" Then
        branchCriteria = "=" & branch
        ws.Range("AppointmentsRecords").AutoFilter Field:=branchIdColumnIdx, Criteria1:=branchCriteria, Operator:=xlAnd
    End If
    
    If patientIdEntry <> "" Then
        patientIdCriteria = "=" & patientIdEntry
        ws.Range("AppointmentsRecords").AutoFilter Field:=patientIdColumnIdx, Criteria1:=patientIdCriteria, Operator:=xlAnd
    End If
    
    Call scrollToTop
    
    Call protectSheet
    
End Sub


Private Sub filterPractice()

    Unprotect
    
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(appointmentsWsName)
    
    Dim patientsWs As Worksheet
    Set patientsWs = ActiveWorkbook.Worksheets(patientsWsName)
    
    patientIdEntry = ws.Range("AppointmentsCriteria").Value
    
    Call unfilter
    
    branch = patientsWs.Range("PatientsPractice")
    If branch <> "" Then
        branchCriteria = "=" & branch
        ws.Range("AppointmentsRecords").AutoFilter Field:=branchIdColumnIdx, Criteria1:=branchCriteria, Operator:=xlAnd
    End If
    
    Call scrollToTop
    
    Call protectSheet
    
End Sub


Private Sub addAppointment()
    
    Dim ws As Worksheet
    Set ws = Worksheets(appointmentsWsName)
    
    Unprotect
    
    ' Check if the user has entered patient id in the search field
    searchPatientId = Range("AppointmentsCriteria").Value
    If searchPatientId = "" Then
        Call MsgBox("Enter 'Patient ID' in search field", vbExclamation)
        Range("AppointmentsCriteria").Select
        protectSheet
        Exit Sub
    End If
    
    ' Asks the user if they are sure about creating the new appointment for that patient
    searchPatientName = Range("AppointmentsPatientName").Value
    response = MsgBox("Add appointment for: " & searchPatientId & " ?", vbQuestion + vbYesNo, "New appointment")
    If response = vbNo Then
        protectSheet
        Exit Sub
    End If
        
    ' Finds the first empty row and the last populated one
    emptyRow = -1
    copyRow = -1
    
    For currentRow = headerRow To maxRow
        currentId = Range(patientIdColumn & currentRow).Value
        If currentId = "" Then
            emptyRow = currentRow
            copyRow = emptyRow - 1
            Exit For
        End If
    Next currentRow
    
    Call unfilter
    
    ' Copys the last record to the first empty row to maintain the formating and data validations
    Range(copyRow & ":" & copyRow).Copy
    Range(emptyRow & ":" & emptyRow).Select
    ws.Paste
    Application.CutCopyMode = False
    
    ' Fills the patient id
    Range(patientIdColumn & emptyRow).Value = searchPatientId
    
    ' Clear new entry fields
    clearFieldsFromColumn1 = complaintColumn
    clearFieldsToColumn1 = transactionNotesColumn
    clearFieldsFromColumn2 = costColumn
    clearFieldsToColumn2 = receiptColumn
    
    Range(clearFieldsFromColumn1 & emptyRow & ":" & clearFieldsToColumn1 & emptyRow).ClearContents
    Range(clearFieldsFromColumn2 & emptyRow & ":" & clearFieldsToColumn2 & emptyRow).ClearContents
    
    ' Enter the current date
    Range(dateColumn & emptyRow).Value = Date
    
    ' Filter the records
    Call searchAppointment
    
    ' Select the first field to enter
    selectColumn = dateColumn
    Range(selectColumn & emptyRow).Select
        
    
End Sub

Private Sub selectSearch()

    Range("AppointmentsCriteria").Select

End Sub

Private Sub unfilter()
    
    Dim ws As Worksheet
    Set ws = Worksheets(appointmentsWsName)
    
    On Error Resume Next
        ws.showAllData
    On Error GoTo 0

End Sub

Private Sub clearSearch()

    Unprotect
    
    Dim ws As Worksheet
    Set ws = Worksheets(appointmentsWsName)
    
    Range("AppointmentsCriteria").Value = ""
    
    Call protectSheet

End Sub

Private Sub protectSheet()

    Protect drawingObjects:=True, Contents:=True, Scenarios:=True, _
    AllowFiltering:=True, AllowDeletingRows:=True, AllowFormattingRows:=True, _
    AllowFormattingColumns:=True, AllowFormattingCells:=True
    
End Sub

Private Sub Worksheet_Activate()

    Call performancePre
    Call unprotectAllWs
    
    If practiceUpdated = True Then
        Call clearSearch
        practiceUpdated = False
    End If
    
    Call refreshPivotTableData
    Call filterPractice
    Call selectSearch
    
    Call protectAllWs(True)
    Call performancePost
    
End Sub

Private Sub SearchButton_Click()

    Call performancePre
    Call searchAppointment
    Call selectSearch
    Call performancePost
    
End Sub

Private Sub ClearButton_Click()

    Call performancePre
    Call clearSearch
    Call searchAppointment
    Call selectSearch
    Call performancePost
    
End Sub

Private Sub AddAppointmentButton_Click()

    Call performancePre
    Call addAppointment
    Call performancePost
    
End Sub




