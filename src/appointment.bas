
Private Const wsName = "Appointments"

Private Sub searchAppointment()

    Unprotect
    
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(wsName)
    
    criteriaEntry = ws.Range("AppointmentsCriteria").Value
    
    Call unfilter
    
    If criteriaEntry = "" Then
        Call unfilter
    Else
        criteria = "=" & ws.Range("AppointmentsCriteria")
        ws.Range("AppointmentsRecords").AutoFilter Field:=1, Criteria1:=criteria, Operator:=xlAnd
    End If
    
    Call scrollToTop
    
    Call protectSheet
    
End Sub

Private Sub clearSearch()

    Unprotect
    
    Dim ws As Worksheet
    Set ws = Worksheets(wsName)
    
    Range("AppointmentsCriteria").Value = ""
    
    Call protectSheet

End Sub

Private Sub addAppointment()
    
    ' Layout dependent
    headerRow = 6
    maxRow = 1000000000
    idColumn = "A"
    selectColumn = "D"
    clearFieldsFromColumn = "D"
    clearFieldsToColumn = "L"
    currentDateColumn = "C"
    
    Dim ws As Worksheet
    Set ws = Worksheets(wsName)
    
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
    response = MsgBox("Add appointment for patient with ID:  " & searchPatientId & " ?", vbQuestion + vbYesNo, "New appointment")
    If response = vbNo Then
        protectSheet
        Exit Sub
    End If
        
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
    
    Call unfilter
    
    ' Copys the last record to the first empty rowc to maintain the formating and data validations
    Range(copyRow & ":" & copyRow).Copy
    Range(emptyRow & ":" & emptyRow).Select
    ws.Paste
    Application.CutCopyMode = False
    
    ' Put next id
    Range(idColumn & emptyRow).Value = searchPatientId
    
    ' Clear new entry fields
    Range(clearFieldsFromColumn & emptyRow & ":" & clearFieldsToColumn & emptyRow).ClearContents
    
    ' Enter the current date
    Range(currentDateColumn & emptyRow).Value = Date
    
    ' Filter the records
    Call searchAppointment
    
    ' Select the first field to enter
    Range(selectColumn & emptyRow).Select
    
    Call protectSheet
        
    
End Sub

Private Sub selectSearch()

    Range("AppointmentsCriteria").Select

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

    Call searchAppointment
    Call selectSearch

End Sub

Private Sub ClearButton_Click()

    Call clearSearch
    Call searchAppointment
    Call selectSearch

End Sub

Private Sub AddAppointmentButton_Click()

    Call addAppointment

End Sub
