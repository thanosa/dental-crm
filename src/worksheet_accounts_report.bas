Private Sub sortPivotTable1(sortOrder As XlSortOrder, sortField As String, pivotLines As Integer)
    
    ' Worksheet dependent
    ptName = "PivotTable1"
    ptBaseField = "Name"
    
    ActiveSheet.PivotTables(ptName).PivotFields(ptBaseField).AutoSort sortOrder _
        , sortField, ActiveSheet.PivotTables(ptName).PivotColumnAxis. _
        pivotLines(pivotLines), 1
    
    Call scrollToTop

End Sub

Private Sub Worksheet_Activate()

    Call performancePre
    Call unprotectAllWs
    
    Call refreshPivotTableData
    
    Call protectAllWs(True)
    Call performancePost
    
End Sub

Private Sub EditTimelineButton_Click()

    Call performancePre
    
    Call protectAllWs(False)
    
    Call performancePost
    
End Sub

Private Sub LockTimelineButton_Click()

    Call performancePre
    
    Call protectAllWs(True)
    
    Call performancePost
    
End Sub

Private Sub SortByPatientIdButton_Click()
    
    Call performancePre
    Call unprotectAllWs
    
    Call sortPivotTable1(xlAscending, "ID", 1)
    
    Call protectAllWs(True)
    Call performancePost
    
End Sub

Private Sub SortByBalanceNegativeButton_Click()

    Call performancePre
    Call unprotectAllWs
    
    Call sortPivotTable1(xlAscending, "Balance.", 2)
    
    Call protectAllWs(True)
    Call performancePost

End Sub

Private Sub SortByBalancePositiveButton_Click()

    Call performancePre
    Call unprotectAllWs
    
    Call sortPivotTable1(xlDescending, "Balance.", 2)
    
    Call protectAllWs(True)
    Call performancePost
    
End Sub

Private Sub SortByCostButton_Click()

    Call performancePre
    Call unprotectAllWs
    
    Call sortPivotTable1(xlDescending, "Cost.", 3)
    
    Call protectAllWs(True)
    Call performancePost
    
End Sub

Private Sub SortByReceiptButton_Click()
    
    Call performancePre
    Call unprotectAllWs
    
    Call sortPivotTable1(xlDescending, "Receipts.", 4)
    
    Call protectAllWs(True)
    Call performancePost
    
End Sub

Private Sub SortByAppointmentsButton_Click()

    Call performancePre
    Call unprotectAllWs
    
    Call sortPivotTable1(xlDescending, "Appointments", 5)
    
    Call protectAllWs(True)
    Call performancePost
    
End Sub



