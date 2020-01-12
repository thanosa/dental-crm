Private Sub SortByPatientIdButton_Click()

    ActiveSheet.PivotTables("PivotTable1").PivotFields("Name").AutoSort xlAscending _
        , "Patient  ID", ActiveSheet.PivotTables("PivotTable1").PivotColumnAxis. _
        PivotLines(1), 1

End Sub

Private Sub SortByAppointmentsButton_Click()

    ActiveSheet.PivotTables("PivotTable1").PivotFields("Name").AutoSort _
        xlDescending, "Appointments", ActiveSheet.PivotTables("PivotTable1"). _
        PivotColumnAxis.PivotLines(2), 1

End Sub

Private Sub SortByBalanceButton_Click()

    ActiveSheet.PivotTables("PivotTable1").PivotFields("Name").AutoSort _
        xlDescending, "Sum of Balance", ActiveSheet.PivotTables("PivotTable1"). _
        PivotColumnAxis.PivotLines(3), 1

End Sub

Private Sub SortByCostButton_Click()

    ActiveSheet.PivotTables("PivotTable1").PivotFields("Name").AutoSort _
        xlDescending, "Sum of Cost", ActiveSheet.PivotTables("PivotTable1"). _
        PivotColumnAxis.PivotLines(4), 1

End Sub

Private Sub SortByReceiptButton_Click()

    ActiveSheet.PivotTables("PivotTable1").PivotFields("Name").AutoSort _
        xlDescending, "Sum of Receipt", ActiveSheet.PivotTables("PivotTable1"). _
        PivotColumnAxis.PivotLines(5), 1

End Sub
