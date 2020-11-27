Public Sub scrollToTop()
    
    ActiveWindow.ScrollRow = 1
    
End Sub

Public Sub refreshPivotTableData()
    
    Dim ws As Worksheet
    Dim pt As PivotTable
    
    For Each ws In ActiveWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.PivotCache.Refresh
        Next pt
    Next ws
    
End Sub

Public Sub protectAllWs(drawingObjects As Boolean)

    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets
        ws.Protect drawingObjects:=drawingObjects, Contents:=True, Scenarios:=True, _
        AllowFiltering:=True, AllowDeletingRows:=True, AllowFormattingRows:=True, _
        AllowFormattingColumns:=True, AllowFormattingCells:=True, AllowUsingPivotTables:=True
    Next ws
    
End Sub

Public Sub unprotectAllWs()

    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets
        ws.Unprotect
    Next ws
    
End Sub

