Private Sub Worksheet_Activate()

    Call performancePre
    Call unprotectAllWs
    
    Call refreshPivotTableData
    Call scrollToTop
    
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

