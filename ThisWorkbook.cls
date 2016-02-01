
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call cleanTable
End Sub

Private Sub Workbook_Open()
    Sheet1.Activate
    Call cleanTable
    Call newGame
    
End Sub
