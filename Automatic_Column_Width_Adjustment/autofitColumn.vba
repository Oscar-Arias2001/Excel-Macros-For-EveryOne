Private Sub workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    Application.ScreenUpdating = False
    Sh.Cells.EntireColumn.AutoFit
    Application.ScreenUpdating = True
End Sub