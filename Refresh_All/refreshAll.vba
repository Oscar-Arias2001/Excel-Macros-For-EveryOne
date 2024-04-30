Sub RefreshAllWorksheets()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        ActiveWorkbook.RefreshAll
    Next ws
End Sub