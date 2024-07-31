Sub FormatDocument()
    Dim ws As Worksheet

    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = True Then
            ws.Activate
            ws.Cells(1, 1).Select
            ActiveWindow.Zoom = 100
        End If
    Next ws

    If Worksheets(1).Visible = True Then Worksheets(1).Activate
End Sub
