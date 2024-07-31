Sub ListSheetNamesToSheet()
    Dim ws As Worksheet
    Dim outputSheet As Worksheet
    Dim sheetArray() As String
    Dim i As Integer
    Dim sheetCount As Integer

    sheetCount = ActiveWorkbook.Sheets.Count

    ReDim sheetArray(1 To sheetCount)

    i = 1
    For Each ws In ActiveWorkbook.Sheets
        sheetArray(i) = ws.Name
        i = i + 1
    Next ws

    Set outputSheet = ActiveWorkbook.Sheets.Add
    outputSheet.Name = "シート名一覧"
    outputSheet.Activate

    For i = LBound(sheetArray) To UBound(sheetArray)
        outputSheet.Cells(i, 1).Value = sheetArray(i)
    Next i
End Sub
