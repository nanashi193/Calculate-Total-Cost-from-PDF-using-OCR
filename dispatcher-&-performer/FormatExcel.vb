Sub FormatExcel()

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets("Sheet1")
    
    Dim i As Long
    For i = ws.UsedRange.Rows.Count To 1 Step -1
        If ws.Cells(i, 1).Value = "" Then
            ws.Rows(i).Delete
        End If
    Next i
    
    ' Column widths
    ws.Columns("A:A").ColumnWidth = 21.47
    ws.Columns("B:B").ColumnWidth = 15.13
    ws.Columns("C:C").ColumnWidth = 19.2
    ws.Columns("D:D").ColumnWidth = 20.33
    ws.Columns("E:E").ColumnWidth = 26.33
    ws.Columns("F:F").ColumnWidth = 28.33
    ws.Columns("G:G").ColumnWidth = 14.07
    
    ws.Rows("1:1").RowHeight = 39
    
    With ws.Range("A1:G1")
        .Interior.Color = 65535
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ws.Columns("C:C").HorizontalAlignment = xlCenter
    
    ActiveWorkbook.Save
    
End Sub