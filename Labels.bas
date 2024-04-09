'Version 1.0
'Creaded by Marshall

Sub Update()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim dataRange As Range
    Dim dataArray() As Variant
    Dim sheetName As String
    Dim i As Long, j As Long
    
    ' Set the worksheet where the data is located
    Set ws = ThisWorkbook.Worksheets("Data")
    
    ' Find the last row in column A (assuming data starts from A8)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Check if there is data in the range
    If lastRow >= 8 Then
        ' Read the data into a 2D array
        Set dataRange = ws.Range("A8:C" & lastRow)
        dataArray = dataRange.Value
        
        ' Clear the data range
        dataRange.ClearContents
        
        ' Fill in the list of sheet names in column A
        For i = 1 To ThisWorkbook.Sheets.Count
            ws.Cells(i, 1).Value = ThisWorkbook.Sheets(i).Name
        Next i
        
        ' Loop through the dataArray and copy data to corresponding sheets
        For i = LBound(dataArray, 1) To UBound(dataArray, 1)
            sheetName = dataArray(i, 1)
            For j = 1 To ThisWorkbook.Sheets.Count
                If ThisWorkbook.Sheets(j).Name = sheetName Then
                    ' Copy data to the corresponding sheet
                    ThisWorkbook.Sheets(j).Cells(ThisWorkbook.Sheets(j).Rows.Count, "A").End(xlUp).Offset(1, 0).Resize(1, UBound(dataArray, 2)).Value = _
                        Application.Index(dataArray, i, 0)
                    Exit For
                End If
            Next j
        Next i
        
        MsgBox "Data processed and populated into sheets.", vbInformation
    Else
        MsgBox "No data found in the specified range.", vbExclamation
    End If
End Sub
