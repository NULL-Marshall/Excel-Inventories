'Version 1.3
'Creaded by Marshall

Sub Update()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim dataArray() As Variant
    Dim searchRange As Range
    Dim foundCell As Range
    Dim i As Long

    Application.ScreenUpdating = False

    Set ws = ThisWorkbook.Sheets("Data")
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    ReDim dataArray(1 To ThisWorkbook.Sheets.count, 1 To 3)
    Set searchRange = ws.Range("A8:A" & lastRow)
    
    For i = 1 To ThisWorkbook.Sheets.count
        dataArray(i, 1) = ThisWorkbook.Sheets(i).name
        Set foundCell = searchRange.Find(What:=ThisWorkbook.Sheets(i).name, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not foundCell Is Nothing Then
            dataArray(i, 2) = foundCell.Offset(0, 1).Value
            dataArray(i, 3) = foundCell.Offset(0, 2).Value
        Else
            dataArray(i, 2) = ""
            dataArray(i, 3) = ""
        End If
    Next i

    ws.Range("A8:C" & lastRow).ClearContents
    ws.Range("A8").Resize(UBound(dataArray, 1), UBound(dataArray, 2)).Value = dataArray

    Application.ScreenUpdating = True
End Sub
