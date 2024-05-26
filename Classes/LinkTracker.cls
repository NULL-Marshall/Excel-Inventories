Option Explicit

Private sheetName() As String
Private sheetKeys() As Variant
Private sheetLength() As Long
Private rowDeleted() As Boolean
Private sheetUsed() As Boolean

Private Function Index(s As String) As Long
    Dim i As Long
    For i = LBound(sheetName) To UBound(sheetName)
        If sheetName(i) = s Then
            Index = i
            Exit Function
        End If
    Next i
    Index = 0
End Function

Public Function DeleteEvent(ws As Worksheet) As Boolean
    Dim i As Long
    i = Index(ws.name)
    
    If sheetLength(i) > ws.UsedRange.Rows.Count Then
        DeleteEvent = True
    Else
        DeleteEvent = False
    End If
End Function

Public Function InsertEvent(ws As Worksheet) As Boolean
    Dim i As Long
    i = Index(ws.name)
    
    MsgBox sheetLength(i) & " | " & ws.UsedRange.Rows.Count
    If sheetLength(i) < ws.UsedRange.Rows.Count Then
        InsertEvent = True
    Else
        InsertEvent = False
    End If
End Function

Public Function key(s As String, r As Long) As String
    Dim i As Long
    i = Index(s)
    
    key = sheetKeys(i)(r - 1, 1)
End Function

Public Function Used(s As String) As Boolean
    Dim i As Long
    i = Index(s)
    
    Used = sheetUsed(i)
End Function

Public Sub Mark(s As String)
    Dim i As Long
    i = Index(s)
    sheetUsed(i) = True
End Sub

Public Sub Unmark(s As String)
    Dim i As Long
    i = Index(s)
    sheetUsed(i) = False
End Sub

Public Sub UpdateKeys(ws As Worksheet)
    Dim links As Worksheet
    Dim tempKeys() As Variant
    
    Dim searchRange As Range
    Dim foundCell As Range
    Dim keysRange As Range
    Dim lastRow As Long
    Dim i As Long
    
    i = Index(ws.name)
    Set links = Sheets("Linking")
    lastRow = links.Cells(links.Rows.Count, "A").End(xlUp).Row
    Set searchRange = links.Range(links.Cells(2, 1), links.Cells(lastRow, 5))
    Set foundCell = searchRange.Find(What:=ws.name, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not foundCell Is Nothing Then
        lastRow = ws.Cells(ws.Rows.Count, foundCell.Offset(0, 1).Value).End(xlUp).Row
        Set keysRange = ws.Range(ws.Cells(2, foundCell.Offset(0, 1).Value), ws.Cells(lastRow, foundCell.Offset(0, 1).Value))
        ReDim tempKeys(2 To lastRow, 1 To 1)
        tempKeys = keysRange.Value
        sheetKeys(i) = tempKeys
        sheetLength(i) = ws.UsedRange.Rows.Count
    End If
End Sub

Public Sub Reset()
    Dim ws As Worksheet
    Dim links As Worksheet
    Dim sheetIndex As Long
    
    Dim searchRange As Range
    Dim foundCell As Range
    Dim lastRow As Long
    
    ReDim sheetName(1 To 1) As String
    ReDim sheetKeys(1 To 1) As Variant
    ReDim sheetLength(1 To 1) As Long
    ReDim rowDeleted(1 To 1) As Boolean
    ReDim sheetUsed(1 To 1) As Boolean
    
    Set links = Sheets("Linking")
    lastRow = links.Cells(links.Rows.Count, "A").End(xlUp).Row
    Set searchRange = links.Range(links.Cells(2, 1), links.Cells(lastRow, 5))
    
    For Each ws In ThisWorkbook.Sheets
        Set foundCell = searchRange.Find(What:=ws.name, LookIn:=xlValues, LookAt:=xlWhole)
        If Not foundCell Is Nothing Then
            sheetIndex = sheetIndex + 1
            
            ReDim Preserve sheetName(1 To sheetIndex) As String
            ReDim Preserve sheetKeys(1 To sheetIndex) As Variant
            ReDim Preserve sheetLength(1 To sheetIndex) As Long
            ReDim Preserve rowDeleted(1 To sheetIndex) As Boolean
            ReDim Preserve sheetUsed(1 To sheetIndex) As Boolean
            
            sheetName(sheetIndex) = ws.name
            Call UpdateKeys(ws)
            rowDeleted(sheetIndex) = False
            sheetUsed(sheetIndex) = False
        End If
    Next ws
End Sub
