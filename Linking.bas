'Version 1.1
'Creaded by Marshall

Sub Process(cell As Range, target As Range)
    Dim checkCell As Range
    Dim isValid As Boolean
    
    Set checkCell = cell.Worksheet.Cells(cell.Row + 1, target.Column)

    isValid = False
    If checkCell.Value <> "" Then ' Check if cell is not empty
        If Right(checkCell.Value, 1) <> "-" Then
            isValid = True
        End If
    End If

    If isValid Then
        MsgBox "Valid: The cell below " & cell.Address & " in column " & target.Column & " does not end with '-'"
    Else
        MsgBox "Not Valid: The cell below " & cell.Address & " in column " & target.Column & " ends with '-' or is empty"
    End If
End Sub

Public Sub Check(ws As Worksheet, target As Range)
    Dim Links As Worksheet
    Dim searchRange As Range
    Dim cell As Range
    Dim lastRow As Long

    Set Links = Sheets("Linking")
    lastRow = Links.Cells(Links.Rows.count, "A").End(xlUp).Row
    Set searchRange = Links.Range(Links.Cells(2, 1), Links.Cells(lastRow, 1))
    
    For Each cell In searchRange
        If cell.Value = ws.Name Then
            MsgBox "Found a link for " & wsName & " in " & cell.Address
            Call Linking.Process(cell, target)
        End If
    Next cell
End Sub
