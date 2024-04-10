'Version 1.0
'Creaded by Marshall

Public Sub Check(ws as Worksheet)
    Dim Links As Worksheet
    Dim searchRange As Range
    Dim cell As Range
    Dim lastRow As Long

    Set Links = Sheets("Linking")
    lastRow = Links.Cells(Links.Rows.count, "A").End(xlUp).Row
    Set searchRange = Links.Range(Links.Cells(2, 1), Links.Cells(lastRow, 1))
    
    For Each cell In searchRange
        If cell.Value = ws.Name Then
            ' Call linking code (replace MsgBox with your linking code)
            MsgBox "Found a link for " & wsName & " in " & cell.Address
        End If
    Next cell
End Sub
