'Version 1.2
'Creaded by Marshall

Dim tracker As New LinkTracker

Function PreviousElement(cell As Range, target As Range) As Long
    Dim searchSheet As Worksheet
    Dim searchRange As Range
    Dim foundCell As Range
    Dim lastRow As Long
    Dim i As Long
    
    Set searchSheet = ThisWorkbook.Worksheets(cell.Offset(0, 2).Value)
    lastRow = searchSheet.Cells(searchSheet.Rows.Count, cell.Offset(0, 4).Value).End(xlUp).Row
    Set searchRange = searchSheet.Range(searchSheet.Cells(1, cell.Offset(0, 4).Value), searchSheet.Cells(lastRow, cell.Offset(0, 4).Value))
    
    For i = target.Row To 2 Step -1
        Set foundCell = searchRange.Find(What:=tracker.key(target.Worksheet.name, i), LookIn:=xlValues, LookAt:=xlWhole)
        If Not foundCell Is Nothing Then
            PreviousElement = foundCell.Row
            Exit Function
        End If
    Next i
    
    PreviousElement = 1
End Function


Function RemoveLastCharacter(ByVal str As String) As String
    Dim charactersToCheck As String
    charactersToCheck = "+-*_=!"
    If str = "" Then
        RemoveLastCharacter = str
    ElseIf InStr(charactersToCheck, Right(str, 1)) > 0 Then
        RemoveLastCharacter = Left(str, Len(str) - 1)
    Else
        RemoveLastCharacter = str
    End If
End Function

Function ValidateCells(cell As Range, target As Range) As Boolean
    Dim searchRange As Range
    Dim searchCell As Range
    Dim copyColumn As Variant
    Dim colCount As Long
    Dim clear As Boolean
    Dim total As Long
    Dim check As Long

    colCount = cell.Worksheet.Cells(cell.Row + 1, cell.Worksheet.Columns.Count).End(xlToLeft).Column
    Set searchRange = cell.Worksheet.Range(cell.Offset(1, 0), cell.Offset(1, colCount - 1))
    
    clear = True
    total = 0
    check = 0
    
    For Each searchCell In searchRange
        If Not isEmpty(searchCell.Value) Then
            copyColumn = trim(RemoveLastCharacter(searchCell.Value))
            If Right(searchCell.Value, 1) = "+" Then
                If isEmpty(target.Worksheet.Cells(target.Row, searchCell.Column).Value) Or target.Worksheet.Cells(target.Row, searchCell.Column).Value = 0 Then
                    clear = False
                    Exit For
                End If
            ElseIf Right(searchCell.Value, 1) = "-" Then
                If Not (isEmpty(target.Worksheet.Cells(target.Row, searchCell.Column).Value) Or target.Worksheet.Cells(target.Row, searchCell.Column).Value = 0) Then
                    clear = False
                    Exit For
                End If
            ElseIf Right(searchCell.Value, 1) = "*" Then
                total = total + 1
                If isEmpty(target.Worksheet.Cells(target.Row, searchCell.Column).Value) Or target.Worksheet.Cells(target.Row, searchCell.Column).Value = 0 Then
                    check = check + 1
                End If
            End If
        End If
    Next searchCell
    
    ValidateCells = clear And (total > check)
End Function

Sub Update(copyRange As Range, pasteRange As Range, referenceRange As Range, Optional inverted As Boolean = False)
    Dim destArray() As Variant
    Dim searchCell As Range
    Dim columnLetter As String
    Dim columnIndex As Long
    Dim i As Long
    
    Application.ScreenUpdating = False
    destArray = pasteRange.Value

    For Each searchCell In referenceRange
        If Not searchCell Is Nothing And searchCell.Value <> "" Then
            columnLetter = RemoveLastCharacter(searchCell.Value)
            columnIndex = Range(columnLetter & "1").Column
            If Right(searchCell.Value, 1) = "_" Then
                If inverted Then
                    pasteRange.Worksheet.Cells(pasteRange.Row, searchCell.Column).Interior.Color = copyRange.Worksheet.Cells(copyRange.Row, columnIndex).Interior.Color
                Else
                    pasteRange.Worksheet.Cells(pasteRange.Row, columnIndex).Interior.Color = copyRange.Worksheet.Cells(copyRange.Row, searchCell.Column).Interior.Color
                End If
            ElseIf Right(searchCell.Value, 1) = "=" Then
                If inverted Then
                    pasteRange.Worksheet.Cells(pasteRange.Row, searchCell.Column).Interior.Color = copyRange.Worksheet.Cells(copyRange.Row, columnIndex).Interior.Color
                    destArray(1, searchCell.Column) = copyRange.Cells(columnIndex)
                Else
                    pasteRange.Worksheet.Cells(pasteRange.Row, columnIndex).Interior.Color = copyRange.Worksheet.Cells(copyRange.Row, searchCell.Column).Interior.Color
                    destArray(1, columnIndex) = copyRange.Cells(searchCell.Column)
                End If
            Else
                If inverted Then
                    destArray(1, searchCell.Column) = copyRange.Cells(columnIndex)
                Else
                    destArray(1, columnIndex) = copyRange.Cells(searchCell.Column)
                End If
            End If
        End If
    Next searchCell
    pasteRange.Value = destArray
    Application.ScreenUpdating = True
End Sub

Sub CopyLink(cell As Range, target As Range)
    Dim key As String
    
    Dim copySheet As Worksheet
    Dim pasteSheet As Worksheet
    Dim referenceSheet As Worksheet
    
    Dim searchRange As Range
    Dim copyRange As Range
    Dim pasteRange As Range
    Dim referenceRange As Range
    
    Dim foundCell As Range
    Dim lastRow As Long
    
    key = tracker.key(target.Worksheet.name, target.Row)
    
    Set copySheet = target.Worksheet
    Set pasteSheet = ThisWorkbook.Worksheets(cell.Offset(0, 2).Value)
    Set referenceSheet = Sheets("Linking")
    
    lastRow = pasteSheet.Cells(pasteSheet.Rows.Count, cell.Offset(0, 4).Value).End(xlUp).Row
    Set searchRange = pasteSheet.Range(pasteSheet.Cells(1, cell.Offset(0, 4).Value), pasteSheet.Cells(lastRow, cell.Offset(0, 4).Value))
    Set foundCell = searchRange.Find(What:=key, LookIn:=xlValues, LookAt:=xlWhole)
    
     If Not foundCell Is Nothing Then
        Set copyRange = copySheet.Range(copySheet.Cells(target.Row, 1), copySheet.Cells(target.Row, copySheet.UsedRange.Columns.Count))
        Set pasteRange = pasteSheet.Range(pasteSheet.Cells(foundCell.Row, 1), pasteSheet.Cells(foundCell.Row, pasteSheet.UsedRange.Columns.Count))
        Set referenceRange = referenceSheet.Range(referenceSheet.Cells(cell.Row + 1, 1), referenceSheet.Cells(cell.Row + 1, referenceSheet.Cells(cell.Row + 1, referenceSheet.Columns.Count).End(xlToLeft).Column))
        Call Linking.Update(copyRange, pasteRange, referenceRange)
    End If
End Sub

Sub ListLink(cell As Range, target As Range)
    Dim key As String
    
    Dim copySheet As Worksheet
    Dim pasteSheet As Worksheet
    Dim referenceSheet As Worksheet
    
    Dim searchRange As Range
    Dim copyRange As Range
    Dim pasteRange As Range
    Dim referenceRange As Range
    
    Dim foundCell As Range
    Dim lastRow As Long
    Dim valid As Boolean

    key = tracker.key(target.Worksheet.name, target.Row)
    
    Set copySheet = target.Worksheet
    Set pasteSheet = ThisWorkbook.Worksheets(cell.Offset(0, 2).Value)
    Set referenceSheet = Sheets("Linking")
    
    lastRow = pasteSheet.Cells(pasteSheet.Rows.Count, cell.Offset(0, 4).Value).End(xlUp).Row
    Set searchRange = pasteSheet.Range(pasteSheet.Cells(1, cell.Offset(0, 4).Value), pasteSheet.Cells(lastRow, cell.Offset(0, 4).Value))
    Set foundCell = searchRange.Find(What:=key, LookIn:=xlValues, LookAt:=xlWhole)
    
    valid = ValidateCells(cell, target)

    If valid Then
        If foundCell Is Nothing Then
            lastRow = PreviousElement(cell, target)
            If lastRow = 1 Then
                pasteSheet.Rows(lastRow + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
            Else
                pasteSheet.Rows(lastRow + 1).Insert Shift:=xlDown
            End If
            Set foundCell = pasteSheet.Cells(lastRow + 1, cell.Offset(0, 4).Value)
        End If
        
        If tracker.DeleteEvent(target.Worksheet) Then
            foundCell.EntireRow.Delete
        Else
            Set copyRange = copySheet.Range(copySheet.Cells(target.Row, 1), copySheet.Cells(target.Row, copySheet.UsedRange.Columns.Count))
            Set pasteRange = pasteSheet.Range(pasteSheet.Cells(foundCell.Row, 1), pasteSheet.Cells(foundCell.Row, pasteSheet.UsedRange.Columns.Count))
            Set referenceRange = referenceSheet.Range(referenceSheet.Cells(cell.Row + 1, 1), referenceSheet.Cells(cell.Row + 1, referenceSheet.Cells(cell.Row + 1, referenceSheet.Columns.Count).End(xlToLeft).Column))
            Call Linking.Update(copyRange, pasteRange, referenceRange)
        End If
    Else
        If Not foundCell Is Nothing Then
            foundCell.EntireRow.Delete
        End If
    End If
End Sub

Sub PushLink(cell As Range, target As Range)
    Dim key As String
    
    Dim copySheet As Worksheet
    Dim pasteSheet As Worksheet
    Dim referenceSheet As Workshee
    
    Dim searchRange As Range
    Dim copyRange As Range
    Dim pasteRange As Range
    Dim referenceRange As Range
    
    Dim foundCell As Range
    Dim lastRow As Long

    key = tracker.key(target.Worksheet.name, target.Row)
    
    Set copySheet = target.Worksheet
    Set pasteSheet = ThisWorkbook.Worksheets(cell.Offset(0, 2).Value)
    Set referenceSheet = Sheets("Linking")
    
    lastRow = pasteSheet.Cells(pasteSheet.Rows.Count, cell.Offset(0, 4).Value).End(xlUp).Row
    Set searchRange = pasteSheet.Range(pasteSheet.Cells(1, cell.Offset(0, 4).Value), pasteSheet.Cells(lastRow, cell.Offset(0, 4).Value))
    Set foundCell = searchRange.Find(What:=key, LookIn:=xlValues, LookAt:=xlWhole)

    If foundCell Is Nothing Then
        lastRow = PreviousElement(cell, target)
        If lastRow = 1 Then
            pasteSheet.Rows(lastRow + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
        Else
            pasteSheet.Rows(lastRow + 1).Insert Shift:=xlDown
        End If
        foundCell = pasteSheet.Cells(lastRow, cell.Offset(0, 4).Value)
    End If
    
    Set copyRange = copySheet.Range(copySheet.Cells(target.Row, 1), copySheet.Cells(target.Row, copySheet.UsedRange.Columns.Count))
    Set pasteRange = pasteSheet.Range(pasteSheet.Cells(foundCell.Row, 1), pasteSheet.Cells(foundCell.Row, pasteSheet.UsedRange.Columns.Count))
    Set referenceRange = referenceSheet.Range(referenceSheet.Cells(cell.Row + 1, 1), referenceSheet.Cells(cell.Row + 1, referenceSheet.Cells(cell.Row + 1, referenceSheet.Columns.Count).End(xlToLeft).Column))
    Call Linking.Update(copyRange, pasteRange, referenceRange)
End Sub

Sub PullLink(cell As Range, target As Range)
    Dim key As String
    
    Dim copySheet As Worksheet
    Dim pasteSheet As Worksheet
    Dim referenceSheet As Worksheet
    
    Dim searchRange As Range
    Dim copyRange As Range
    Dim pasteRange As Range
    Dim referenceRange As Range
    
    Dim foundCell As Range
    Dim lastRow As Long
    
    key = tracker.key(target.Worksheet.name, target.Row)
    
    Set copySheet = target.Worksheet
    Set pasteSheet = ThisWorkbook.Worksheets(cell.Offset(0, 2).Value)
    Set referenceSheet = Sheets("Linking")
    
    lastRow = pasteSheet.Cells(pasteSheet.Rows.Count, cell.Offset(0, 4).Value).End(xlUp).Row
    Set searchRange = pasteSheet.Range(pasteSheet.Cells(1, cell.Offset(0, 4).Value), pasteSheet.Cells(lastRow, cell.Offset(0, 4).Value))
    Set foundCell = searchRange.Find(What:=key, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not foundCell Is Nothing Then
        Set copyRange = copySheet.Range(copySheet.Cells(target.Row, 1), copySheet.Cells(target.Row, copySheet.UsedRange.Columns.Count))
        Set pasteRange = pasteSheet.Range(pasteSheet.Cells(foundCell.Row, 1), pasteSheet.Cells(foundCell.Row, pasteSheet.UsedRange.Columns.Count))
        Set referenceRange = referenceSheet.Range(referenceSheet.Cells(cell.Row + 1, 1), referenceSheet.Cells(cell.Row + 1, referenceSheet.Cells(cell.Row + 1, referenceSheet.Columns.Count).End(xlToLeft).Column))
        Call Linking.Update(pasteRange, copyRange, referenceRange, True)
    End If
End Sub

Public Sub check(ws As Worksheet, target As Range)
    Dim links As Worksheet
    Dim searchRange As Range
    Dim cell As Range
    Dim lastRow As Long
    Dim checkCell As Range
    Dim valid As Boolean
    Dim mode As String

    Set links = Sheets("Linking")
    lastRow = links.Cells(links.Rows.Count, "A").End(xlUp).Row
    Set searchRange = links.Range(links.Cells(2, 1), links.Cells(lastRow, 1))
    
    For Each cell In searchRange
        If cell.Value = ws.name Then
            mode = "Skip"
            valid = False
            For Each checkCell In target
                If links.Cells(cell.Row + 1, checkCell.Column).Value <> "" And Right(links.Cells(cell.Row + 1, checkCell.Column).Value, 1) <> "!" Then
                    valid = True
                End If
            Next checkCell
            If Not tracker.Used(cell.Offset(0, 2).Value) And valid Then
                mode = cell.Offset(0, 5).Value
            End If
            
            If mode <> "Skip" And Not (tracker.DeleteEvent(ws) And mode = "List") Then
                tracker.UpdateKeys ws
            End If
            
            If mode = "Pull" Then
                tracker.Mark ws.name
                tracker.Mark cell.Offset(0, 2).Value
            ElseIf Not mode = "Skip" Then
                tracker.Mark ws.name
            End If
            
            Select Case mode
                Case "Copy"
                    Call Linking.CopyLink(cell, target)
                Case "List"
                    Call Linking.ListLink(cell, target)
                Case "Push"
                    Call Linking.PushLink(cell, target)
                Case "Pull"
                    Call Linking.PullLink(cell, target)
                Case "Skip"
                Case Else
                    MsgBox "Unrecognized linking mode: " & mode
            End Select

            If mode = "Pull" Then
                tracker.Unmark ws.name
                tracker.Unmark cell.Offset(0, 2).Value
            ElseIf Not mode = "Skip" Then
                tracker.Unmark ws.name
            End If
        End If
    Next cell
End Sub

Public Sub Intialize(ws As Worksheet)
    tracker.Reset
End Sub
