Sub LoadModule(name As String)
    Dim url As String
    Dim httpRequest As Object
    Dim responseBody As String
    Dim existingModule As Object
    Dim newModule As Object
    
    url = "https://raw.githubusercontent.com/NULL-Marshall/Excel-Inventories/main/" & name & ".bas"
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    httpRequest.Open "GET", url, False
    httpRequest.send
    
    If httpRequest.Status = 200 Then
        responseBody = httpRequest.responseText

        On Error Resume Next
        Set existingModule = ThisWorkbook.VBProject.VBComponents(name)
        On Error GoTo 0
        
        If Not existingModule Is Nothing Then
            existingModule.CodeModule.DeleteLines 1, existingModule.CodeModule.CountOfLines
            existingModule.CodeModule.AddFromString responseBody
            MsgBox "Module '" & name & "' updated successfully!", vbInformation
        Else
            Set newModule = ThisWorkbook.VBProject.VBComponents.Add(1)
            newModule.Name = name
            newModule.CodeModule.AddFromString responseBody
            MsgBox "Module '" & name & "' imported successfully!", vbInformation
        End If
    Else
        MsgBox "Failed to import module. Error: " & httpRequest.Status, vbExclamation
    End If
    
    Set httpRequest = Nothing
End Sub

Sub CheckModule(name As String, version As String, desc As String)
    Dim WS As Worksheet
    Dim lastRow As Long
    Dim foundCell As Range
    Dim searchRange As Range
    
    Set WS = ThisWorkbook.Worksheets("Modules")
    lastRow = WS.Cells(WS.Rows.count, "A").End(xlUp).Row
    Set searchRange = WS.Range(WS.Cells(2, 1), WS.Cells(lastRow, 1))
    Set foundCell = searchRange.Find(What:=name, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not foundCell Is Nothing Then
        If foundCell.Offset(0, 1).Value < Val(version) Then
            If MsgBox("A new version of '" & name & "' is available. Do you want to update?", vbQuestion + vbYesNo) = vbYes Then
                foundCell.Offset(0, 1).Value = version
                foundCell.Offset(0, 2).Value = Format(Date, "MM/DD/YYYY")
                foundCell.Offset(0, 3).Value = desc
    
                LoadModule name
            Else
                MsgBox "Update canceled by user.", vbInformation
            End If
        End If
    Else
        If MsgBox("Module '" & name & "' is not installed. Do you want to install?", vbQuestion + vbYesNo) = vbYes Then
            WS.Cells(lastRow + 1, 1).Value = name
            WS.Cells(lastRow + 1, 2).Value = version
            WS.Cells(lastRow + 1, 3).Value = Format(Date, "MM/DD/YYYY")
            WS.Cells(lastRow + 1, 4).Value = desc

            LoadModule name
        Else
            MsgBox "Install canceled by user.", vbInformation
        End If
    End If
End Sub

Public Sub UpdateModules()
    Dim url As String
    Dim httpRequest As Object
    Dim responseBody As String
    Dim linesArray() As String
    Dim dataArray2D() As String
    Dim i As Long
    
    url = "https://raw.githubusercontent.com/NULL-Marshall/Excel-Inventories/main/Versions.txt"
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    httpRequest.Open "GET", url, False
    httpRequest.send
    
    If httpRequest.Status = 200 Then
        responseBody = httpRequest.responseText
        linesArray = Split(responseBody, vbCrLf)
        ReDim dataArray2D(1 To UBound(linesArray) + 1, 1 To 3)
        
        For i = LBound(linesArray) To UBound(linesArray)
            If linesArray(i) <> "" Then
                Dim parts() As String
                parts = Split(linesArray(i), " | ")
                
                dataArray2D(i + 1, 1) = parts(0)
                dataArray2D(i + 1, 2) = parts(1)
                dataArray2D(i + 1, 3) = Left(parts(2), Len(parts(2)) - 1)
            End If
        Next i
    Else
        MsgBox "Failed to retrieve file. Error: " & httpRequest.Status, vbExclamation
    End If

    Set httpRequest = Nothing

    For i = LBound(dataArray2D, 1) To UBound(dataArray2D, 1)
        If Not IsEmpty(dataArray2D(i, 1)) Then
            Call CheckModule(dataArray2D(i, 1), dataArray2D(i, 2), dataArray2D(i, 3))
        End If
    Next i
End Sub
