'Version 1.4
'Creaded by Marshall

Sub Load(name As String, isModule As Boolean)
    Dim url As String
    Dim httpRequest As Object
    Dim responseBody As String
    Dim existingComponent As Object
    Dim newComponent As Object
    Dim componentType As Integer

    ' Set the URL and component type based on whether it's a module or a class
    If isModule Then
        url = "https://raw.githubusercontent.com/NULL-Marshall/Excel-Inventories/main/Modules/" & name & ".bas"
        componentType = 1  ' vbext_ct_StdModule is 1
    Else
        url = "https://raw.githubusercontent.com/NULL-Marshall/Excel-Inventories/main/Classes/" & name & ".cls"
        componentType = 2  ' vbext_ct_ClassModule is 2
    End If
    
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    httpRequest.Open "GET", url, False
    httpRequest.send
    
    If httpRequest.Status = 200 Then
        responseBody = httpRequest.responseText

        On Error Resume Next
        Set existingComponent = ThisWorkbook.VBProject.VBComponents(name)
        On Error GoTo 0
        
        If Not existingComponent Is Nothing Then
            existingComponent.CodeModule.DeleteLines 1, existingComponent.CodeModule.CountOfLines
            existingComponent.CodeModule.AddFromString responseBody
            MsgBox "Component '" & name & "' updated successfully!", vbInformation
        Else
            Set newComponent = ThisWorkbook.VBProject.VBComponents.Add(componentType)
            newComponent.Name = name
            newComponent.CodeModule.AddFromString responseBody
            MsgBox "Component '" & name & "' imported successfully!", vbInformation
        End If
    Else
        MsgBox "Failed to import component. Error: " & httpRequest.Status, vbExclamation
    End If
    
    Set httpRequest = Nothing
End Sub

Sub check(name As String, version As String, desc As String, isModule As Boolean)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim foundCell As Range
    Dim searchRange As Range
    
    Set ws = ThisWorkbook.Worksheets("Modules")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Set searchRange = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 1))
    Set foundCell = searchRange.Find(What:=name, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not foundCell Is Nothing Then
        If foundCell.Offset(0, 1).Value < Val(version) Then
            If MsgBox("A new version of '" & name & "' is available. Do you want to update?", vbQuestion + vbYesNo) = vbYes Then
                foundCell.Offset(0, 1).Value = version
                foundCell.Offset(0, 2).Value = Format(Date, "MM/DD/YYYY")
                foundCell.Offset(0, 3).Value = desc
    
                Control.Load name, isModule
            Else
                MsgBox "Update canceled by user.", vbInformation
            End If
        End If
    Else
        If MsgBox("Component '" & name & "' is not installed. Do you want to install?", vbQuestion + vbYesNo) = vbYes Then
            ws.Cells(lastRow + 1, 1).Value = name
            ws.Cells(lastRow + 1, 2).Value = version
            ws.Cells(lastRow + 1, 3).Value = Format(Date, "MM/DD/YYYY")
            ws.Cells(lastRow + 1, 4).Value = desc

            Control.Load name, isModule
        Else
            MsgBox "Install canceled by user.", vbInformation
        End If
    End If
End Sub


Public Sub Update()
    Dim url As String
    Dim httpRequest As Object
    Dim responseBody As String
    Dim linesArray() As String
    Dim dataArray2D() As String
    Dim i As Long
    Dim isModule As Boolean
    
    url = "https://raw.githubusercontent.com/NULL-Marshall/Excel-Inventories/main/Versions.txt"
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    httpRequest.Open "GET", url, False
    httpRequest.send
    
    If httpRequest.Status = 200 Then
        responseBody = httpRequest.responseText
        linesArray = Split(responseBody, vbCrLf)
        If UBound(linesArray) = 0 Then
            linesArray = Split(responseBody, vbLf)
        End If
        If UBound(linesArray) = 0 Then
            linesArray = Split(responseBody, vbCr)
        End If
        ReDim dataArray2D(1 To UBound(linesArray) + 1, 1 To 4)
        
        For i = LBound(linesArray) To UBound(linesArray)
            If linesArray(i) <> "" Then
                Dim parts() As String
                parts = Split(linesArray(i), " | ")
                
                dataArray2D(i + 1, 1) = parts(0)
                dataArray2D(i + 1, 2) = parts(1)
                dataArray2D(i + 1, 3) = parts(2)
                dataArray2D(i + 1, 4) = Left(parts(3), Len(parts(3)))
            End If
        Next i
    Else
        MsgBox "Failed to retrieve file. Error: " & httpRequest.Status, vbExclamation
    End If

    Set httpRequest = Nothing

    For i = LBound(dataArray2D, 1) To UBound(dataArray2D, 1) - 1
        If Not IsEmpty(dataArray2D(i, 2)) Then
            isModule = (dataArray2D(i, 1) = "M")
            Call check(dataArray2D(i, 2), dataArray2D(i, 3), dataArray2D(i, 4), isModule)
        End If
    Next i
End Sub

