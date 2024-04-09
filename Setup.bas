Sub LoadModule(name As String)
    Dim url As String
    Dim httpRequest As Object
    Dim responseBody As String
    Dim existingModule As Object
    Dim newModule As Object
    
    ' GitHub raw file URL
    url = "https://raw.githubusercontent.com/NULL-Marshall/Excel-Inventories/main/" + name + ".bas"
    
    ' Create HTTP request object
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    
    ' Send HTTP request
    httpRequest.Open "GET", url, False
    httpRequest.send
    
    ' Check if request was successful
    If httpRequest.Status = 200 Then
        ' Get response body (file content)
        responseBody = httpRequest.responseText
        
        On Error Resume Next
        ' Attempt to access an existing module with the same name
        Set existingModule = ThisWorkbook.VBProject.VBComponents(fileName)
        On Error GoTo 0
        
        If Not existingModule Is Nothing Then
            ' Module already exists, replace its code with the new code
            existingModule.CodeModule.DeleteLines 1, existingModule.CodeModule.CountOfLines
            existingModule.CodeModule.AddFromString responseBody
            MsgBox "Module '" & name & "' updated successfully!", vbInformation
        Else
            ' Module does not exist, create a new module and add the code
            Set newModule = ThisWorkbook.VBProject.VBComponents.Add(1)
            newModule.name = name
            newModule.CodeModule.AddFromString responseBody
            MsgBox "Module '" & name & "' imported successfully!", vbInformation
        End If
    Else
        MsgBox "Failed to import module. Error: " & httpRequest.Status, vbExclamation
    End If
    
    ' Clean up
    Set httpRequest = Nothing

End Sub

Sub CheckModule(searchArray() As Variant)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim foundCell As Range
    Dim searchRange As Range
    
    Set ws = ThisWorkbook.Worksheets("Modules")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Set searchRange = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 1))
    Set foundCell = searchRange.Find(What:=searchArray(0), LookIn:=xlValues, LookAt:=xlWhole)
    
    ' Check if the value is found
    If Not foundCell Is Nothing Then
        If foundCell.Offset(0, 1).Value < searchArray(1) Then
            foundCell.Offset(0, 1).Value = searchArray(1)
            foundCell.Offset(0, 2).Value = Format(Date, "MM/DD/YYYY")
            foundCell.Offset(0, 3).Value = searchArray(2)
            
            LoadModule (searchArray(0))
        End If
    Else
        ws.Cells(lastRow + 1, 1).Value = searchArray(0)
        ws.Cells(lastRow + 1, 2).Value = searchArray(1)
        ws.Cells(lastRow + 1, 3).Value = Format(Date, "MM/DD/YYYY")
        ws.Cells(lastRow + 1, 4).Value = searchArray(2)

        LoadModule (searchArray(0))
    End If
End Sub

Sub UpdateModules()
    'VARIABLES
    Dim url As String
    Dim httpRequest As Object
    Dim responseBody As String
    
    Dim linesArray() As String
    Dim dataArray() As String
    Dim dataArray2D() As String
    
    Dim i As Long
    Dim j As Long
    
    ' GET VERSIONS
    url = "https://raw.githubusercontent.com/NULL-Marshall/Excel-Inventories/main/Versions.txt.bas"
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    httpRequest.Open "GET", url, False
    httpRequest.send
    
    ' CHECK REQUEST STAUS
    If httpRequest.Status = 200 Then
        responseBody = httpRequest.responseText
        
        'PARSE RESPONSE INTO ARRAY
        linesArray = Split(responseBody, vbCrLf)
        ReDim dataArray2D(1 To UBound(linesArray) + 1, 1 To 3)
        
        For i = LBound(linesArray) To UBound(linesArray)
            dataArray = Split(linesArray(i), " | ")
            If UBound(dataArray) >= 2 Then
                dataArray2D(i + 1, 1) = Trim(dataArray(0))
                dataArray2D(i + 1, 2) = Trim(dataArray(1))
                dataArray2D(i + 1, 3) = Trim(dataArray(2))
            Else
                ' Handle invalid lines (if needed)
                dataArray2D(i + 1, 1) = ""
                dataArray2D(i + 1, 2) = ""
                dataArray2D(i + 1, 3) = ""
            End If
        Next i
    Else
        ' Display error message if request fails
        MsgBox "Failed to retrieve file. Error: " & httpRequest.Status, vbExclamation
    End If
    
    'CLOSE REQUEST
    Set httpRequest = Nothing

    
    For j = LBound(dataArray2D, 1) To UBound(dataArray2D, 1)
        Call CheckModule(dataArray(j))
    Next j
End Sub
