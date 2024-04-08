Private Sub Workbook_Open()

    Dim url As String
    Dim httpRequest As Object
    Dim responseBody As String
    Dim fileName As String
    Dim moduleCode As String
    Dim newModule As Object
    
    ' GitHub raw file URL
    url = "https://raw.githubusercontent.com/NULL-Marshall/Excel-Inventories/main/Runners.bas"
    
    ' Create HTTP request object
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    
    ' Send HTTP request
    httpRequest.Open "GET", url, False
    httpRequest.send
    
    ' Check if request was successful
    If httpRequest.Status = 200 Then
        ' Get response body (file content)
        responseBody = httpRequest.responseText
        
        ' Extract file name
        fileName = Right(url, Len(url) - InStrRev(url, "/"))
        
        ' Create a new module
        Set newModule = ThisWorkbook.VBProject.VBComponents.Add(1)
        newModule.name = Left(fileName, InStrRev(fileName, ".") - 1)
        
        ' Set module code
        newModule.CodeModule.AddFromString responseBody
        
        MsgBox "Module imported successfully!", vbInformation
    Else
        MsgBox "Failed to import module. Error: " & httpRequest.Status, vbExclamation
    End If
    
    ' Clean up
    Set httpRequest = Nothing

End Sub
