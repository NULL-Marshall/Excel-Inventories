'Pull information from 'Data' sheet to propogate Headers and Footers.
'Version 1.0

Sub UpdateRunners(WS As Worksheet)
    Dim Data As Worksheet
    Dim searchRange As Range
    Dim foundCell As Range
    Dim updated As String
    Dim description As String
    
    Application.ScreenUpdating = False
    
    WS.PageSetup.TopMargin = 60
    WS.PageSetup.BottomMargin = 60
    
    Set Data = Sheets("Data")
    Set searchRange = Data.Range("A8:A25")
    Set foundCell = searchRange.Find(What:=WS.name, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not foundCell Is Nothing Then
        description = foundCell.Offset(0, 1).Value
        updated = foundCell.Offset(0, 2).Value
    Else
        MsgBox "Data not found for: " + WS.name
        description = ""
        updated = ""
    End If
    
    WS.PageSetup.RightHeader = Data.Range("B2").Value + vbCr + "Updated: " + updated
    WS.PageSetup.LeftHeader = "&B&16" + WS.name + " " + description + "&B&11" + vbCr + Data.Range("A3").Value + ": " + Data.Range("B3").Value
    WS.PageSetup.LeftFooter = Data.Range("A4").Value + ": " + Data.Range("B4").Value + vbCr + Data.Range("A5").Value + ": " + Data.Range("B5").Value
    WS.PageSetup.RightFooter = "Page &P of &N"
    
    Application.ScreenUpdating = True
End Sub
