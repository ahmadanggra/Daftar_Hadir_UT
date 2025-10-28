Sub LookupFromExcel(site_id As String, site_name As String, site_ref As String)
    Dim doc As Document
    Set doc = ActiveDocument

    ' Turn off screenupdate to speedup looping proccess
    Application.ScreenUpdating = False
    
    ' Update site_id (Work Package) & site_name (Lokasi)
    doc.Variables("site_id").Value = site_id
    doc.Variables("site_name").Value = site_name
    ' Update site_ref
    doc.CustomDocumentProperties("site_ref").Value = site_ref
    
    ' Update frame/chassis data
    Dim Chassis_path As String
    Chassis_path = ActiveDocument.path & "\Frame_Report.xlsx"
    ChassisSN Chassis_path, doc
    
    ' Update cardboard data
    Dim Card_path As String
    Card_path = ActiveDocument.path & "\Card_Report.xlsx"
    CardReport Card_path, doc

    ' Turn on screenupdate again
    Application.ScreenUpdating = True
    ' Update all doc varialbe
    doc.Fields.Update
End Sub

Function ChassisSN(path As String, doc As Document)
    Dim xlApp As Object
    Dim xlBook As Object
    Dim ws As Object
    Dim lookupValue As String
    Dim result As Variant
    Dim excelFile As String
    Dim appDataPath As String

    ' --- Path to your Excel file ---
    ' --- Must using fullpath ---
    excelFile = path

    ' --- Create Excel instance ---
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False  ' Optional: make Excel visible

    ' --- Open the workbook ---
    Set xlBook = xlApp.Workbooks.Open(excelFile)
    Set ws = xlBook.Sheets("Frame Report") ' change as needed

    ' --- Value to search ---
    lookupValue = doc.CustomDocumentProperties("site_ref").Value

    ' --- Perform VLOOKUP using regex ---
    ' --- Must using fullpath ---
    appDataPath = Environ("APPDATA")
    On Error Resume Next
    result = xlApp.Run("'" & appDataPath & "\Microsoft\AddIns\Convert Date to String.xlam" & "'!VLookupRegex", lookupValue, ws.Range("A4:I127"), 9, 3, "*M24*")
    If IsError(result) Then
        doc.Variables("chassis_sn").Value = "-"
    Else
        doc.Variables("chassis_sn").Value = CStr(result)
    End If
    On Error GoTo 0

    ' --- Clean up ---
    xlBook.Close False
    xlApp.Quit
    Set ws = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
End Function

Function CardReport(path As String, doc As Document)
    Dim xlApp As Object
    Dim xlBook As Object
    Dim ws As Object
    Dim lookupValue As String
    Dim result As Variant
    Dim excelFile As String
    Dim appDataPath As String

    ' --- Path to your Excel file ---
    ' --- Must using fullpath ---
    excelFile = path

    ' --- Create Excel instance ---
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False  ' Optional: make Excel visible

    ' --- Open the workbook ---
    Set xlBook = xlApp.Workbooks.Open(excelFile)
    Set ws = xlBook.Sheets("Card Report") ' change as needed

    ' --- Value to search ---
    lookupValue = doc.CustomDocumentProperties("site_ref").Value

    ' --- Perform VLOOKUP using regex ---
    ' --- Must using fullpath ---
    appDataPath = Environ("APPDATA")
    On Error Resume Next
    '=VLookupRegexAll(A1;'Card Report_2025-10-27_08-17-06'!A4:P1501;15; 5;"*CXP*";16; "Used")
    ' Look for DCP
    result = xlApp.Run("'" & appDataPath & "\Microsoft\AddIns\Convert Date to String.xlam" & "'!VLookupRegexAll", lookupValue, ws.Range("A4:P1501"), 15, 5, "*DCP*", 16, "Used")
    If IsError(result) Then
        doc.Variables("dcp").Value = "-"
    Else
        doc.Variables("dcp").Value = CStr(result)
    End If
    ' Look for CXP
    result = xlApp.Run("'" & appDataPath & "\Microsoft\AddIns\Convert Date to String.xlam" & "'!VLookupRegexAll", lookupValue, ws.Range("A4:P1501"), 15, 5, "*CXP*", 16, "Used")
    If IsError(result) Then
        doc.Variables("cxp").Value = "-"
    Else
        doc.Variables("cxp").Value = CStr(result)
    End If
    ' Look for V8T402
    result = xlApp.Run("'" & appDataPath & "\Microsoft\AddIns\Convert Date to String.xlam" & "'!VLookupRegexAll", lookupValue, ws.Range("A4:P1501"), 15, 5, "V8T402", 16, "Used")
    If IsError(result) Then
        doc.Variables("mt402").Value = "-"
    Else
        doc.Variables("mt402").Value = CStr(result)
    End If
    ' Look for S7N402
    result = xlApp.Run("'" & appDataPath & "\Microsoft\AddIns\Convert Date to String.xlam" & "'!VLookupRegexAll", lookupValue, ws.Range("A4:P1501"), 15, 5, "S7N402", 16, "Used")
    If IsError(result) Then
        doc.Variables("mn402").Value = "-"
    Else
        doc.Variables("mn402").Value = CStr(result)
    End If
    On Error GoTo 0

    ' --- Clean up ---
    xlBook.Close False
    xlApp.Quit
    Set ws = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
End Function



