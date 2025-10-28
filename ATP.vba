Sub LookupFromExcel()
    Dim Chassis_path As String
    Chassis_path = ActiveDocument.path & "\Frame_Report.xlsx"
    ChassisSN Chassis_path
    Dim Card_path As String
    Card_path = ActiveDocument.path & "\Card_Report.xlsx"
    CardReport Card_path
End Sub

Function ChassisSN(path As String)
    Dim xlApp As Object
    Dim xlBook As Object
    Dim ws As Object
    Dim lookupValue As String
    Dim result As Variant
    Dim excelFile As String
    Dim appDataPath As String
    Dim doc As Document
    Set doc = ActiveDocument

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
        doc.Fields.Update  ' update any { DOCVARIABLE } fields in document
    End If
    On Error GoTo 0

    ' --- Handle result ---
    'If IsError(result) Then
        'MsgBox "Value not found!"
    'Else
        'MsgBox "Found: " & result
    'End If

    ' --- Clean up ---
    xlBook.Close False
    xlApp.Quit
    Set ws = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
End Function

Function CardReport(path As String)
    Dim xlApp As Object
    Dim xlBook As Object
    Dim ws As Object
    Dim lookupValue As String
    Dim result As Variant
    Dim excelFile As String
    Dim appDataPath As String
    Dim doc As Document
    Set doc = ActiveDocument

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
        doc.Fields.Update  ' update any { DOCVARIABLE } fields in document
    End If
    ' Look for CXP
    result = xlApp.Run("'" & appDataPath & "\Microsoft\AddIns\Convert Date to String.xlam" & "'!VLookupRegexAll", lookupValue, ws.Range("A4:P1501"), 15, 5, "*CXP*", 16, "Used")
    If IsError(result) Then
        doc.Variables("cxp").Value = "-"
    Else
        doc.Variables("cxp").Value = CStr(result)
        doc.Fields.Update  ' update any { DOCVARIABLE } fields in document
    End If
    ' Look for V8T402
    result = xlApp.Run("'" & appDataPath & "\Microsoft\AddIns\Convert Date to String.xlam" & "'!VLookupRegexAll", lookupValue, ws.Range("A4:P1501"), 15, 5, "V8T402", 16, "Used")
    If IsError(result) Then
        doc.Variables("mt402").Value = "-"
    Else
        doc.Variables("mt402").Value = CStr(result)
        doc.Fields.Update  ' update any { DOCVARIABLE } fields in document
    End If
    ' Look for S7N402
    result = xlApp.Run("'" & appDataPath & "\Microsoft\AddIns\Convert Date to String.xlam" & "'!VLookupRegexAll", lookupValue, ws.Range("A4:P1501"), 15, 5, "S7N402", 16, "Used")
    If IsError(result) Then
        doc.Variables("mn402").Value = "-"
    Else
        doc.Variables("mn402").Value = CStr(result)
        doc.Fields.Update  ' update any { DOCVARIABLE } fields in document
    End If
    On Error GoTo 0

    ' --- Handle result ---
    'If IsError(result) Then
        'MsgBox "Value not found!"
    'Else
        'MsgBox "Found: " & result
    'End If

    ' --- Clean up ---
    xlBook.Close False
    xlApp.Quit
    Set ws = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
End Function



