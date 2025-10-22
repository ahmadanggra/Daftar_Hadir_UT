Function DateToWord(dateStr As String) As String
    Dim dt As Date
    Dim dayOfMonth As Integer
    Dim weekOfMonth As Integer
    Dim monthAbbr As String
    Dim yearShort As String
    
    ' Check if cell is empty
    If dateStr = "" Or dateStr = "NA" Then
        DateToWord = ""
        Exit Function
    End If
    ' Convert string to Date (expects dd-mmm-yy format, e.g., "20-Sep-25")
    dt = DateValue(dateStr)
    
    dayOfMonth = Day(dt)
    weekOfMonth = Int((dayOfMonth - 1) / 7) + 1
    monthAbbr = Format(dt, "mmm")
    yearShort = "'" & Format(dt, "yy")
    
    ' Force days 29, 30, 31 ? W4
    If dayOfMonth >= 29 Then
        DateToWord = "W4 " & AdjustMonth(monthAbbr) & yearShort
    Else
        DateToWord = "W" & weekOfMonth & " " & AdjustMonth(monthAbbr) & yearShort
    End If
End Function

Function DateToWordv2(date1 As String, date2 As String) As String
    Dim dt As Date
    Dim dayOfMonth As Integer
    Dim weekOfMonth As Integer
    Dim monthAbbr As String
    Dim yearShort As String
    
    ' Convert string to Date (expects dd-mmm-yy format, e.g., "20-Sep-25")
    If (date1 = "" Or date1 = "NA") And (date2 = "" Or date2 = "") Then
        DateToWordv2 = ""
        Exit Function
    ElseIf date1 = "" Or date1 = "NA" Then
        dt = DateValue(date2)
    ElseIf date2 = "" Or date2 = "NA" Then
        dt = DateValue(date1)
    ElseIf DateValue(date1) > DateValue(date2) Then
        dt = DateValue(date1)
    Else
        dt = DateValue(date2)
    End If
    
    dayOfMonth = Day(dt)
    weekOfMonth = Int((dayOfMonth - 1) / 7) + 1
    monthAbbr = Format(dt, "mmm")
    yearShort = "'" & Format(dt, "yy")
    
    ' Force days 29, 30, 31 ? W4
    If dayOfMonth >= 29 Then
        DateToWordv2 = "W4 " & AdjustMonth(monthAbbr) & yearShort
    Else
        DateToWordv2 = "W" & weekOfMonth & " " & AdjustMonth(monthAbbr) & yearShort
    End If
End Function

Function AdjustMonth(month As String) As String
    If month = "Jun" Then
        month = "Juni"
    ElseIf month = "Jul" Then
        month = "Juli"
    ElseIf month = "Oct" Then
        month = "Okt"
    End If
    AdjustMonth = month
End Function

Function VLookupStatus(LookupValue As String, TableRange As Range, ColNum As Long, StatusIndex As Long, StatusValue As String) As Variant
    Dim r As Range
    Dim i As Long
    
    ' Loop through each row in the table
    For i = 1 To TableRange.Rows.Count
        ' Match Site ID and Status
        If TableRange.Cells(i, 1).Value = LookupValue And _
           TableRange.Cells(i, StatusIndex).Value = StatusValue Then
           
           VLookupStatus = TableRange.Cells(i, ColNum).Value
           Exit Function
        End If
    Next i
    
    ' If not found
    VLookupStatus = CVErr(xlErrNA)
End Function
