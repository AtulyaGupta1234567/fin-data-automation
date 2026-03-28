Sub FetchNAVByCode()
    Dim http As Object
    Dim url As String
    Dim result As String
    Dim lines As Variant
    Dim i As Long, j As Long
    Dim ws As Worksheet
    Dim schemeCode As String
    Dim splitLine As Variant
    Dim fundName As String
    Dim nav As String
    Dim navDate As String
    Dim found As Boolean
    Dim sheetName As String

    ' Automatically find the correct sheet if needed
    For Each ws In ThisWorkbook.Sheets
        If ws.Name Like "*Your Sheet*" Or ws.Name Like "*Sheet1*" Then
            sheetName = ws.Name
            Exit For
        End If
    Next ws

    ' If no valid sheet is found, exit
    If sheetName = "" Then
        MsgBox "Error: Sheet not found! Please check sheet name.", vbCritical
        Exit Sub
    End If

    ' Set the correct worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' Use AMFI NAV TXT from portal.amfiindia.com
    url = "https://portal.amfiindia.com/spages/NAVAll.txt?" & Format(Now, "YYYYMMDDhhmmss")

    ' Create HTTP object (server-compatible)
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    http.Send
    
    ' Check if data was fetched successfully
    If http.Status <> 200 Then
        MsgBox "Error: Unable to fetch NAV data! Status " & http.Status, vbCritical
        Exit Sub
    End If

    ' Store result
    result = http.responseText
    If Len(result) = 0 Then
        MsgBox "Error: Empty response from AMFI!", vbCritical
        Exit Sub
    End If

    ' Split data into lines
    lines = Split(result, vbCrLf)
    
    ' Loop through scheme codes in Column B (Starting from Row 3)
    For i = 3 To ws.Cells(Rows.Count, 2).End(xlUp).Row
        schemeCode = Trim(ws.Cells(i, 2).Value) ' Read scheme code from column B
        schemeCode = Replace(schemeCode, " ", "") ' Remove all spaces

        ' Skip empty scheme codes
        If schemeCode = "" Then GoTo NextIteration
        
        nav = ""
        fundName = ""
        navDate = ""
        found = False

        ' Debugging: Check if the Scheme Code is being read
        Debug.Print "Checking Scheme Code: " & schemeCode

        ' Search for scheme code in NAV data
        For j = LBound(lines) To UBound(lines)
            splitLine = Split(lines(j), ";")
            If UBound(splitLine) >= 5 Then
                If Replace(Trim(splitLine(1)), " ", "") = schemeCode Then
                    fundName = Trim(splitLine(3)) ' Fund Name
                    nav = Trim(splitLine(4)) ' NAV
                    navDate = Trim(splitLine(5)) ' NAV Date

                    If IsDate(navDate) Then
                        found = True
                        Debug.Print "Match Found - " & fundName & ": " & nav & " (Date: " & navDate & ")"
                    End If

                    Exit For
                End If
            End If
        Next j

        ' Update sheet only if found
        If found Then
            ws.Cells(i, 3).Value = fundName ' Column C - Fund Name
            ws.Cells(i, 5).Value = navDate ' Column E - NAV Date
            ws.Cells(i, 6).Value = nav ' Column F - Current NAV
        End If

NextIteration:
    Next i

    ' Notify user
    MsgBox "NAVs Updated Successfully!", vbInformation
End Sub
