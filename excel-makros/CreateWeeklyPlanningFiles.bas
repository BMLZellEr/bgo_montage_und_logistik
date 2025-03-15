Sub CreateWeeklyPlanningFiles()
    ' Purpose: Create a file for each week (KW) of the year 2025
    ' Each file will have the correct dates set in row 2
    
    Dim fso As Object
    Dim baseFolder As String
    Dim year2025Folder As String
    Dim i As Integer
    Dim wb As Workbook
    Dim tempWb As Workbook
    Dim ws As Worksheet
    Dim mondayDate As Date
    Dim weekNum As Integer
    Dim successCount As Integer
    Dim failCount As Integer
    
    ' Create FileSystemObject for advanced file operations
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Store reference to current workbook
    Set tempWb = ThisWorkbook
    
    ' Define the base folder (adjust if needed)
    baseFolder = "C:\Users\bmlzeller\OneDrive - BGO Holding GmbH\Dokumente\BGO_Holding\Vorlagen_BGO\NOS\"
    
    ' Make sure the base folder ends with a backslash
    If Right(baseFolder, 1) <> "\" Then
        baseFolder = baseFolder & "\"
    End If
    
    year2025Folder = baseFolder & "2025\"
    
    ' Create 2025 folder if it doesn't exist
    On Error Resume Next
    If Not fso.FolderExists(year2025Folder) Then
        fso.CreateFolder year2025Folder
    End If
    If Err.Number <> 0 Then
        MsgBox "Error creating folder: " & year2025Folder & vbCrLf & "Error: " & Err.Description, vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Calculate the first Monday of week 1, 2025 using ISO week dating
    ' According to ISO 8601, week 1 is the week with the first Thursday of the year
    ' For 2025, week 1 starts on Monday, December 30, 2024
    mondayDate = DateSerial(2024, 12, 30)
    
    ' Get the correct number of weeks in 2025 (usually 52, but can be 53)
    Dim totalWeeks As Integer
    totalWeeks = ISOWeeksInYear(2025)
    
    ' Display a message to the user to confirm
    If MsgBox("This will create " & totalWeeks & " weekly files for 2025 in:" & vbCrLf & _
              year2025Folder & vbCrLf & vbCrLf & _
              "First week starts on " & Format(mondayDate, "dd.mm.yyyy") & vbCrLf & _
              "Do you want to continue?", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    ' Initialize counters
    successCount = 0
    failCount = 0
    
    ' Add delay between save operations
    Application.ScreenUpdating = False
    
    ' Create files for each week
    For weekNum = 1 To totalWeeks
        ' Format the week number with leading zero if needed
        Dim weekStr As String
        If weekNum < 10 Then
            weekStr = "0" & weekNum
        Else
            weekStr = weekNum
        End If
        
        ' Define the output file path
        Dim outputPath As String
        outputPath = year2025Folder & "KW" & weekStr & ".xlsm"
        
        ' Check if file already exists
        Dim fileExists As Boolean
        fileExists = fso.fileExists(outputPath)
        
        ' Skip file if it exists (to avoid errors)
        If fileExists Then
            Debug.Print "File already exists: " & outputPath & " - Skipping"
            GoTo NextWeek
        End If
        
        ' Create a new workbook
        Set wb = Workbooks.Add
        
        ' Define the sheet order
        Dim sheetOrder(1 To 6) As String
        sheetOrder(1) = "Tourenplan Dispo"
        sheetOrder(2) = "Wochenplan SC Graz"
        sheetOrder(3) = "Wochenplan SC Klagenfurt"
        sheetOrder(4) = "Wochenplan SC Innsbruck"
        sheetOrder(5) = "Wochenplan SC Bregenz"
        sheetOrder(6) = "SYWA-KW12_fein"
        
        ' First, copy all the needed sheets from the template
        Application.DisplayAlerts = False
        For i = 1 To UBound(sheetOrder)
            On Error Resume Next
            tempWb.Sheets(sheetOrder(i)).Copy After:=wb.Sheets(wb.Sheets.count)
            
            ' If sheet doesn't exist in template, create a blank one
            If Err.Number <> 0 Then
                wb.Sheets.Add After:=wb.Sheets(wb.Sheets.count)
                wb.Sheets(wb.Sheets.count).Name = sheetOrder(i)
                Err.Clear
            End If
            On Error GoTo 0
        Next i
        
        ' Handle default sheets (Sheet1, Sheet2, Sheet3, Tabelle1)
        ' Move Tabelle1 to the end if it exists, otherwise delete default sheets
        Dim defaultSheetFound As Boolean
        defaultSheetFound = False
        
        For Each ws In wb.Sheets
            If ws.Name = "Tabelle1" Then
                ' Move Tabelle1 to the end
                ws.Move After:=wb.Sheets(wb.Sheets.count)
                defaultSheetFound = True
            ElseIf ws.Name = "Sheet1" Or ws.Name = "Sheet2" Or ws.Name = "Sheet3" Then
                ' Delete other default sheets
                ws.Delete
            End If
        Next ws
        
        Application.DisplayAlerts = True
        
        ' Set the sheet to work with
        On Error Resume Next
        Set ws = wb.Sheets("Tourenplan Dispo")
        If ws Is Nothing Then
            Set ws = wb.Sheets(1)
        End If
        On Error GoTo 0
        
        ' Update the Monday date in cell B2
        ws.Range("B2").value = mondayDate
        
        ' Update the KW number in the header
        Dim yearForHeader As Integer
        If year(mondayDate) = 2024 Then
            yearForHeader = 2025
        Else
            yearForHeader = year(mondayDate)
        End If
        ws.Range("M1").value = "KW " & weekStr & " / " & yearForHeader
        
        ' Activate the first sheet (Tourenplan Dispo)
        wb.Sheets("Tourenplan Dispo").Activate
        
        ' Try to save the file with enhanced error handling
        On Error Resume Next
        
        ' Make sure the file doesn't exist before trying to save
        If fso.fileExists(outputPath) Then
            On Error Resume Next
            fso.DeleteFile outputPath, True
            Err.Clear
            On Error GoTo 0
        End If
        
        ' Try saving with different methods
        wb.SaveAs Filename:=outputPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        
        If Err.Number <> 0 Then
            ' If first method fails, try alternative saving method
            Err.Clear
            
            ' Try alternative method
            Dim altPath As String
            altPath = year2025Folder & "KW" & weekStr & "_temp.xlsm"
            wb.SaveAs Filename:=altPath, FileFormat:=52 ' Use file format integer code
            
            If Err.Number = 0 Then
                ' If temp save succeeded, try to rename the file
                On Error Resume Next
                Name altPath As outputPath
                If Err.Number <> 0 Then
                    ' If rename fails, at least we have the temp file
                    Debug.Print "Created file with alternate name: " & altPath
                    successCount = successCount + 1
                Else
                    Debug.Print "Successfully created file: " & outputPath
                    successCount = successCount + 1
                End If
            Else
                Debug.Print "Failed to create file: " & outputPath & " - Error: " & Err.Description
                failCount = failCount + 1
            End If
        Else
            Debug.Print "Successfully created file: " & outputPath
            successCount = successCount + 1
        End If
        On Error GoTo 0
        
        ' Close the workbook without saving (already saved)
        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0
        
        ' Add a small delay between saves (may help with network issues)
        Application.Wait Now + TimeSerial(0, 0, 1)
        
NextWeek:
        ' Move to next Monday
        mondayDate = mondayDate + 7
        
        ' Update status to user every 10 files
        If weekNum Mod 10 = 0 Then
            Application.StatusBar = "Created " & weekNum & " of " & totalWeeks & " files..."
        End If
    Next weekNum
    
    ' Clean up
    Set ws = Nothing
    Set wb = Nothing
    Set fso = Nothing
    
    ' Reset Excel UI
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    ' Show results
    MsgBox "Process completed!" & vbCrLf & _
           "Successfully created: " & successCount & " files" & vbCrLf & _
           "Failed: " & failCount & " files", vbInformation
End Sub

Function ISOWeeksInYear(ByVal year As Integer) As Integer
    ' Function to calculate the number of ISO weeks in a given year
    ' A year has 53 weeks if:
    ' 1. The year starts on a Thursday, or
    ' 2. The year starts on a Wednesday in a leap year
    
    Dim jan1 As Date
    jan1 = DateSerial(year, 1, 1)
    
    Select Case Weekday(jan1, vbMonday)
        Case 4 ' Thursday
            ISOWeeksInYear = 53
        Case 3 ' Wednesday
            If IsLeapYear(year) Then
                ISOWeeksInYear = 53
            Else
                ISOWeeksInYear = 52
            End If
        Case Else
            ISOWeeksInYear = 52
    End Select
End Function

Function IsLeapYear(ByVal year As Integer) As Boolean
    ' Check if the given year is a leap year
    IsLeapYear = (year Mod 4 = 0 And year Mod 100 <> 0) Or (year Mod 400 = 0)
End Function

