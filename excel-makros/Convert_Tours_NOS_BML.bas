Sub Convert_Tours_NOS_BML()
    ' Purpose: Convert data from NOS_Tourenkonzept sheet to Tourenplan_BML_* sheets
    ' Using existing template worksheets instead of creating new files
    
    Dim sourceWb As Workbook
    Dim sourceWs As Worksheet
    Dim destWs As Worksheet
    Dim fso As Object
    Dim baseFolder As String
    Dim i As Integer, j As Integer
    Dim areaName As Variant
    Dim tourNumber As String
    Dim tourName As String
    Dim tourTime As String
    Dim containerCount As Integer
    Dim workersField As String
    Dim hasNochPlatz As Boolean
    Dim isDirectContainer As Boolean
    Dim containerNumbers() As String
    Dim sourceRow As Integer
    Dim destRow As Integer
    Dim destCol As Integer
    Dim dayOffset As Integer
    Dim areaCode As Variant
    Dim weekDay As Integer
    Dim containerIndex As Integer
    Dim currentKW As Integer
    Dim currentYear As Integer
    Dim strDate As String
    Dim mondayDate As Date
    
    ' Create areas array
    Dim areas As Variant
    areas = Array("Graz", "Linz", "Klagenfurt", "Innsbruck", "Wiener Neudorf", "Wien")
    
    ' Create area code mapping
    Dim areaCodes As Object
    Set areaCodes = CreateObject("Scripting.Dictionary")
    areaCodes.Add "Innsbruck", "15"
    areaCodes.Add "Wien", "11"
    areaCodes.Add "Wiener Neudorf", "12"  ' "Neudörfel" changed to "Wiener Neudorf"
    areaCodes.Add "Linz", "13"
    areaCodes.Add "Salzburg", "14"
    areaCodes.Add "Kärnten", "16"
    areaCodes.Add "Klagenfurt", "16"    ' Alias for Kärnten
    areaCodes.Add "Graz", "17"
    areaCodes.Add "Dornbirn", "18"
    areaCodes.Add "Deutschland", "43"
    areaCodes.Add "Budapest", "45"
    areaCodes.Add "Bratislava", "47"
    
    ' Create FileSystemObject for file operations
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Set source workbook as the active workbook
    Set sourceWb = ActiveWorkbook
    
    ' Get source worksheet (the NOS_Tourenkonzept sheet)
    On Error Resume Next
    Set sourceWs = sourceWb.Sheets("NOS_Tourenkonzept")
    If sourceWs Is Nothing Then
        MsgBox "Could not find the NOS_Tourenkonzept sheet. Please check the workbook structure.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Get current KW (calendar week) from user
    currentKW = InputBox("Enter the calendar week number (KW):", "Calendar Week", Format(Date, "ww"))
    If currentKW <= 0 Or currentKW > 53 Then
        MsgBox "Invalid calendar week number. Please enter a value between 1 and 53.", vbExclamation
        Exit Sub
    End If
    
    ' Get the current year or ask user
    currentYear = InputBox("Enter the year:", "Year", Year(Date))
    If currentYear < 2024 Or currentYear > 2030 Then
        MsgBox "Invalid year. Please enter a value between 2024 and 2030.", vbExclamation
        Exit Sub
    End If
    
    ' Calculate the first day (Monday) of the given calendar week
    mondayDate = GetMondayOfWeek(currentYear, currentKW)
    
    ' Define base folder for saving files (or ask user)
    baseFolder = GetFolderPath()
    If baseFolder = "" Then Exit Sub
    
    ' Make sure the folder ends with a backslash
    If Right(baseFolder, 1) <> "\" Then
        baseFolder = baseFolder & "\"
    End If
    
    ' Verify source worksheet has the expected structure
    If Not ValidateSourceWorksheet(sourceWs) Then
        MsgBox "The source worksheet does not have the expected structure. Please check the format.", vbExclamation
        Exit Sub
    End If
    
    ' Process each area
    For Each areaName In areas
        ' Skip this iteration if the area template doesn't exist
        On Error Resume Next
        Dim templateName As String
        templateName = "Tourenplan_BML_" & areaName
        Set destWs = sourceWb.Sheets(templateName)
        If destWs Is Nothing Then
            ' Template doesn't exist, log and continue to next area
            Debug.Print "Skipping " & areaName & " - template not found."
            On Error GoTo 0
            GoTo NextArea
        End If
        On Error GoTo 0
        
        ' Reset container index for each area
        Dim areaContainerIndices As Object
        Set areaContainerIndices = CreateObject("Scripting.Dictionary")
        
        ' Clear previous data in destination worksheet (only the tour areas)
        ClearDestinationData destWs
        
        ' Update header and dates
        UpdateHeaderAndDates destWs, mondayDate, currentKW, currentYear
        
        ' Counter for direct containers for each day
        Dim directContainerCounts(5) As Integer  ' 0-indexed for 5 days (Monday to Friday)
        For i = 0 To 4
            directContainerCounts(i) = 0
        Next i
        
        ' Process source data for this area
        For sourceRow = 2 To 70 Step 3  ' Assume tours are in groups of 3 rows
            ' Loop through each day (Monday to Friday)
            For dayOffset = 0 To 4
                ' Calculate column positions for the current day
                Dim sourceStartCol As Integer
                sourceStartCol = 2 + (dayOffset * 4)  ' Starting column for current day (B, F, J, N, R)
                
                ' Extract tour information
                On Error Resume Next  ' In case of empty cells
                tourNumber = Trim(sourceWs.Cells(sourceRow, sourceStartCol).Value)
                tourName = Trim(sourceWs.Cells(sourceRow + 1, sourceStartCol).Value)
                tourTime = Trim(sourceWs.Cells(sourceRow + 2, sourceStartCol).Value)
                
                ' Extract container count (only in source file)
                containerCount = Val(sourceWs.Cells(sourceRow + 2, sourceStartCol + 1).Value)
                If containerCount = 0 Then containerCount = 1  ' Default to 1 if not specified
                
                ' Extract workers field
                workersField = Trim(sourceWs.Cells(sourceRow, sourceStartCol + 3).Value)
                
                ' Check if "Noch Platz" is marked
                hasNochPlatz = IsCheckboxChecked(sourceWs.Cells(sourceRow + 1, sourceStartCol + 3))
                
                ' Check if this is a direct container
                isDirectContainer = IsCheckboxChecked(sourceWs.Cells(sourceRow + 2, sourceStartCol + 3))
                On Error GoTo 0
                
                ' Skip if no tour number or name (empty tour)
                If tourNumber = "" And tourName = "" Then GoTo NextDay
                
                ' Skip if this tour doesn't belong to the current area
                If Not BelongsToArea(tourName, areaName) Then GoTo NextDay
                
                ' Set the weekday (1-5 for Monday-Friday)
                weekDay = dayOffset + 1
                
                ' Get area code (or default to 88 if not found)
                areaCode = GetAreaCode(tourName, areaCodes)
                
                ' For each container, generate a container number and add to destination
                For containerIndex = 1 To containerCount
                    ' Update the container counter for this area and day
                    Dim areaWeekdayKey As String
                    areaWeekdayKey = areaCode & "_" & weekDay
                    
                    If Not areaContainerIndices.Exists(areaWeekdayKey) Then
                        areaContainerIndices.Add areaWeekdayKey, 0
                    End If
                    
                    areaContainerIndices(areaWeekdayKey) = areaContainerIndices(areaWeekdayKey) + 1
                    
                    ' Create the container number
                    Dim containerNumber As String
                    containerNumber = areaCode & "_" & weekDay & "_" & Format(areaContainerIndices(areaWeekdayKey), "00")
                    
                    ' Calculate destination row and column
                    destCol = 2 + (dayOffset * 3)  ' B, E, H, K, N, Q for dates
                    
                    ' If direct container and we have space (max 3 per day), place in direct container section
                    If isDirectContainer And directContainerCounts(dayOffset) < 3 Then
                        destRow = 25 + (directContainerCounts(dayOffset) * 3)
                        directContainerCounts(dayOffset) = directContainerCounts(dayOffset) + 1
                    Else
                        ' Find the next available tour row in the upper section
                        destRow = FindNextAvailableTourRow(destWs, destCol, 3, 23)
                        
                        ' If no space left, skip this container
                        If destRow = -1 Then GoTo NextContainer
                    End If
                    
                    ' Add tour name (with "Noch Platz" if applicable)
                    Dim finalTourName As String
                    finalTourName = tourName
                    If hasNochPlatz Then finalTourName = finalTourName & " - Noch Platz"
                    
                    ' Now add to the destination worksheet
                    AddTourToDestination destWs, destRow, destCol, finalTourName, tourNumber, containerNumber, tourTime, workersField
                    
NextContainer:
                Next containerIndex
                
NextDay:
            Next dayOffset
        Next sourceRow
        
        ' Create a new workbook for this area by copying the template sheets
        Dim newWb As Workbook
        Set newWb = Workbooks.Add
        
        ' Copy the Tourenplan_BML sheet from template to new workbook
        destWs.Copy Before:=newWb.Sheets(1)
        
        ' Try to find and copy the Wochenplan_BML sheet too
        On Error Resume Next
        Dim wochenplanWs As Worksheet
        Set wochenplanWs = sourceWb.Sheets("Wochenplan_BML")
        If Not wochenplanWs Is Nothing Then
            wochenplanWs.Copy After:=newWb.Sheets(newWb.Sheets.Count)
        End If
        On Error GoTo 0
        
        ' Delete any default sheets
        Application.DisplayAlerts = False
        For i = newWb.Sheets.Count To 1 Step -1
            If newWb.Sheets(i).Name Like "Sheet*" Or newWb.Sheets(i).Name Like "Tabelle*" Then
                newWb.Sheets(i).Delete
            End If
        Next i
        Application.DisplayAlerts = True
        
        ' Save the workbook with the appropriate name
        Dim destFileName As String
        destFileName = baseFolder & "Tourenplan_BML_NOS_KW" & Format(currentKW, "00") & "_" & areaName & ".xlsm"
        
        ' Try to save the file with alternatives if the first attempt fails
        On Error Resume Next
        
        ' Attempt 1: Try standard SaveAs
        newWb.SaveAs Filename:=destFileName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        
        ' Check if we had an error
        If Err.Number <> 0 Then
            Err.Clear
            
            ' Attempt 2: Try using numeric file format value
            newWb.SaveAs Filename:=destFileName, FileFormat:=52  ' 52 = xlOpenXMLWorkbookMacroEnabled
            
            ' Check if we still have an error
            If Err.Number <> 0 Then
                MsgBox "Error saving file: " & destFileName & vbCrLf & "Error: " & Err.Description, vbExclamation
            End If
        End If
        On Error GoTo 0
        
        ' Close the workbook
        newWb.Close SaveChanges:=False
        
        ' Notify user
        Debug.Print "Created file for " & areaName & ": " & destFileName
        
NextArea:
    Next areaName
    
    ' Clean up
    Set fso = Nothing
    Set areaCodes = Nothing
    Set areaContainerIndices = Nothing
    
    MsgBox "Conversion completed successfully!", vbInformation
End Sub

Function ValidateSourceWorksheet(ws As Worksheet) As Boolean
    ' Check if the source worksheet has the expected structure
    ' For simplicity, we'll just check if there's data in expected cells
    
    ' Check column headers or other identifiers
    On Error Resume Next
    
    ' Perform some basic checks (modify as needed)
    If ws.Cells(2, 2).Value = "" And ws.Cells(2, 6).Value = "" Then
        ValidateSourceWorksheet = False
        Exit Function
    End If
    
    On Error GoTo 0
    ValidateSourceWorksheet = True
End Function

Function GetFolderPath() As String
    ' Ask user for folder path
    Dim folderPath As String
    folderPath = InputBox("Enter the folder path to save the files:", "Save Location", "C:\Temp\")
    
    ' Create folder if it doesn't exist
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If folderPath <> "" And Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
        If Err.Number <> 0 Then
            MsgBox "Error creating folder: " & folderPath & vbCrLf & "Error: " & Err.Description, vbExclamation
            folderPath = ""
        End If
    End If
    
    On Error GoTo 0
    Set fso = Nothing
    GetFolderPath = folderPath
End Function

Sub UpdateHeaderAndDates(ws As Worksheet, mondayDate As Date, kwNumber As Integer, yearNumber As Integer)
    ' Update the header and dates in the destination worksheet
    
    ' Set header
    ws.Range("M1").Value = "KW " & Format(kwNumber, "00") & " / " & yearNumber
    
    ' Set date row (row 2)
    Dim i As Integer
    For i = 0 To 4
        Dim currentDate As Date
        currentDate = mondayDate + i
        
        ' Set date in columns B, E, H, K, N, Q
        ws.Cells(2, 2 + (i * 3)).Value = currentDate
        ws.Cells(2, 2 + (i * 3)).NumberFormat = "dd.mm.yyyy"
    Next i
End Sub

Sub ClearDestinationData(ws As Worksheet)
    ' Clear only the tour data areas of the destination worksheet
    ' Upper tour section (rows 3-23)
    ws.Range("B3:T23").ClearContents
    
    ' Direct container section (rows 25-33)
    ws.Range("B25:T33").ClearContents
    
    ' Personnel count (row 34)
    ws.Range("D34:T34").ClearContents
    
    ' Keep row/column formatting intact
End Sub

Function BelongsToArea(ByVal tourName As String, ByVal areaName As Variant) As Boolean
    ' Check if a tour belongs to the specified area
    ' This is a simplified check - modify as needed
    Dim strAreaName As String
    
    ' Convert variant to string
    strAreaName = CStr(areaName)
    
    ' Check if the area name appears in the tour name
    If InStr(1, tourName, strAreaName, vbTextCompare) > 0 Then
        BelongsToArea = True
        Exit Function
    End If
    
    ' Special cases
    Select Case strAreaName
        Case "Klagenfurt"
            If InStr(1, tourName, "Kärnten", vbTextCompare) > 0 Then
                BelongsToArea = True
                Exit Function
            End If
        
        Case "Wiener Neudorf"
            If InStr(1, tourName, "Neudörf", vbTextCompare) > 0 Then
                BelongsToArea = True
                Exit Function
            End If
        
        ' Add other special cases as needed
    End Select
    
    BelongsToArea = False
End Function

Function GetAreaCode(ByVal tourName As String, ByVal areaCodes As Object) As String
    ' Get the area code based on tour name
    Dim areaKey As Variant
    
    ' Check all area names from the dictionary
    For Each areaKey In areaCodes.Keys
        If InStr(1, CStr(tourName), CStr(areaKey), vbTextCompare) > 0 Then
            GetAreaCode = CStr(areaCodes(areaKey))
            Exit Function
        End If
    Next areaKey
    
    ' Default if no match
    GetAreaCode = "88"
End Function

Function FindNextAvailableTourRow(ws As Worksheet, col As Integer, startRow As Integer, endRow As Integer) As Integer
    ' Find the next available row for a tour
    Dim i As Integer
    
    ' Check for rows in groups of 3 (one tour takes 3 rows)
    For i = startRow To endRow Step 3
        ' Check if the tour name field is empty
        If Trim(ws.Cells(i, col).Value) = "" Then
            FindNextAvailableTourRow = i
            Exit Function
        End If
    Next i
    
    ' No available row found
    FindNextAvailableTourRow = -1
End Function

Sub AddTourToDestination(ws As Worksheet, row As Integer, col As Integer, tourName As String, tourNumber As String, containerNumber As String, tourTime As String, workersField As String)
    ' Add a tour to the destination worksheet
    
    ' Clear the cells first (in case of previous data)
    ws.Range(ws.Cells(row, col), ws.Cells(row + 2, col + 2)).ClearContents
    
    ' 1. Tour name in the top row (merged cells)
    ws.Cells(row, col).Value = tourName
    
    ' Merge the tour name cells if not already merged
    On Error Resume Next
    If Not ws.Cells(row, col).MergeCells Then
        ws.Range(ws.Cells(row, col), ws.Cells(row, col + 2)).MergeCells = True
    End If
    On Error GoTo 0
    
    ' 2. Tour number in the second row
    ws.Cells(row + 1, col).Value = tourNumber
    
    ' 3. Container number in the second row, last column
    ws.Cells(row + 1, col + 2).Value = containerNumber
    
    ' 4. Workers field in the third row
    ws.Cells(row + 2, col).Value = workersField
    
    ' 5. Tour time in the third row, last column
    ws.Cells(row + 2, col + 2).Value = tourTime
    
    ' 6. Add the fraction field to the left of the container number
    ws.Cells(row + 1, col + 1).Value = "1/1"  ' Default to 1/1 for now
    
    ' Add any additional formatting or content as needed
End Sub

Function IsCheckboxChecked(cell As Range) As Boolean
    ' Check if a cell contains a checked checkbox
    ' This is a simplified check and may need adjustment based on how checkboxes are implemented
    
    On Error Resume Next
    
    ' Simplify the checkbox detection since we're getting errors
    ' Let's just check the cell value
    
    ' Option 1: Check for TRUE/FALSE values
    If TypeName(cell.Value) = "Boolean" Then
        IsCheckboxChecked = cell.Value
        Exit Function
    End If
    
    ' Option 2: Check for 0/1 values (common for form controls)
    If IsNumeric(cell.Value) Then
        IsCheckboxChecked = (cell.Value = 1)
        Exit Function
    End If
    
    ' Option 3: Check for text indicators
    If TypeName(cell.Value) = "String" Then
        Dim cellText As String
        cellText = LCase(Trim(cell.Value))
        
        If cellText = "true" Or cellText = "ja" Or cellText = "yes" Or _
           cellText = "x" Or cellText = "?" Or cellText = "?" Then
            IsCheckboxChecked = True
            Exit Function
        End If
    End If
    
    ' Option 4: Look for a shape in the cell
    Dim targetAddress As String
    targetAddress = cell.Address
    
    ' Check for form checkboxes
    On Error Resume Next
    Dim chkObj As Object
    For Each chkObj In cell.Worksheet.CheckBoxes
        If Not chkObj Is Nothing Then
            If chkObj.TopLeftCell.Address = targetAddress Then
                IsCheckboxChecked = (chkObj.Value = 1) Or (chkObj.Value = True)
                Exit Function
            End If
        End If
    Next chkObj
    
    ' The most reliable way is usually just to ask the user (but we'll return False as default)
    On Error GoTo 0
    IsCheckboxChecked = False
End Function

Function GetMondayOfWeek(yearNum As Integer, weekNum As Integer) As Date
    ' Get the Monday date of a specific ISO week and year
    Dim janFirst As Date
    Dim dayOfWeek As Integer
    Dim daysToAdd As Integer
    
    ' January 1st of the given year
    janFirst = DateSerial(yearNum, 1, 1)
    
    ' Day of week for January 1st (1 = Monday, 7 = Sunday)
    dayOfWeek = weekDay(janFirst, vbMonday)
    
    ' Calculate days to add to get to the first Monday of the year
    If dayOfWeek = 1 Then
        daysToAdd = 0  ' Jan 1 is already a Monday
    Else
        daysToAdd = 9 - dayOfWeek  ' Days until first Monday
    End If
    
    ' Monday of week 1
    Dim week1Monday As Date
    week1Monday = janFirst + daysToAdd
    
    ' Add weeks to get to the desired week
    GetMondayOfWeek = week1Monday + (7 * (weekNum - 1))
    
    ' Special handling for weeks that belong to the previous year
    If weekNum = 1 And Month(week1Monday) = 12 Then
        ' Week 1 belongs to the following year, so we need the Monday of ISO week 1
        GetMondayOfWeek = week1Monday
    End If
End Function
