Sub SetupMainControls()
    ' Create the main controls (dropdown and all buttons)
    
    ' First, set up the Gebiet dropdown
    SetupGebietDropdown
    
    ' Then set up the KW switcher (button only, no dropdown)
    SetupKWSwitcher
    
    ' Add the additional buttons
    AddActionButtons
    
    MsgBox "Steuerelemente wurden erstellt. Klicken Sie auf 'KW wechseln' zum Ändern der Kalenderwoche."
End Sub

Sub SetupGebietDropdown()
    ' Create a dropdown in W1
    Dim sheetList As String
    Dim i As Integer
    
    ' Build list of sheets
    sheetList = ""
    For i = 1 To ThisWorkbook.Sheets.count
        If ThisWorkbook.Sheets(i).Name Like "Tourenplan_BML_*" Then
            sheetList = sheetList & ThisWorkbook.Sheets(i).Name & ","
        End If
    Next i
    
    ' Remove trailing comma
    If Len(sheetList) > 0 Then
        sheetList = Left(sheetList, Len(sheetList) - 1)
    End If
    
    ' Create dropdown
    With ActiveSheet.Range("W1").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:=sheetList
        .InCellDropdown = True
    End With
    
    ' Add label
    ActiveSheet.Range("V1").Value = "Gebiet:"
End Sub

Sub SetupKWSwitcher()
    ' Create a KW switcher control (just a button, no dropdown)
    
    ' Add the button
    On Error Resume Next
    Dim btnKW As Button
    Set btnKW = ActiveSheet.Buttons.Add(ActiveSheet.Range("X1").Left - 80, _
                                       ActiveSheet.Range("X1").Top, 80, 20)
    With btnKW
        .Caption = "KW wechseln"
        .OnAction = "SwitchKW"
        .Name = "btnSwitchKW"
    End With
    On Error GoTo 0
End Sub

Sub AddActionButtons()
    ' Add all action buttons
    
    ' Main data load button
    On Error Resume Next
    Dim btnLoad As Button
    Set btnLoad = ActiveSheet.Buttons.Add(ActiveSheet.Range("V1").Left, _
                                         ActiveSheet.Range("V1").Top + 30, 120, 20)
    With btnLoad
        .Caption = "Daten laden"
        .OnAction = "SimpleTransferWithDates"
        .Name = "btnLoadData"
    End With
    
    ' Transfer Tours button
    Dim btnTours As Button
    Set btnTours = ActiveSheet.Buttons.Add(ActiveSheet.Range("V1").Left + 130, _
                                          ActiveSheet.Range("V1").Top + 30, 120, 20)
    With btnTours
        .Caption = "Transfer Tours"
        .OnAction = "TransferTours"
        .Name = "btnTransferTours"
    End With
    
    ' Transfer WABs button
    Dim btnWABs As Button
    Set btnWABs = ActiveSheet.Buttons.Add(ActiveSheet.Range("V1").Left + 260, _
                                         ActiveSheet.Range("V1").Top + 30, 120, 20)
    With btnWABs
        .Caption = "Transfer WABs"
        .OnAction = "TransferWABs"
        .Name = "btnTransferWABs"
    End With
    
    ' Print button
    Dim btnPrint As Button
    Set btnPrint = ActiveSheet.Buttons.Add(ActiveSheet.Range("V1").Left + 390, _
                                          ActiveSheet.Range("V1").Top + 30, 120, 20)
    With btnPrint
        .Caption = "Drucken"
        .OnAction = "PrintWorksheet"
        .Name = "btnPrint"
    End With
    On Error GoTo 0
End Sub

Sub PrintWorksheet()
    ' Print the NOS_Tourenkonzept_Print worksheet
    
    ' Save the currently active sheet to return to later
    Dim currentSheet As Worksheet
    Set currentSheet = ActiveSheet
    
    ' Check if the print worksheet exists
    Dim printSheet As Worksheet
    On Error Resume Next
    Set printSheet = ThisWorkbook.Worksheets("NOS_Tourenkonzept_Print")
    
    If printSheet Is Nothing Then
        MsgBox "Das Druckblatt 'NOS_Tourenkonzept_Print' wurde nicht gefunden.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Activate the print sheet
    printSheet.Activate
    
    ' Display the print dialog
    Application.Dialogs(xlDialogPrint).Show
    
    ' Return to the original sheet
    currentSheet.Activate
End Sub

Sub SwitchKW()
    ' Simple KW switcher using fixed 2025 dates
    Dim kwInput As String
    Dim kwNumber As Integer
    Dim firstMonday As Date
    Dim currentKW As Integer
    
    ' Get the current KW based on B1 (corrected from B2)
    On Error Resume Next
    If IsDate(ActiveSheet.Range("B1").Value) Then
        ' Get KW based on the date in B1
        currentKW = GetKW2025FromDate(CDate(ActiveSheet.Range("B1").Value))
    Else
        currentKW = GetKW2025FromDate(Date)
    End If
    On Error GoTo 0
    
    ' Ask for the desired KW using an InputBox
    kwInput = InputBox("Bitte geben Sie die gewünschte Kalenderwoche ein (1-53):" & vbCrLf & _
                       "Aktuelle KW: " & currentKW, "KW wechseln", currentKW)
    
    ' Validate input
    If kwInput = "" Then
        ' User clicked Cancel
        Exit Sub
    End If
    
    If Not IsNumeric(kwInput) Then
        MsgBox "Bitte geben Sie eine gültige Zahl ein.", vbExclamation
        Exit Sub
    End If
    
    kwNumber = CInt(kwInput)
    
    If kwNumber < 1 Or kwNumber > 53 Then
        MsgBox "Bitte geben Sie eine Zahl zwischen 1 und 53 ein.", vbExclamation
        Exit Sub
    End If
    
    ' Check if we're already in the selected KW
    If kwNumber = currentKW Then
        MsgBox "Sie befinden sich bereits in KW" & kwNumber & ".", vbInformation
        Exit Sub
    End If
    
    ' Get the Monday date for the selected KW from our fixed calendar
    firstMonday = GetDateForKW2025(kwNumber)
    
    ' Save the current file before switching
    If SaveWorkbookToSpecialFolder(currentKW) Then
        ' Successfully saved, proceed with switch
        
        ' First clear all data before changing the date
        ClearAllData
        
        ' Now update B1 with the first Monday date of selected KW
        ActiveSheet.Range("B1").Value = firstMonday
        
        ' Update the KW label in the sheet (cell W2)
        ActiveSheet.Range("X1").Value = "Tourenplan für KW" & kwNumber & " (" & _
                                      Format(firstMonday, "dd.mm.yyyy") & " - " & _
                                      Format(firstMonday + 4, "dd.mm.yyyy") & ")"
        
        MsgBox "Erfolgreich zu KW" & kwNumber & " gewechselt. Das Datum in B1 wurde auf " & _
               Format(firstMonday, "dd.mm.yyyy") & " gesetzt.", vbInformation
    Else
        ' File could not be saved, don't switch
        MsgBox "KW-Wechsel abgebrochen, da die aktuelle Datei nicht gespeichert werden konnte.", vbExclamation
    End If
End Sub

Function GetDateForKW2025(kwNumber As Integer) As Date
    ' Returns the Monday date for a given KW in 2025
    ' Using the simple formula: KW1_Monday + (kwNumber-1) * 7
    
    ' KW1 Monday for 2025 is Dec 30, 2024
    Dim kw1Monday As Date
    kw1Monday = DateSerial(2024, 12, 30)
    
    ' Calculate the Monday for the requested KW
    GetDateForKW2025 = kw1Monday + ((kwNumber - 1) * 7)
End Function

Function GetKW2025FromDate(targetDate As Date) As Integer
    ' Calculate which KW a date belongs to based on days since KW1 start
    ' This is much simpler and more reliable than the previous approach
    
    ' KW1 Monday for 2025 is Dec 30, 2024
    Dim kw1Monday As Date
    kw1Monday = DateSerial(2024, 12, 30)
    
    ' Handle dates before KW1
    If targetDate < kw1Monday Then
        GetKW2025FromDate = 1
        Exit Function
    End If
    
    ' Calculate days since KW1 start
    Dim daysSinceKW1 As Integer
    daysSinceKW1 = targetDate - kw1Monday
    
    ' Calculate KW number (integer division by 7 and add 1)
    Dim kwNumber As Integer
    kwNumber = (daysSinceKW1 \ 7) + 1
    
    ' Cap at KW53
    If kwNumber > 53 Then kwNumber = 53
    
    GetKW2025FromDate = kwNumber
End Function

Function SaveWorkbookToSpecialFolder(currentKW As Integer) As Boolean
    ' Save with shorter filename to OneDrive path
    On Error GoTo ErrorHandler
    
    ' Get the current user's profile path and username
    Dim userProfile As String
    Dim userName As String
    userProfile = Environ("USERPROFILE")
    userName = Environ("USERNAME")
    
    ' Setup the target folder path with the specified OneDrive path
    Dim targetFolder As String
    targetFolder = userProfile & "\OneDrive - BGO Holding GmbH\BML_Dispo - Planung NOS - Planung NOS\10_Excel_Wocheneinteilung_Intern_NOS\Autosave"
    
    ' Create the folder if it doesn't exist
    On Error Resume Next
    If Dir(targetFolder, vbDirectory) = "" Then
        MkDir targetFolder
    End If
    On Error GoTo ErrorHandler
    
    ' Generate timestamp with shorter format
    Dim timeStamp As String
    timeStamp = Format(Now, "dd_MM_HH_mm")
    
    ' Generate shorter filename with KW and username
    Dim newFileName As String
    newFileName = targetFolder & "\NOS_TK_" & userName & "_" & timeStamp & ".xlsx"
    
    ' Use the simple SaveCopyAs approach which is more reliable
    ThisWorkbook.SaveCopyAs newFileName
    
    MsgBox "Datei erfolgreich gespeichert unter:" & vbCrLf & newFileName, vbInformation
    
    SaveWorkbookToSpecialFolder = True
    Exit Function
    
ErrorHandler:
    MsgBox "Fehler beim Speichern: " & Err.Description & vbCrLf & _
           "Versuchter Pfad: " & targetFolder, vbExclamation
    SaveWorkbookToSpecialFolder = False
End Function

Sub ClearAllData()
    ' Clear all data fields in main sheet and subsheets as specified
    Application.ScreenUpdating = False
    
    ' Clear main sheet (NOS_Tourenkonzept)
    Dim nosSheet As Worksheet
    On Error Resume Next
    Set nosSheet = ThisWorkbook.Worksheets("NOS_Tourenkonzept")
    If Not nosSheet Is Nothing Then
        ' Monday (B & C columns)
        ClearRangeAndCheckboxes nosSheet, "B2:C100", "E2:E100"
        
        ' Clear the specific cells in E column that you mentioned
        ClearSpecificCells nosSheet, "E", Array(2, 5, 8, 11, 14, 17, 20, 23, 26, 29, 32, 35, 38, 41, 44)
        
        ' Tuesday (F & G columns)
        ClearRangeAndCheckboxes nosSheet, "F2:G100", "I2:I100"
        
        ' Clear the specific cells in I column
        ClearSpecificCells nosSheet, "I", Array(2, 5, 8, 11, 14, 17, 20, 23, 26, 29, 32, 35, 38, 41, 44)
        
        ' Wednesday (J & K columns)
        ClearRangeAndCheckboxes nosSheet, "J2:K100", "M2:M100"
        
        ' Clear the specific cells in M column
        ClearSpecificCells nosSheet, "M", Array(2, 5, 8, 11, 14, 17, 20, 23, 26, 29, 32, 35, 38, 41, 44)
        
        ' Thursday (N & O columns)
        ClearRangeAndCheckboxes nosSheet, "N2:O100", "Q2:Q100"
        
        ' Clear the specific cells in Q column
        ClearSpecificCells nosSheet, "Q", Array(2, 5, 8, 11, 14, 17, 20, 23, 26, 29, 32, 35, 38, 41, 44)
        
        ' Friday (R & S columns)
        ClearRangeAndCheckboxes nosSheet, "R2:S100", "U2:U100"
        
        ' Clear the specific cells in U column
        ClearSpecificCells nosSheet, "U", Array(2, 5, 8, 11, 14, 17, 20, 23, 26, 29, 32, 35, 38, 41, 44)
    End If
    On Error GoTo 0
    
    ' Clear all subsheets
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "Tourenplan_BML_*" Then
            ' Clear Tour fields (B3:S33) - leave headers and car information intact
            ClearTourFields ws, "B3:S33"
            
            ' Clear WAB fields (B35:S38) - leave headers intact
            ClearWABFields ws, "B35:S38"
        End If
    Next ws
    
    Application.ScreenUpdating = True
End Sub

Sub ClearRangeAndCheckboxes(ws As Worksheet, dataRange As String, checkboxRange As String)
    ' Clear data range and set checkboxes to FALSE
    
    ' Clear content of data range
    ws.Range(dataRange).ClearContents
    
    ' Set checkboxes to FALSE
    Dim cell As Range
    For Each cell In ws.Range(checkboxRange)
        If IsCheckbox(cell) Then
            cell.Value = False
        End If
    Next cell
End Sub

Sub ClearSpecificCells(ws As Worksheet, columnLetter As String, rowNumbers As Variant)
    ' Clear specific cells in a column
    Dim i As Integer
    
    For i = LBound(rowNumbers) To UBound(rowNumbers)
        Dim cellAddress As String
        cellAddress = columnLetter & rowNumbers(i)
        
        ' Clear the cell
        ws.Range(cellAddress).ClearContents
        
        ' Set checkbox to FALSE if it's a checkbox
        If IsCheckbox(ws.Range(cellAddress)) Then
            ws.Range(cellAddress).Value = False
        End If
    Next i
End Sub

Function IsCheckbox(cell As Range) As Boolean
    ' Check if a cell contains a checkbox or similar control
    ' This is a simple check - it assumes checkboxes are already present
    ' and just need to be set to FALSE
    
    On Error Resume Next
    IsCheckbox = (TypeName(cell.Value) = "Boolean")
    On Error GoTo 0
End Function

Sub ClearTourFields(ws As Worksheet, tourRange As String)
    ' Clear Tour fields but preserve formatting and headers
    
    ' Get the range
    Dim rng As Range
    Set rng = ws.Range(tourRange)
    
    ' Clear the content
    rng.ClearContents
    
    ' Remove background colors but keep borders
    Dim cell As Range
    For Each cell In rng
        cell.Interior.colorIndex = xlNone
    Next cell
End Sub

Sub ClearWABFields(ws As Worksheet, wabRange As String)
    ' Clear WAB fields but preserve formatting and headers
    
    ' Get the range
    Dim rng As Range
    Set rng = ws.Range(wabRange)
    
    ' Clear the content
    rng.ClearContents
    
    ' Remove background colors but keep borders
    Dim cell As Range
    For Each cell In rng
        cell.Interior.colorIndex = xlNone
    Next cell
End Sub

Sub SimpleTransferWithDates()
    ' This is the transfer function with date synchronization
    
    ' Store active sheet reference
    Dim mainSheet As Worksheet
    Set mainSheet = ActiveSheet
    
    ' Get selected sheet
    Dim selectedSheet As String
    selectedSheet = mainSheet.Range("W1").Value
    
    If selectedSheet = "" Then
        MsgBox "Bitte wählen Sie ein Gebiet aus."
        Exit Sub
    End If
    
    ' Get source sheet
    Dim sourceWs As Worksheet
    On Error Resume Next
    Set sourceWs = ThisWorkbook.Worksheets(selectedSheet)
    If Err.Number <> 0 Or sourceWs Is Nothing Then
        MsgBox "Blatt nicht gefunden: " & selectedSheet
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Turn off screen updates for smoother operation
    Application.ScreenUpdating = False
    
    ' Clear old data safely
    On Error Resume Next
    mainSheet.Range("W2:AP80").UnMerge
    On Error GoTo 0
    mainSheet.Range("W2:AP80").Clear
    
    ' Copy source data
    sourceWs.Range("A1:S80").Copy
    
    ' Paste to target
    mainSheet.Range("W2").PasteSpecial xlPasteAll
    
    ' Clear clipboard
    Application.CutCopyMode = False
    
    ' Set font size to 10
    mainSheet.Range("W2:AP80").Font.Size = 10
    
    ' Merge SC-Leiter thing
    mainSheet.Range("W2:W3").Merge
    
    ' Get date from B1 (corrected from B2)
    Dim startDate As Date
    If IsDate(mainSheet.Range("B1").Value) Then
        startDate = mainSheet.Range("B1").Value
    Else
        startDate = Date ' Use current date as fallback
    End If

    
    ' Set the title with KW info
    Dim kwNumber As Integer
    kwNumber = GetKW2025FromDate(startDate)
    
    mainSheet.Range("X1").Value = "Tourenplan für KW" & kwNumber & " (" & _
                                Format(startDate, "dd.mm.yyyy") & " - " & _
                                Format(startDate + 4, "dd.mm.yyyy") & ")"
    
    mainSheet.Range("X1").Font.Bold = True
    mainSheet.Range("X1").Font.Size = 11
    
    ' Update date headers (X3, AA3, AD3, AG3, AJ3, AM3)
    For i = 0 To 5 ' Monday to Saturday
        Dim dayDate As Date
        dayDate = startDate + i
        
        ' Format date string
        Dim dateStr As String
        dateStr = Format(dayDate, "dddd, dd.mm.yyyy")
        
        ' Calculate column position (X, AA, AD, etc.)
        Dim colIndex As Integer
        colIndex = 24 + (i * 3) ' X is column 24, each day takes 3 columns
        
        ' Update header if it exists
        mainSheet.Cells(3, colIndex).Value = dateStr
        
        ' Merge cells for the header if needed
        If Not mainSheet.Cells(3, colIndex).mergeCells Then
            mainSheet.Range(mainSheet.Cells(3, colIndex), mainSheet.Cells(3, colIndex + 2)).Merge
        End If
        
        ' Format header
        With mainSheet.Cells(3, colIndex)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .Font.Size = 10
            .Interior.Color = RGB(220, 220, 220)
        End With
    Next i
    On Error GoTo 0
    
    ' Make sure we're back on the main sheet
    mainSheet.Activate
    
    ' Turn screen updating back on
    Application.ScreenUpdating = True
    
    MsgBox "Daten mit aktualisierten Datumsangaben übertragen!"
End Sub

Sub TransferTours()
    ' Transfer tours using row mapping pattern with direct container logic
    ' and color coding as requested
    
    Dim mainWs As Worksheet
    Dim neudorfWs As Worksheet
    Dim mainCarRow As Integer, subCarRow As Integer
    Dim day As Integer, sourceCol As Integer, destCol As Integer
    Dim tourName As String, tourNumber As String, tourTime As String
    Dim workers As String, wabCount As Integer
    Dim hasNochPlatz As Boolean, isDirectContainer As Boolean
    Dim carsProcessed As Integer, toursAdded As Integer
    Dim directContainersUsed(6) As Integer ' Counter for direct containers for each day
    
    ' Dictionary to track LLnr numbers and their colors
    Dim llnrDict As Object
    Set llnrDict = CreateObject("Scripting.Dictionary")
    Dim llnrColorIndex As Integer
    llnrColorIndex = 3 ' Start with color index 3 (red is 3, blue is 5, etc.)
    
    ' Array to track WAB counters for each day
    Dim wabCounters(6) As Integer
    For day = 0 To 6
        wabCounters(day) = 1 ' Start WAB counter at 1 for each day
    Next day
    
    ' Get worksheets
    Set mainWs = ThisWorkbook.Worksheets("NOS_Tourenkonzept")
    
    ' Get selected sheet from dropdown or default to WNeudorf if nothing selected
    Dim selectedSheet As String
    selectedSheet = ThisWorkbook.Sheets("NOS_Tourenkonzept").Range("W1").Value
    
    If selectedSheet = "" Then
        Set neudorfWs = ThisWorkbook.Worksheets("Tourenplan_BML_WNeudorf")
    Else
        On Error Resume Next
        Set neudorfWs = ThisWorkbook.Worksheets(selectedSheet)
        If neudorfWs Is Nothing Then
            Set neudorfWs = ThisWorkbook.Worksheets("Tourenplan_BML_WNeudorf")
        End If
        On Error GoTo 0
    End If
    
    ' Clear tour areas in destination - make sure to clear all potential WAB cells
    neudorfWs.Range("B3:S33").ClearContents
    neudorfWs.Range("B3:S33").Interior.colorIndex = xlNone
    neudorfWs.Range("B55:S72").ClearContents ' Clear direct container area
    neudorfWs.Range("B55:S72").Interior.colorIndex = xlNone
    
    ' Initialize counters
    carsProcessed = 0
    toursAdded = 0
    
    ' Initialize direct container counters
    For day = 0 To 6
        directContainersUsed(day) = 0
    Next day
    
    ' Create an array to store tours and their WAB counts for each day
    Dim wabToursByDay(0 To 6) As Object
    
    ' Initialize collections for each day
    For day = 0 To 6
        Set wabToursByDay(day) = CreateObject("Scripting.Dictionary")
    Next day
    
    ' First pass - collect all tours with WABs by day and time
    mainCarRow = 5
    Dim rowIndex As Integer
    rowIndex = 0 ' Use for preserving original row order
    
    Do While mainCarRow <= 22
        rowIndex = rowIndex + 1 ' Increment for each car row
        
        ' Process all days for this car (0-4 for Mon-Fri)
        For day = 0 To 4
            ' Calculate column positions
            sourceCol = 2 + (day * 4)  ' B, F, J, N, R, V, Z
            
            ' Get tour name
            tourName = Trim(mainWs.Cells(mainCarRow + 1, sourceCol).Value)
            
            ' Only process if there's a tour
            If tourName <> "" Then
                ' Get tour details
                tourTime = Trim(mainWs.Cells(mainCarRow + 2, sourceCol).Value)
                
                ' IMPORTANT: Get WAB count from the cell 1 to the right of tour time
                wabCount = GetWabCount(mainWs.Cells(mainCarRow + 2, sourceCol + 2))
                
                ' Format tour time if needed (ensure it's in hh:mm format)
                If tourTime <> "" Then
                    ' Check if time needs formatting
                    If Not InStr(tourTime, ":") > 0 Then
                        ' Try to format as hh:mm
                        On Error Resume Next
                        If IsNumeric(tourTime) Then
                            ' Convert numeric value to time
                            Dim timeValue As Double
                            timeValue = CDbl(tourTime)
                            tourTime = Format(timeValue, "hh:mm")
                        End If
                        On Error GoTo 0
                    End If
                End If
                
                ' Only add tours with WABs
                If wabCount > 0 Then
                    ' Store tour info as key=time_rowIndex, value=wabCount
                    ' Adding rowIndex ensures unique keys even with same time
                    Dim timeKey As String
                    timeKey = Format(tourTime, "hh:mm") & "_" & Format(rowIndex, "000")
                    
                    ' Add to dictionary for this day
                    wabToursByDay(day).Add timeKey, wabCount
                End If
            End If
        Next day
        
        ' Move to next car
        mainCarRow = mainCarRow + 3
    Loop
    
    ' Now generate WAB numbers for each day based on the tours
    Dim wabNumberDict As Object
    Set wabNumberDict = CreateObject("Scripting.Dictionary")
    
    ' Process each day
    For day = 0 To 6
        Dim wabCounter As Integer
        wabCounter = 1 ' Start at 1 for each day
        
        ' Skip empty days
        If wabToursByDay(day).count > 0 Then
            ' Get sorted list of keys for this day
            Dim timeKeys As Variant
            timeKeys = wabToursByDay(day).keys
            
            ' Sort the keys (important for time-based ordering)
            SortKeys timeKeys
            
            ' Debug output
            Debug.Print "Day " & (day + 1) & " has " & UBound(timeKeys) + 1 & " WAB tours"
            
            ' Process tours in time order
            Dim i As Integer
            For i = 0 To UBound(timeKeys)
                ' Get WAB count for this tour
                Dim tourWabCount As Integer
                tourWabCount = wabToursByDay(day).Item(timeKeys(i))
                
                ' Extract original time part from the key
                Dim timePart As String
                timePart = Left(timeKeys(i), 5) ' "hh:mm"
                
                ' Generate WAB number
                Dim wabNumber As String
                
                If tourWabCount = 1 Then
                    ' Single WAB format: 12_W_AA
                    wabNumber = "12_" & (day + 1) & "_" & Format(wabCounter, "00")
                    ' Debug output
                    Debug.Print "  Time: " & timePart & ", Single WAB: " & wabNumber
                Else
                    ' Multiple WAB format: 12_W_AA/BB
                    Dim endWab As Integer
                    endWab = wabCounter + tourWabCount - 1
                    wabNumber = "12_" & (day + 1) & "_" & Format(wabCounter, "00") & "/" & Format(endWab, "00")
                    ' Debug output
                    Debug.Print "  Time: " & timePart & ", Multiple WABs: " & wabNumber & " (count=" & tourWabCount & ")"
                End If
                
                ' Store WAB number with tour info (use original time and count for lookup)
                Dim lookupKey As String
                lookupKey = day & "|" & timePart & "|" & tourWabCount & "|" & timeKeys(i)
                wabNumberDict.Add lookupKey, wabNumber
                
                ' Increment counter for next tour
                wabCounter = wabCounter + tourWabCount
            Next i
        End If
    Next day
    
    ' Second pass - process the tours with assigned WAB numbers
    mainCarRow = 5
    subCarRow = 3
    rowIndex = 0
    
    Do While mainCarRow <= 22
        ' Process this car
        carsProcessed = carsProcessed + 1
        rowIndex = rowIndex + 1
        
        ' Process all days for this car (0-4 for Mon-Fri)
        For day = 0 To 4
            ' Calculate column positions
            sourceCol = 2 + (day * 4)  ' B, F, J, N, R, V, Z
            destCol = 2 + (day * 3)    ' B, E, H, K, N, Q, T
            
            ' Get tour name
            tourName = Trim(mainWs.Cells(mainCarRow + 1, sourceCol).Value)
            
            ' Only process if there's a tour
            If tourName <> "" Then
                ' Get tour details
                tourNumber = Trim(mainWs.Cells(mainCarRow, sourceCol).Value)
                workers = Trim(mainWs.Cells(mainCarRow, sourceCol + 3).Value)
                tourTime = Trim(mainWs.Cells(mainCarRow + 2, sourceCol).Value)
                
                ' Get WAB count for this tour - use the cell 2 columns right of the tour time
                wabCount = GetWabCount(mainWs.Cells(mainCarRow + 2, sourceCol + 2))
                
                ' Format tour time if needed
                If tourTime <> "" Then
                    If Not InStr(tourTime, ":") > 0 Then
                        On Error Resume Next
                        If IsNumeric(tourTime) Then
                            Dim timeValue2 As Double
                            timeValue2 = CDbl(tourTime)
                            tourTime = Format(timeValue2, "hh:mm")
                        End If
                        On Error GoTo 0
                    End If
                End If
                
                ' Get WAB number if applicable
                Dim wabNumbers As String
                wabNumbers = ""
                
                If wabCount > 0 Then
                    ' Create unique time key including row index
                    timeKey = Format(tourTime, "hh:mm") & "_" & Format(rowIndex, "000")
                    
                    ' Look up WAB number in our dictionary - try specific key first
                    lookupKey = day & "|" & Format(tourTime, "hh:mm") & "|" & wabCount & "|" & timeKey
                    
                    ' Try to find WAB number using various matching strategies
                    If wabNumberDict.Exists(lookupKey) Then
                        ' Exact match with time, count, and row
                        wabNumbers = wabNumberDict.Item(lookupKey)
                        Debug.Print "Found exact WAB match: " & wabNumbers & " for " & lookupKey
                    Else
                        ' Try with just day, time and count
                        Dim foundKey As String
                        foundKey = ""
                        
                        ' Try to find by day, time and count
                        Dim allKeys As Variant
                        allKeys = wabNumberDict.keys
                        
                        For i = 0 To UBound(allKeys)
                            Dim keyParts As Variant
                            keyParts = Split(allKeys(i), "|")
                            
                            ' Match day and time part (first 5 chars of keyParts(1))
                            If CInt(keyParts(0)) = day And _
                               Left(keyParts(1), 5) = Format(tourTime, "hh:mm") And _
                               CInt(keyParts(2)) = wabCount Then
                                foundKey = allKeys(i)
                                Exit For
                            End If
                        Next i
                        
                        If foundKey <> "" Then
                            wabNumbers = wabNumberDict.Item(foundKey)
                            Debug.Print "Found partial WAB match: " & wabNumbers & " for " & lookupKey
                        Else
                            ' Generate a WAB number as fallback
                            If wabCount = 1 Then
                                wabNumbers = "12_" & (day + 1) & "_" & Format(wabCounters(day), "00")
                            Else
                                endWab = wabCounters(day) + wabCount - 1
                                wabNumbers = "12_" & (day + 1) & "_" & Format(wabCounters(day), "00") & "/" & Format(endWab, "00")
                            End If
                            wabCounters(day) = wabCounters(day) + wabCount
                            Debug.Print "Generated fallback WAB: " & wabNumbers & " for " & lookupKey
                        End If
                    End If
                End If
                
                ' Check special flags
                Dim checkboxCol As Integer
                checkboxCol = sourceCol + 3
                hasNochPlatz = CheckboxValue(mainWs.Cells(mainCarRow + 1, checkboxCol))
                isDirectContainer = CheckboxValue(mainWs.Cells(mainCarRow + 2, sourceCol + 3))
                
                ' Modify name if "Noch Platz" is checked
                If hasNochPlatz Then
                    If InStr(tourName, "(Noch Platz)") = 0 Then
                        tourName = tourName & " (Noch Platz)"
                    End If
                End If
                
                ' Process LLnr numbers for color matching
                Dim llnrColor As Long
                llnrColor = ProcessLLnr(tourNumber, llnrDict, llnrColorIndex)
                
                ' Check if this should go to direct container area
                If isDirectContainer And directContainersUsed(day) < 6 Then
                    ' Add to direct container area
                    Dim directRow As Integer
                    directRow = 55 + (directContainersUsed(day) * 3)
                    
                    ' Copy tour data to direct container section
                    CopyTourToDestination neudorfWs, directRow, destCol, tourName, tourNumber, _
                                         tourTime, workers, wabCount, hasNochPlatz, isDirectContainer, llnrColor, wabNumbers
                    
                    ' Regular tour - copy to car's row
                    CopyTourToDestination neudorfWs, subCarRow, destCol, tourName, tourNumber, _
                                         tourTime, workers, wabCount, hasNochPlatz, isDirectContainer, llnrColor, wabNumbers
                    
                    ' Increment direct container counter for this day
                    directContainersUsed(day) = directContainersUsed(day) + 1
                Else
                    ' Regular tour - copy to car's row
                    CopyTourToDestination neudorfWs, subCarRow, destCol, tourName, tourNumber, _
                                         tourTime, workers, wabCount, hasNochPlatz, isDirectContainer, llnrColor, wabNumbers
                End If
                
                toursAdded = toursAdded + 1
            End If
        Next day
        
        ' Move to next car
        mainCarRow = mainCarRow + 3
        subCarRow = subCarRow + 3
    Loop
    
    ' Show results
    MsgBox "Transfer complete:" & vbCrLf & _
           "Cars processed: " & carsProcessed & vbCrLf & _
           "Tours added: " & toursAdded, vbInformation
End Sub

' Helper function to sort array of keys
Sub SortKeys(keys As Variant)
    Dim i As Integer, j As Integer
    Dim temp As Variant
    
    ' Simple bubble sort
    For i = 0 To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If keys(i) > keys(j) Then
                temp = keys(i)
                keys(i) = keys(j)
                keys(j) = temp
            End If
        Next j
    Next i
End Sub

Sub CopyTourToDestination(ws As Worksheet, row As Integer, col As Integer, tourName As String, _
                        tourNumber As String, tourTime As String, workers As String, _
                        wabCount As Integer, hasNochPlatz As Boolean, isDirectContainer As Boolean, _
                        Optional llnrColor As Long = -1, Optional wabNumbers As String = "")
    ' Copy the data
    ws.Cells(row, col).Value = tourName
    ws.Cells(row + 1, col).Value = tourNumber
    
    ' Format the workers text according to specification (2 BML + 2 LP)
    ws.Cells(row + 2, col).Value = FormatWorkers(workers)
    
    ' Place tour time in the correct cell
    ws.Cells(row + 2, col + 2).Value = tourTime
    ws.Cells(row + 2, col + 2).NumberFormat = "hh:mm"
    
    ' Apply direct container background if needed
    If isDirectContainer Then
        ' Apply light background color to the entire row section
        Dim bgRange As Range
        Set bgRange = ws.Range(ws.Cells(row, col), ws.Cells(row + 2, col + 2))
        bgRange.Interior.Color = RGB(240, 240, 255) ' Light blue background
    End If
    
    ' Format LLnr (tour number) with custom formatting
    With ws.Cells(row + 1, col)
        .NumberFormat = "@" ' Text format
        
        ' Apply color highlighting for matching LLnr numbers if provided
        If llnrColor <> -1 Then
            .Interior.Color = llnrColor
        End If
    End With
    
    ' Format worker numbers
    With ws.Cells(row + 2, col + 2)
        ' If workers is less than 2, make it red
        If IsNumeric(workers) Then
            If CInt(workers) < 2 Then
                .Font.Color = RGB(255, 0, 0) ' Red
                .Font.Bold = True
            End If
        End If
    End With
    
    ' Apply "Noch Platz" formatting
    Dim nameCell As Range
    Set nameCell = ws.Cells(row, col)  ' This is where tourName is placed
    
    ' Ensure cells are merged horizontally before applying text
    On Error Resume Next
    If Not nameCell.mergeCells Then
        ' Merge the tourName cell with 2 cells to the right
        ws.Range(ws.Cells(row, col), ws.Cells(row, col + 2)).Merge
    End If
    On Error GoTo 0
    
    ' Format the entire tour name cell with center alignment
    nameCell.HorizontalAlignment = xlCenter
    
    ' Check if the cell contains "(Noch Platz)"
    If InStr(nameCell.Value, "(Noch Platz)") > 0 Then
        Dim regularPart As String
        Dim nochPlatzStart As Integer
        
        nochPlatzStart = InStr(nameCell.Value, "(Noch Platz)")
        regularPart = Left(nameCell.Value, nochPlatzStart - 1)
        
        ' Set the cell to just the regular part first
        nameCell.Value = regularPart
        
        ' Then add the "(Noch Platz)" as red and bold text
        nameCell.Characters(Start:=Len(regularPart) + 1, Length:=12).Text = "(Noch Platz)"
        
        With nameCell.Characters(Start:=Len(regularPart) + 1, Length:=12).Font
            .Bold = True
            .Color = RGB(255, 0, 0) ' Red
        End With
    End If
    
    ' If we have WAB numbers for this tour, place them in the yellow box
    If wabNumbers <> "" And wabCount > 0 Then
        ' Choose correct placement based on your spreadsheet structure
        ' Based on your screenshot, it appears the WAB numbers should go in a yellow box
        Dim wabCell As Range
        
        ' Choose the correct position based on your images
        Set wabCell = ws.Cells(row + 1, col + 2)
        
        
        ' Place WAB number in the yellow box
        wabCell.Value = wabNumbers
        
        ' Format the WAB number cell
        With wabCell
            .Interior.Color = RGB(255, 255, 0) ' Yellow
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.Color = RGB(0, 0, 0) ' Black border
        End With
        
        ' Set tour name in first column
        ws.Cells(row, col).Value = tourName
        
        ' Remerge the cells on either side, skipping the WAB cell
        On Error Resume Next
        ' Don't merge them - just format them individually
        ws.Cells(row, col).HorizontalAlignment = xlLeft
        ws.Cells(row, col + 2).HorizontalAlignment = xlLeft
        On Error GoTo 0
    End If
End Sub

' Improved function to check checkbox value
Function CheckboxValue(cell As Range) As Boolean
    ' Default to False
    CheckboxValue = False
    
    ' Check for various checkbox representations
    If Not IsEmpty(cell.Value) Then
        ' Check for Excel TRUE/FALSE or WAHR/FALSCH
        If cell.Value = True Or _
           UCase(cell.Value) = "TRUE" Or _
           UCase(cell.Value) = "WAHR" Then
            CheckboxValue = True
        End If
    End If
End Function

' Function to format workers text (e.g., "2 BML + 2 LP")
Function FormatWorkers(workersText As String) As String
    Dim result As String
    result = workersText
    
    ' If it's just a number, format as "X BML"
    If IsNumeric(workersText) Then
        result = workersText & " BML"
    End If
    
    ' If it contains a "+", ensure proper spacing
    If InStr(workersText, "+") > 0 Then
        Dim parts As Variant
        parts = Split(workersText, "+")
        
        ' First part
        If IsNumeric(Trim(parts(0))) Then
            parts(0) = Trim(parts(0)) & " BML"
        End If
        
        ' Second part
        If UBound(parts) >= 1 Then
            If IsNumeric(Trim(parts(1))) Then
                parts(1) = Trim(parts(1)) & " LP"
            End If
            
            ' Combine with proper spacing
            result = Trim(parts(0)) & " + " & Trim(parts(1))
        End If
    End If
    
    FormatWorkers = result
End Function

' Improved function to get WAB count from a cell
Function GetWabCount(cell As Range) As Integer
    ' Default to 0
    GetWabCount = 0
    
    ' Empty cell = 0 WABs
    If IsEmpty(cell.Value) Then
        GetWabCount = 0
    ' Numeric value = that number of WABs
    ElseIf IsNumeric(cell.Value) Then
        GetWabCount = CInt(cell.Value)
    Else
        ' Try to extract a number if it's text
        Dim cellText As String
        cellText = Trim(CStr(cell.Value))
        
        ' If it's a number stored as text
        If IsNumeric(cellText) Then
            GetWabCount = CInt(cellText)
        Else
            ' Default to 1 for any other non-empty cell
            GetWabCount = 1
        End If
    End If
End Function

' Function to process LLnr and assign colors to matching numbers
Function ProcessLLnr(llnrText As String, llnrDict As Object, ByRef colorIndex As Integer) As Long
    ' Default color (none)
    ProcessLLnr = -1
    
    ' If empty, return default
    If Trim(llnrText) = "" Then Exit Function
    
    ' Split the LLnr by possible separators
    Dim llnrParts As Variant
    ' Replace all possible separators with space
    llnrText = Replace(llnrText, "+", " ")
    llnrText = Replace(llnrText, ",", " ")
    llnrText = Replace(llnrText, ";", " ")
    llnrText = Replace(llnrText, "-", " ")
    
    ' Clean up multiple spaces
    Do While InStr(llnrText, "  ") > 0
        llnrText = Replace(llnrText, "  ", " ")
    Loop
    
    llnrParts = Split(Trim(llnrText), " ")
    
    ' Check each 5-digit number
    Dim i As Integer, j As Integer
    Dim matchFound As Boolean
    matchFound = False
    
    ' First check if any part matches an existing key
    For i = 0 To UBound(llnrParts)
        If Len(Trim(llnrParts(i))) = 5 And IsNumeric(Trim(llnrParts(i))) Then
            If llnrDict.Exists(Trim(llnrParts(i))) Then
                ' Match found - use existing color
                ProcessLLnr = llnrDict(Trim(llnrParts(i)))
                matchFound = True
                Exit For
            End If
        End If
    Next i
    
    ' If no match found, assign a new color and add all parts to dictionary
    If Not matchFound Then
        ' Get a new color
        Dim newColor As Long
        newColor = GetNextColor(colorIndex)
        colorIndex = colorIndex + 1
        
        ' Add all parts to dictionary with this color
        For i = 0 To UBound(llnrParts)
            If Len(Trim(llnrParts(i))) = 5 And IsNumeric(Trim(llnrParts(i))) Then
                If Not llnrDict.Exists(Trim(llnrParts(i))) Then
                    llnrDict.Add Trim(llnrParts(i)), newColor
                End If
            End If
        Next i
        
        ProcessLLnr = newColor
    End If
End Function

' Function to get the next color in sequence
Function GetNextColor(index As Integer) As Long
    Dim colors(14) As Long
    
    ' Define a set of distinct colors
    colors(0) = RGB(255, 200, 200) ' Light red
    colors(1) = RGB(200, 255, 200) ' Light green
    colors(2) = RGB(200, 200, 255) ' Light blue
    colors(3) = RGB(255, 255, 200) ' Light yellow
    colors(4) = RGB(255, 200, 255) ' Light magenta
    colors(5) = RGB(200, 255, 255) ' Light cyan
    colors(6) = RGB(255, 230, 200) ' Light orange
    colors(7) = RGB(230, 200, 255) ' Light purple
    colors(8) = RGB(200, 255, 230) ' Light teal
    colors(9) = RGB(230, 255, 200) ' Light lime
    colors(10) = RGB(240, 240, 240) ' Light gray
    colors(11) = RGB(255, 220, 220) ' Very light red
    colors(12) = RGB(220, 255, 220) ' Very light green
    colors(13) = RGB(220, 220, 255) ' Very light blue
    
    ' Return the color at the current index, cycling through the array
    GetNextColor = colors(index Mod 14)
End Function

Sub TransferWABs()
    ' Placeholder for future WAB transfer functionality
    MsgBox "WAB-Transfer Funktion wird implementiert."
End Sub
