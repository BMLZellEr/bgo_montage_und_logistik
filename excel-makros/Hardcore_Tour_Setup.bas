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
    
    ' Dictionary to track LLnr numbers (5-digit numbers) and their counts
    ' Key = 5-digit number, Value = count of appearances
    Dim llnrCountDict As Object
    Set llnrCountDict = CreateObject("Scripting.Dictionary")
    
    ' Dictionary to track LLnr numbers and their assigned colors
    ' Key = 5-digit number, Value = color
    Dim llnrColorDict As Object
    Set llnrColorDict = CreateObject("Scripting.Dictionary")
    
    Dim llnrColorIndex As Integer
    llnrColorIndex = 0 ' Start with first color
    
    ' Array to track WAB counters for each day
    Dim wabCounters(6) As Integer
    For day = 0 To 6
        wabCounters(day) = 1 ' Start WAB counter at 1 for each day
    Next day
    
    ' Define array of colors for direct tours (unique colors for each direct tour within a day)
    Dim directTourColors(14) As Long
    directTourColors(0) = RGB(220, 230, 255)  ' Light blue
    directTourColors(1) = RGB(255, 230, 220)  ' Light orange
    directTourColors(2) = RGB(220, 255, 220)  ' Light green
    directTourColors(3) = RGB(255, 220, 255)  ' Light magenta
    directTourColors(4) = RGB(255, 255, 220)  ' Light yellow
    directTourColors(5) = RGB(230, 230, 255)  ' Light purple
    directTourColors(6) = RGB(230, 255, 230)  ' Light mint
    directTourColors(7) = RGB(200, 220, 255)  ' Slightly darker blue
    directTourColors(8) = RGB(255, 210, 200)  ' Slightly darker orange
    directTourColors(9) = RGB(200, 255, 200)  ' Slightly darker green
    directTourColors(10) = RGB(255, 200, 245) ' Slightly darker magenta
    directTourColors(11) = RGB(255, 255, 200) ' Slightly darker yellow
    directTourColors(12) = RGB(210, 210, 255) ' Slightly darker purple
    directTourColors(13) = RGB(210, 255, 210) ' Slightly darker mint
    
    ' Dictionaries to track direct tour colors by day (key = tour number, value = color)
    Dim directTourColorsByDay(0 To 6) As Object
    For day = 0 To 6
        Set directTourColorsByDay(day) = CreateObject("Scripting.Dictionary")
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
    neudorfWs.Range("B3:S33").Font.Bold = False
    neudorfWs.Range("B55:S72").ClearContents ' Clear direct container area
    neudorfWs.Range("B55:S72").Interior.colorIndex = xlNone
    neudorfWs.Range("B55:S72").Font.Bold = False
    ' Also clear Lager container area
    neudorfWs.Range("B74:S81").ClearContents
    neudorfWs.Range("B74:S81").Interior.colorIndex = xlNone
    neudorfWs.Range("B74:S81").Font.Bold = False
    
    ' Initialize counters
    carsProcessed = 0
    toursAdded = 0
    
    ' Initialize direct container counters
    For day = 0 To 6
        directContainersUsed(day) = 0
    Next day
    
    ' First pass - scan all tour numbers to identify 5-digit numbers
    ' and count how many times each appears
    mainCarRow = 5
    
    Do While mainCarRow <= 22
        ' Process all days for this car (0-4 for Mon-Fri)
        For day = 0 To 4
            ' Calculate column positions
            sourceCol = 2 + (day * 4)  ' B, F, J, N, R, V, Z
            
            ' Get tour number
            tourNumber = Trim(mainWs.Cells(mainCarRow, sourceCol).Value)
            
            ' Process 5-digit numbers in tour number
            If tourNumber <> "" Then
                ' Clean the tour number - replace separators with spaces
                Dim cleanNumber As String
                cleanNumber = Replace(Replace(Replace(Replace(tourNumber, "+", " "), ",", " "), ";", " "), "-", " ")
                
                ' Remove multiple spaces
                Do While InStr(cleanNumber, "  ") > 0
                    cleanNumber = Replace(cleanNumber, "  ", " ")
                Loop
                
                ' Split by spaces
                Dim numParts As Variant
                numParts = Split(Trim(cleanNumber), " ")
                
                ' Check each part for 5-digit numbers
                Dim i As Integer
                For i = 0 To UBound(numParts)
                    If Len(Trim(numParts(i))) = 5 And IsNumeric(Trim(numParts(i))) Then
                        ' Found a 5-digit number
                        Dim fiveDigitNum As String
                        fiveDigitNum = Trim(numParts(i))
                        
                        ' Add to count dictionary or increment count
                        If llnrCountDict.Exists(fiveDigitNum) Then
                            llnrCountDict(fiveDigitNum) = llnrCountDict(fiveDigitNum) + 1
                        Else
                            llnrCountDict.Add fiveDigitNum, 1
                        End If
                    End If
                Next i
            End If
        Next day
        
        ' Move to next car
        mainCarRow = mainCarRow + 3
    Loop
    
    ' Assign colors only to LLnr numbers that appear more than once
    Dim llnrKey As Variant
    For Each llnrKey In llnrCountDict.keys
        If llnrCountDict(llnrKey) > 1 Then
            ' This 5-digit number appears multiple times - assign a color
            llnrColorDict.Add llnrKey, GetNextColor(llnrColorIndex)
            llnrColorIndex = llnrColorIndex + 1
        End If
    Next llnrKey
    
    ' Create a collection to store all tours with WABs in time order
    Dim wabTourCollection As New Collection
    Dim rowIndex As Integer
    
    ' Scan all tours to collect WAB info
    mainCarRow = 5
    rowIndex = 0
    
    Do While mainCarRow <= 22
        rowIndex = rowIndex + 1
        
        ' Process all days for this car
        For day = 0 To 4
            ' Calculate column position for the day
            sourceCol = 2 + (day * 4)  ' B, F, J, N, R, V, Z
            
            ' Check if there's a tour name
            tourName = Trim(mainWs.Cells(mainCarRow + 1, sourceCol).Value)
            
            If tourName <> "" Then
                ' Get tour details
                tourNumber = Trim(mainWs.Cells(mainCarRow, sourceCol).Value)
                tourTime = Trim(mainWs.Cells(mainCarRow + 2, sourceCol).Value)
                
                ' Get WAB count from the cell directly to the right of tour time
                Dim wabCellText As String
                wabCellText = Trim(CStr(mainWs.Cells(mainCarRow + 2, sourceCol + 1).Value))
                wabCount = 0 ' Default to 0
                
                ' Extract the number from text like "1 WAB's D" or "2 WAB's D"
                If wabCellText <> "" Then
                    If IsNumeric(Left(wabCellText, 1)) Then
                        wabCount = CInt(Left(wabCellText, 1))
                    End If
                End If
                
                ' Format time properly
                If tourTime <> "" Then
                    If Not InStr(tourTime, ":") > 0 Then
                        On Error Resume Next
                        If IsNumeric(tourTime) Then
                            Dim timeValue As Double
                            timeValue = CDbl(tourTime)
                            tourTime = Format(timeValue, "hh:mm")
                        End If
                        On Error GoTo 0
                    End If
                End If
                
                ' Check if this is a direct container
                isDirectContainer = CheckboxValue(mainWs.Cells(mainCarRow + 2, sourceCol + 3))
                
                ' Assign color for direct containers
                If isDirectContainer Then
                    Dim directTourKey As String
                    directTourKey = tourNumber & "|" & tourName
                    
                    If Not directTourColorsByDay(day).Exists(directTourKey) Then
                        Dim colorIndex As Integer
                        colorIndex = directTourColorsByDay(day).count Mod UBound(directTourColors)
                        directTourColorsByDay(day).Add directTourKey, directTourColors(colorIndex)
                    End If
                End If
                
                ' If this tour has WABs, add it to our collection
                If wabCount > 0 Then
                    ' Store as: day|time|wabCount|rowIndex|tourName|tourNumber|isDirectContainer
                    Dim tourData As String
                    tourData = day & "|" & Format(tourTime, "hh:mm") & "|" & wabCount & "|" & _
                               rowIndex & "|" & tourName & "|" & tourNumber & "|" & (isDirectContainer And 1)
                    
                    ' Add to collection
                    wabTourCollection.Add tourData
                End If
            End If
        Next day
        
        mainCarRow = mainCarRow + 3
    Loop
    
    ' Sort the tours by day and time
    Dim wabTourArray() As String
    ReDim wabTourArray(1 To wabTourCollection.count)
    
    For i = 1 To wabTourCollection.count
        wabTourArray(i) = wabTourCollection(i)
    Next i
    
    ' Simple bubble sort to order by day and time
    Dim j As Integer, temp As String
    For i = LBound(wabTourArray) To UBound(wabTourArray) - 1
        For j = i + 1 To UBound(wabTourArray)
            Dim parts1 As Variant, parts2 As Variant
            parts1 = Split(wabTourArray(i), "|")
            parts2 = Split(wabTourArray(j), "|")
            
            ' Compare day
            Dim day1 As Integer, day2 As Integer
            day1 = CInt(parts1(0))
            day2 = CInt(parts2(0))
            
            ' Compare time
            Dim time1 As String, time2 As String
            time1 = parts1(1)
            time2 = parts2(1)
            
            ' Sort by day, then by time
            If day1 > day2 Or (day1 = day2 And time1 > time2) Then
                temp = wabTourArray(i)
                wabTourArray(i) = wabTourArray(j)
                wabTourArray(j) = temp
            End If
        Next j
    Next i
    
    ' Create a dictionary to store WAB numbers
    Dim wabNumberDict As Object
    Set wabNumberDict = CreateObject("Scripting.Dictionary")
    
    ' Reset WAB counters to 1 for each day
    For day = 0 To 6
        wabCounters(day) = 1
    Next day
    
    ' Process the sorted tours and assign WAB numbers
    For i = LBound(wabTourArray) To UBound(wabTourArray)
        ' Parse the tour data
        Dim tourParts As Variant
        tourParts = Split(wabTourArray(i), "|")
        
        day = CInt(tourParts(0))
        Dim timeStr As String
        timeStr = tourParts(1)
        wabCount = CInt(tourParts(2))
        rowIndex = CInt(tourParts(3))
        tourName = tourParts(4)
        tourNumber = tourParts(5)
        isDirectContainer = (tourParts(6) = "1")
        
        ' Generate WAB number for this tour
        Dim wabNumber As String
        wabNumber = FormatWABNumber(day, wabCounters(day), wabCount)
        
        ' Create lookup key
        Dim lookupKey As String
        lookupKey = day & "|" & timeStr & "|" & wabCount & "|" & rowIndex
        
        ' Add to dictionary
        wabNumberDict.Add lookupKey, wabNumber
        
        ' Debug output
        Debug.Print "Day " & (day + 1) & ", Time: " & timeStr & ", Tour: " & tourName & _
                   ", WAB: " & wabNumber & " (count=" & wabCount & ")"
        
        ' Increment WAB counter for next tour
        wabCounters(day) = wabCounters(day) + wabCount
    Next i
    
    ' Create dictionary to store all WAB tours by day
    Dim wabToursByDay(0 To 6) As Object
    For day = 0 To 6
        Set wabToursByDay(day) = CreateObject("Scripting.Dictionary")
    Next day
    
    ' Process the tours one final time to put everything together
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
                
                ' Get WAB count from the cell directly to the right of tour time
                wabCellText = Trim(CStr(mainWs.Cells(mainCarRow + 2, sourceCol + 1).Value))
                wabCount = 0 ' Default to 0
                
                ' Extract the number from text like "1 WAB's D" or "2 WAB's D"
                If wabCellText <> "" Then
                    If IsNumeric(Left(wabCellText, 1)) Then
                        wabCount = CInt(Left(wabCellText, 1))
                    End If
                End If
                
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
                
                ' Check if this is a direct container
                isDirectContainer = CheckboxValue(mainWs.Cells(mainCarRow + 2, sourceCol + 3))
                
                ' Get the color for this direct tour if applicable
                Dim directTourColor As Long
                directTourColor = -1 ' Default (no color)
                
                If isDirectContainer Then
                    ' Create the same key used in the first pass
                    directTourKey = tourNumber & "|" & tourName
                    
                    ' Lookup the color for this direct tour
                    If directTourColorsByDay(day).Exists(directTourKey) Then
                        directTourColor = directTourColorsByDay(day).Item(directTourKey)
                    End If
                End If
                
                ' Get WAB number if applicable
                wabNumber = ""
                
                If wabCount > 0 Then
                    ' Create lookup key
                    lookupKey = day & "|" & Format(tourTime, "hh:mm") & "|" & wabCount & "|" & rowIndex
                    
                    ' Look up WAB number in our dictionary
                    If wabNumberDict.Exists(lookupKey) Then
                        wabNumber = wabNumberDict.Item(lookupKey)
                    Else
                        ' If not found for some reason, use a fallback
                        wabNumber = FormatWABNumber(day, wabCounters(day), wabCount)
                        wabCounters(day) = wabCounters(day) + wabCount
                    End If
                End If
                
                ' Check special flags
                Dim checkboxCol As Integer
                checkboxCol = sourceCol + 3
                hasNochPlatz = CheckboxValue(mainWs.Cells(mainCarRow + 1, checkboxCol))
                
                ' Modify name if "Noch Platz" is checked
                If hasNochPlatz Then
                    If InStr(tourName, "(Noch Platz)") = 0 Then
                        tourName = tourName & " (Noch Platz)"
                    End If
                End If
                
                ' Check for 5-digit numbers in this tour that need coloring
                Dim llnrColor As Long
                llnrColor = -1 ' Default (no color)
                Dim matchingLLnr As String
                matchingLLnr = ""
                
                ' Process the tour number to extract 5-digit numbers
                cleanNumber = Replace(Replace(Replace(Replace(tourNumber, "+", " "), ",", " "), ";", " "), "-", " ")
                
                ' Remove multiple spaces
                Do While InStr(cleanNumber, "  ") > 0
                    cleanNumber = Replace(cleanNumber, "  ", " ")
                Loop
                
                ' Split by spaces and look for matching 5-digit numbers
                numParts = Split(Trim(cleanNumber), " ")
                
                For i = 0 To UBound(numParts)
                    If Len(Trim(numParts(i))) = 5 And IsNumeric(Trim(numParts(i))) Then
                        ' This is a 5-digit number
                        fiveDigitNum = Trim(numParts(i))
                        
                        ' Check if this is a number that appears multiple times
                        If llnrColorDict.Exists(fiveDigitNum) Then
                            ' This is a number that needs coloring
                            llnrColor = llnrColorDict(fiveDigitNum)
                            matchingLLnr = fiveDigitNum
                            Exit For ' Use the first match found
                        End If
                    End If
                Next i
                
                ' Check if this should go to direct container area
                If isDirectContainer And directContainersUsed(day) < 6 Then
                    ' Add to direct container area
                    Dim directRow As Integer
                    directRow = 55 + (directContainersUsed(day) * 3)
                    
                    ' Copy tour data to direct container section with the unique direct tour color
                    CopyTourToDestination neudorfWs, directRow, destCol, tourName, tourNumber, _
                                         tourTime, workers, wabCount, hasNochPlatz, _
                                         isDirectContainer, llnrColor, matchingLLnr, wabNumber, directTourColor
                    
                    ' Regular tour - copy to car's row with the same unique direct tour color
                    CopyTourToDestination neudorfWs, subCarRow, destCol, tourName, tourNumber, _
                                         tourTime, workers, wabCount, hasNochPlatz, _
                                         isDirectContainer, llnrColor, matchingLLnr, wabNumber, directTourColor
                    
                    ' Increment direct container counter for this day
                    directContainersUsed(day) = directContainersUsed(day) + 1
                Else
                    ' Regular tour - copy to car's row
                    CopyTourToDestination neudorfWs, subCarRow, destCol, tourName, tourNumber, _
                                         tourTime, workers, wabCount, hasNochPlatz, _
                                         isDirectContainer, llnrColor, matchingLLnr, wabNumber, directTourColor
                End If
                
                toursAdded = toursAdded + 1
            End If
        Next day
        
        ' Move to next car
        mainCarRow = mainCarRow + 3
        subCarRow = subCarRow + 3
    Loop
    
    ' NEW PART: Process Lager-Container fields for non-direct tours with WABs
    
    ' Define colors for Lager containers
    Dim lgrColors(3) As Long
    lgrColors(0) = RGB(220, 245, 255) ' Soft Sky Blue
    lgrColors(1) = RGB(245, 220, 255) ' Gentle Lavender
    lgrColors(2) = RGB(225, 255, 220) ' Mint Mist
    lgrColors(3) = RGB(255, 245, 215) ' Warm Peach Glow
    
    ' Use unique variable names to avoid conflicts with existing variables
    Dim lgrNonDirectWabs(0 To 4, 0 To 3) As String  ' Array to store non-direct WAB numbers [day, index]
    Dim lgrWabCount(0 To 4) As Integer  ' Count of non-direct WABs per day
    
    ' Initialize counts
    For lgrDay = 0 To 4
        lgrWabCount(lgrDay) = 0
    Next lgrDay
    
    ' Scan through the destination sheet to find non-direct tours with WAB numbers
    For lgrDay = 0 To 4
        lgrDestCol = 2 + (lgrDay * 3)  ' B, E, H, K, N
        
        ' Scan through rows 3-33, stepping by 3 for each tour
        For lgrRow = 3 To 33 Step 3
            ' Check for WAB number cell (row+1, col+2)
            Dim lgrWabCell As Range
            Set lgrWabCell = neudorfWs.Cells(lgrRow + 1, lgrDestCol + 2)
            
            ' If cell has a value and yellow background (WAB number)
            If lgrWabCell.Value <> "" And lgrWabCell.Interior.Color = RGB(255, 255, 0) Then
                ' Check if this is NOT a direct container (no special background color)
                Dim lgrTourRange As Range
                Set lgrTourRange = neudorfWs.Range(neudorfWs.Cells(lgrRow, lgrDestCol), neudorfWs.Cells(lgrRow + 2, lgrDestCol))
                
                ' If tour is not a direct container (has no background color)
                If lgrTourRange.Interior.colorIndex = xlNone Or lgrTourRange.Interior.Color = RGB(255, 255, 255) Then
                With lgrWabCell
                    .Interior.Color = lgrColors(lgrIdx Mod 4)
                    .Font.Bold = True
                    .HorizontalAlignment = xlCenter
                    .Borders.LineStyle = xlContinuous
                    .Borders.Weight = xlThin
                    .Borders.Color = RGB(0, 0, 0) ' Black border
                End With
                    ' Store the WAB number if we still have space (max 4 per day)
                    If lgrWabCount(lgrDay) < 4 Then
                        lgrNonDirectWabs(lgrDay, lgrWabCount(lgrDay)) = lgrWabCell.Value
                        lgrWabCount(lgrDay) = lgrWabCount(lgrDay) + 1
                    End If
                End If
            End If
        Next lgrRow
    Next lgrDay
    
    ' Now place the non-direct WAB numbers in the Lager-Container fields
    For lgrDay = 0 To 4
        lgrDestCol = 2 + (lgrDay * 3)  ' B, E, H, K, N
        
        ' Process all non-direct WABs for this day
        For lgrIdx = 0 To lgrWabCount(lgrDay) - 1
            If lgrIdx < 4 Then ' Maximum 4 containers per day
                ' Calculate row for this Lager container (rows 74, 76, 78, 80)
                Dim lgrContainerRow As Integer
                lgrContainerRow = 74 + (lgrIdx * 2)
                
                ' Set the WAB number
                neudorfWs.Cells(lgrContainerRow, lgrDestCol).Value = lgrNonDirectWabs(lgrDay, lgrIdx)
                
                ' Format the entire 2x3 container field
                Dim lgrContainerRange As Range
                Set lgrContainerRange = neudorfWs.Cells(lgrContainerRow, lgrDestCol)
                
                ' Apply formatting
                With lgrContainerRange
                    .Interior.Color = lgrColors(lgrIdx Mod 4)
                    .Font.Bold = True
                    .HorizontalAlignment = xlCenter
                    .Borders.LineStyle = xlContinuous
                    .Borders.Weight = xlThin
                    .Borders.Color = RGB(0, 0, 0) ' Black border
                    
                End With
            End If
        Next lgrIdx
    Next lgrDay
    ' END OF NEW PART
    
    ' Show results
    MsgBox "Transfer complete:" & vbCrLf & _
           "Cars processed: " & carsProcessed & vbCrLf & _
           "Tours added: " & toursAdded, vbInformation
End Sub

' Function to properly format WAB numbers based on the count
Function FormatWABNumber(day As Integer, startWab As Integer, wabCount As Integer) As String
    ' For single WAB, format is 12_D_NN
    If wabCount = 1 Then
        FormatWABNumber = "12_" & (day + 1) & "_" & Format(startWab, "00")
    Else
        ' For exactly 2 WABs, use format 12_D_NN|MM (individual numbers separated by |)
        If wabCount = 2 Then
            ' For exactly 2 WABs, use pipe separator
            Dim secondWab As Integer
            secondWab = startWab + 1
            FormatWABNumber = "12_" & (day + 1) & "_" & Format(startWab, "00") & "|" & _
                             Format(secondWab, "00")
        Else
            ' For 3+ WABs, use format 12_D_NN/MM (range with /)
            Dim endWab As Integer
            endWab = startWab + wabCount - 1
            FormatWABNumber = "12_" & (day + 1) & "_" & Format(startWab, "00") & "/" & _
                             Format(endWab, "00")
        End If
    End If
End Function

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
                        Optional llnrColor As Long = -1, Optional llnrMatch As String = "", _
                        Optional wabNumbers As String = "", Optional directTourColor As Long = -1)
    ' Copy the data
    ws.Cells(row, col).Value = tourName
    ws.Cells(row + 1, col).Value = tourNumber
    
    ' Format the workers text according to specification (2 BML + 2 LP)
    ws.Cells(row + 2, col).Value = FormatWorkers(workers)
    
    ' Place tour time in the correct cell
    ws.Cells(row + 2, col + 2).Value = tourTime
    ws.Cells(row + 2, col + 2).NumberFormat = "hh:mm"
    
    ' Add borders to the entire 3x3 tour field
    Dim tourFieldRange As Range
    Set tourFieldRange = ws.Range(ws.Cells(row, col), ws.Cells(row + 2, col + 2))
    
    With tourFieldRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .colorIndex = xlAutomatic
    End With
    
    With tourFieldRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .colorIndex = xlAutomatic
    End With
    
    With tourFieldRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .colorIndex = xlAutomatic
    End With
    
    With tourFieldRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .colorIndex = xlAutomatic
    End With
    
    ' Apply direct container background if needed
    If isDirectContainer Then
        ' Apply the unique background color for this direct tour if provided
        If directTourColor <> -1 Then
            ' Use the unique color for this direct tour
            tourFieldRange.Interior.Color = directTourColor
        Else
            ' Fallback to default light blue background
            tourFieldRange.Interior.Color = RGB(240, 240, 255)
        End If
    End If
    
    ' Format LLnr (tour number) with custom formatting only if it's a match
    With ws.Cells(row + 1, col)
        .NumberFormat = "@" ' Text format
        
        ' Only apply special formatting for matching LLnr numbers
        If llnrColor <> -1 And llnrMatch <> "" Then
            .Interior.Color = llnrColor
            
            ' Add border to make match more obvious
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlMedium
            .Borders.Color = RGB(0, 0, 0) ' Black border
            
            ' Make the text bold to emphasize matched LLnr
            .Font.Bold = True
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
    
    ' Special handling for "(Noch Platz)" text
    If InStr(tourName, "(Noch Platz)") > 0 Then
        ' Clear any previous value
        nameCell.Value = ""
        
        ' Get the parts
        Dim regularPart As String
        Dim nochPlatzPart As String
        
        nochPlatzPart = "(Noch Platz)"
        regularPart = Replace(tourName, nochPlatzPart, "")
        
        ' Set the regular part first
        nameCell.Value = regularPart
        
        ' Add the Noch Platz part with specific formatting
        nameCell.Characters(Start:=Len(regularPart) + 1, Length:=12).Text = nochPlatzPart
        
        ' Format just this part
        With nameCell.Characters(Start:=Len(regularPart) + 1, Length:=12).Font
            .Bold = True
            .Color = RGB(255, 0, 0) ' Red
        End With
    Else
        ' Regular name without special formatting
        nameCell.Value = tourName
    End If
    
    ' If we have WAB numbers for this tour, place them in the yellow box
    If wabNumbers <> "" And wabCount > 0 Then
        ' Choose correct placement based on your spreadsheet structure
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
        
        ' Set tour name in first column if it was lost due to WAB cell
        If nameCell.Value = "" Then
            nameCell.Value = tourName
            
            ' Reapply Noch Platz formatting if needed
            If InStr(tourName, "(Noch Platz)") > 0 Then
                regularPart = Replace(tourName, "(Noch Platz)", "")
                nameCell.Value = regularPart
                nameCell.Characters(Start:=Len(regularPart) + 1, Length:=12).Text = "(Noch Platz)"
                With nameCell.Characters(Start:=Len(regularPart) + 1, Length:=12).Font
                    .Bold = True
                    .Color = RGB(255, 0, 0) ' Red
                End With
            End If
        End If
        
        ' Format the individual cells correctly
        On Error Resume Next
        ws.Cells(row, col).HorizontalAlignment = xlCenter
        ws.Cells(row, col + 2).HorizontalAlignment = xlCenter
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
