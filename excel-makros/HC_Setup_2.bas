Sub SetupMainControls()
    ' Create the main controls (dropdown and all buttons) in a single routine
    Application.ScreenUpdating = False
    
    ' Set up the Gebiet dropdown
    SetupGebietDropdown
    
    ' Set up the KW switcher
    SetupKWSwitcher
    
    ' Add all action buttons in a single call 
    AddActionButtons
    
    Application.ScreenUpdating = True
    MsgBox "Steuerelemente wurden erstellt. Klicken Sie auf 'KW wechseln' zum Ändern der Kalenderwoche."
End Sub

Sub SetupGebietDropdown()
    ' Create dropdown with sheet names matching pattern "Tourenplan_BML_*"
    Dim sheetList As String
    
    ' Build list of sheets using filter
    sheetList = FilterSheetNames("Tourenplan_BML_*")
    
    ' Create dropdown
    With ActiveSheet.Range("W1").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:=sheetList
        .InCellDropdown = True
    End With
    
    ' Add label
    ActiveSheet.Range("V1").Value = "Gebiet:"
End Sub

Function FilterSheetNames(pattern As String) As String
    ' Returns comma-separated list of sheet names matching the pattern
    Dim result As String, i As Integer
    
    result = ""
    For i = 1 To ThisWorkbook.Sheets.count
        If ThisWorkbook.Sheets(i).Name Like pattern Then
            result = result & ThisWorkbook.Sheets(i).Name & ","
        End If
    Next i
    
    ' Remove trailing comma if present
    If Len(result) > 0 Then
        FilterSheetNames = Left(result, Len(result) - 1)
    Else
        FilterSheetNames = ""
    End If
End Function

Sub SetupKWSwitcher()
    ' Create a KW switcher button
    On Error Resume Next
    ActiveSheet.Buttons.Delete  ' Remove existing button if any
    
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
    ' Add all action buttons in a more concise way
    On Error Resume Next
    ' Define button configurations
    Dim buttonConfig As Variant
    buttonConfig = Array( _
        Array("btnLoadData", "Daten laden", "SimpleTransferWithDates"), _
        Array("btnTransferTours", "Transfer Tours", "TransferTours"), _
        Array("btnTransferWABs", "Transfer WABs", "TransferWABs"), _
        Array("btnPrint", "Drucken", "PrintWorksheet") _
    )
    
    ' Base position for first button
    Dim baseLeft As Double, baseTop As Double
    baseLeft = ActiveSheet.Range("V1").Left
    baseTop = ActiveSheet.Range("V1").Top + 30
    
    ' Add each button with consistent spacing
    Dim i As Integer, btn As Button
    For i = 0 To UBound(buttonConfig)
        Set btn = ActiveSheet.Buttons.Add(baseLeft + (i * 130), baseTop, 120, 20)
        With btn
            .Name = buttonConfig(i)(0)
            .Caption = buttonConfig(i)(1)
            .OnAction = buttonConfig(i)(2)
        End With
    Next i
    On Error GoTo 0
End Sub

Sub PrintWorksheet()
    ' Print the NOS_Tourenkonzept_Print worksheet
    Dim currentSheet As Worksheet, printSheet As Worksheet
    
    ' Save current sheet
    Set currentSheet = ActiveSheet
    
    ' Check if print sheet exists
    On Error Resume Next
    Set printSheet = ThisWorkbook.Worksheets("NOS_Tourenkonzept_Print")
    
    If printSheet Is Nothing Then
        MsgBox "Das Druckblatt 'NOS_Tourenkonzept_Print' wurde nicht gefunden.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Print and return to original sheet
    printSheet.Activate
    Application.Dialogs(xlDialogPrint).Show
    currentSheet.Activate
End Sub

Sub SwitchKW()
    ' KW switcher with improved error handling and user experience
    Dim kwInput As String, kwNumber As Integer, firstMonday As Date, currentKW As Integer
    
    ' Get current KW from B1 date
    On Error Resume Next
    If IsDate(ActiveSheet.Range("B1").Value) Then
        currentKW = GetKW2025FromDate(CDate(ActiveSheet.Range("B1").Value))
    Else
        currentKW = GetKW2025FromDate(Date)
    End If
    On Error GoTo 0
    
    ' Ask for desired KW
    kwInput = InputBox("Bitte geben Sie die gewünschte Kalenderwoche ein (1-53):" & vbCrLf & _
                      "Aktuelle KW: " & currentKW, "KW wechseln", currentKW)
    
    ' Handle input validation
    If kwInput = "" Then Exit Sub ' User canceled
    
    If Not IsNumeric(kwInput) Then
        MsgBox "Bitte geben Sie eine gültige Zahl ein.", vbExclamation
        Exit Sub
    End If
    
    kwNumber = CInt(kwInput)
    
    If kwNumber < 1 Or kwNumber > 53 Then
        MsgBox "Bitte geben Sie eine Zahl zwischen 1 und 53 ein.", vbExclamation
        Exit Sub
    End If
    
    ' Check if already in selected KW
    If kwNumber = currentKW Then
        MsgBox "Sie befinden sich bereits in KW" & kwNumber & ".", vbInformation
        Exit Sub
    End If
    
    ' Get Monday date for selected KW
    firstMonday = GetDateForKW2025(kwNumber)
    
    ' Save current file before switching
    If SaveWorkbookToSpecialFolder(currentKW) Then
        ' Clear data and update dates
        Application.ScreenUpdating = False
        
        ClearAllData
        
        ' Update date and KW label
        ActiveSheet.Range("B1").Value = firstMonday
        ActiveSheet.Range("X1").Value = "Tourenplan für KW" & kwNumber & " (" & _
                                      Format(firstMonday, "dd.mm.yyyy") & " - " & _
                                      Format(firstMonday + 4, "dd.mm.yyyy") & ")"
        
        Application.ScreenUpdating = True
        MsgBox "Erfolgreich zu KW" & kwNumber & " gewechselt. Das Datum in B1 wurde auf " & _
               Format(firstMonday, "dd.mm.yyyy") & " gesetzt.", vbInformation
    Else
        MsgBox "KW-Wechsel abgebrochen, da die aktuelle Datei nicht gespeichert werden konnte.", vbExclamation
    End If
End Sub

Function GetDateForKW2025(kwNumber As Integer) As Date
    ' Returns Monday date for a given KW in 2025
    Dim kw1Monday As Date
    kw1Monday = DateSerial(2024, 12, 30) ' KW1 Monday for 2025
    
    GetDateForKW2025 = kw1Monday + ((kwNumber - 1) * 7)
End Function

Function GetKW2025FromDate(targetDate As Date) As Integer
    ' Calculate KW from date for 2025
    Dim kw1Monday As Date, daysSinceKW1 As Integer, kwNumber As Integer
    kw1Monday = DateSerial(2024, 12, 30) ' KW1 Monday for 2025
    
    ' Handle dates before KW1
    If targetDate < kw1Monday Then
        GetKW2025FromDate = 1
        Exit Function
    End If
    
    ' Calculate KW number
    daysSinceKW1 = targetDate - kw1Monday
    kwNumber = (daysSinceKW1 \ 7) + 1
    
    ' Cap at KW53
    If kwNumber > 53 Then kwNumber = 53
    
    GetKW2025FromDate = kwNumber
End Function

Function SaveWorkbookToSpecialFolder(currentKW As Integer) As Boolean
    ' Save workbook to OneDrive path with optimized error handling
    On Error GoTo ErrorHandler
    
    ' Get user profile path and username
    Dim userProfile As String, userName As String, targetFolder As String
    userProfile = Environ("USERPROFILE")
    userName = Environ("USERNAME")
    
    ' Set target folder path
    targetFolder = userProfile & "\OneDrive - BGO Holding GmbH\BML_Dispo - Planung NOS - Planung NOS\10_Excel_Wocheneinteilung_Intern_NOS\Autosave"
    
    ' Create folder if needed
    On Error Resume Next
    If Dir(targetFolder, vbDirectory) = "" Then MkDir targetFolder
    On Error GoTo ErrorHandler
    
    ' Generate filename with timestamp
    Dim newFileName As String
    newFileName = targetFolder & "\NOS_TK_" & userName & "_" & Format(Now, "dd_MM_HH_mm") & ".xlsx"
    
    ' Save a copy
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
    ' Clear all data fields with improved structure
    Application.ScreenUpdating = False
    
    ' Clear main sheet
    Dim nosSheet As Worksheet
    On Error Resume Next
    Set nosSheet = ThisWorkbook.Worksheets("NOS_Tourenkonzept")
    If Not nosSheet Is Nothing Then
        ' Clear data for each day of the week
        ' Monday to Friday with specific pattern
        Dim days As Variant, day As Variant
        days = Array( _
            Array("B2:C100", "E2:E100", "E"), _
            Array("F2:G100", "I2:I100", "I"), _
            Array("J2:K100", "M2:M100", "M"), _
            Array("N2:O100", "Q2:Q100", "Q"), _
            Array("R2:S100", "U2:U100", "U") _
        )
        
        ' Common row numbers for all days
        Dim specialRows As Variant
        specialRows = Array(2, 5, 8, 11, 14, 17, 20, 23, 26, 29, 32, 35, 38, 41, 44)
        
        ' Process each day
        For Each day In days
            ClearRangeAndCheckboxes nosSheet, day(0), day(1)
            ClearSpecificCells nosSheet, day(2), specialRows
        Next day
    End If
    On Error GoTo 0
    
    ' Clear all subsheets with pattern matching
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "Tourenplan_BML_*" Then
            ' Clear while preserving headers
            ClearTourFields ws, "B3:S33"
            ClearWABFields ws, "B35:S38"
        End If
    Next ws
    
    Application.ScreenUpdating = True
End Sub

Sub ClearRangeAndCheckboxes(ws As Worksheet, dataRange As String, checkboxRange As String)
    ' Clear data and reset checkboxes
    ws.Range(dataRange).ClearContents
    
    Dim cell As Range
    For Each cell In ws.Range(checkboxRange)
        If IsCheckbox(cell) Then cell.Value = False
    Next cell
End Sub

Sub ClearSpecificCells(ws As Worksheet, columnLetter As String, rowNumbers As Variant)
    ' Clear specific cells in a column
    Dim i As Integer, cellAddress As String
    
    For i = LBound(rowNumbers) To UBound(rowNumbers)
        cellAddress = columnLetter & rowNumbers(i)
        
        ws.Range(cellAddress).ClearContents
        
        If IsCheckbox(ws.Range(cellAddress)) Then
            ws.Range(cellAddress).Value = False
        End If
    Next i
End Sub

Function IsCheckbox(cell As Range) As Boolean
    ' Check if cell contains boolean value
    On Error Resume Next
    IsCheckbox = (TypeName(cell.Value) = "Boolean")
    On Error GoTo 0
End Function

Sub ClearTourFields(ws As Worksheet, tourRange As String)
    ' Clear tour fields but preserve formatting
    Dim rng As Range, cell As Range
    Set rng = ws.Range(tourRange)
    
    rng.ClearContents
    
    For Each cell In rng
        cell.Interior.colorIndex = xlNone
    Next cell
End Sub

Sub ClearWABFields(ws As Worksheet, wabRange As String)
    ' Clear WAB fields but preserve formatting
    Dim rng As Range, cell As Range
    Set rng = ws.Range(wabRange)
    
    rng.ClearContents
    
    For Each cell In rng
        cell.Interior.colorIndex = xlNone
    Next cell
End Sub

Sub SimpleTransferWithDates()
    ' Transfer data with date synchronization
    
    ' Store reference to active sheet
    Dim mainSheet As Worksheet, sourceWs As Worksheet
    Set mainSheet = ActiveSheet
    
    ' Get selected sheet from dropdown
    Dim selectedSheet As String
    selectedSheet = mainSheet.Range("W1").Value
    
    If selectedSheet = "" Then
        MsgBox "Bitte wählen Sie ein Gebiet aus."
        Exit Sub
    End If
    
    ' Get source sheet
    On Error Resume Next
    Set sourceWs = ThisWorkbook.Worksheets(selectedSheet)
    If Err.Number <> 0 Or sourceWs Is Nothing Then
        MsgBox "Blatt nicht gefunden: " & selectedSheet
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Prepare for transfer
    Application.ScreenUpdating = False
    
    ' Clear old data
    On Error Resume Next
    mainSheet.Range("W2:AP80").UnMerge
    On Error GoTo 0
    mainSheet.Range("W2:AP80").Clear
    
    ' Copy and paste data
    sourceWs.Range("A1:S80").Copy
    mainSheet.Range("W2").PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    
    ' Format transferred data
    mainSheet.Range("W2:AP80").Font.Size = 10
    mainSheet.Range("W2:W3").Merge
    
    ' Get date and KW info
    Dim startDate As Date, kwNumber As Integer
    
    If IsDate(mainSheet.Range("B1").Value) Then
        startDate = mainSheet.Range("B1").Value
    Else
        startDate = Date ' Fallback to current date
    End If
    
    kwNumber = GetKW2025FromDate(startDate)
    
    ' Update title with KW info
    With mainSheet.Range("X1")
        .Value = "Tourenplan für KW" & kwNumber & " (" & _
                Format(startDate, "dd.mm.yyyy") & " - " & _
                Format(startDate + 4, "dd.mm.yyyy") & ")"
        .Font.Bold = True
        .Font.Size = 11
    End With
    
    ' Update date headers for each day
    For i = 0 To 5 ' Monday to Saturday
        Dim dayDate As Date, dateStr As String, colIndex As Integer
        
        dayDate = startDate + i
        dateStr = Format(dayDate, "dddd, dd.mm.yyyy")
        colIndex = 24 + (i * 3) ' Column position (X, AA, AD, etc.)
        
        ' Update header
        mainSheet.Cells(3, colIndex).Value = dateStr
        
        ' Merge cells if needed
        If Not mainSheet.Cells(3, colIndex).MergeCells Then
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
    
    ' Return to main sheet
    mainSheet.Activate
    
    Application.ScreenUpdating = True
    MsgBox "Daten mit aktualisierten Datumsangaben übertragen!"
End Sub

' Improved TransferTours and TransferWABs functions

Sub TransferTours()
    ' Transfer tours with improved structure and performance
    
    ' Constants for colors and other settings
    Const MAX_DIRECT_CONTAINERS As Integer = 6
    Const MAX_LAGER_CONTAINERS As Integer = 4
    
    ' Create dictionaries and trackers
    Dim llnrCountDict As Object, llnrColorDict As Object
    Dim directTourColorsByDay(0 To 6) As Object
    Dim wabCounters(6) As Integer, directContainersUsed(6) As Integer
    
    ' Initialize objects
    Set llnrCountDict = CreateObject("Scripting.Dictionary")
    Set llnrColorDict = CreateObject("Scripting.Dictionary")
    
    ' Initialize day-specific dictionaries
    Dim day As Integer
    For day = 0 To 6
        Set directTourColorsByDay(day) = CreateObject("Scripting.Dictionary")
        wabCounters(day) = 1
        directContainersUsed(day) = 0
    Next day
    
    ' Color arrays for different tour types
    Dim directTourColors(13) As Long, lgrColors(3) As Long
    directTourColors = GetDirectTourColorArray()
    lgrColors = GetLagerContainerColorArray()
    
    ' Get worksheets
    Dim mainWs As Worksheet, targetWs As Worksheet
    Set mainWs = ThisWorkbook.Worksheets("NOS_Tourenkonzept")
    
    ' Get selected sheet or default
    Dim selectedSheet As String
    selectedSheet = mainWs.Range("W1").Value
    
    If selectedSheet = "" Then
        Set targetWs = ThisWorkbook.Worksheets("Tourenplan_BML_WNeudorf")
    Else
        On Error Resume Next
        Set targetWs = ThisWorkbook.Worksheets(selectedSheet)
        If targetWs Is Nothing Then
            Set targetWs = ThisWorkbook.Worksheets("Tourenplan_BML_WNeudorf")
        End If
        On Error GoTo 0
    End If
    
    ' Clear tour areas in destination sheet
    Application.ScreenUpdating = False
    ClearDestinationAreas targetWs
    
    ' PHASE 1: Scan all tour numbers to identify and count 5-digit numbers
    ScanAndCountLLnr mainWs, llnrCountDict
    
    ' Assign colors to LLnr numbers that appear multiple times
    AssignColorsToLLnr llnrCountDict, llnrColorDict
    
    ' PHASE 2: Collect all tours with WABs in time order
    Dim wabTourCollection As New Collection
    CollectWABTours mainWs, wabTourCollection
    
    ' Sort tours by day and time
    Dim wabTourArray() As String
    wabTourArray = SortWABTours(wabTourCollection)
    
    ' PHASE 3: Assign WAB numbers to all tours with WABs
    Dim wabNumberDict As Object
    Set wabNumberDict = CreateObject("Scripting.Dictionary")
    
    AssignWABNumbers wabTourArray, wabNumberDict, wabCounters
    
    ' PHASE 4: Process all tours and copy to destination
    TransferAllTours mainWs, targetWs, llnrColorDict, directTourColorsByDay, wabNumberDict, _
                    directTourColors, directContainersUsed
    
    ' PHASE 5: Process Lager-Container fields for non-direct WABs
    ProcessLagerContainers targetWs, lgrColors
    
    Application.ScreenUpdating = True
    
    ' Show results summary
    MsgBox "Transfer complete!", vbInformation
End Sub

' Helper functions would follow here...

Function GetDirectTourColorArray() As Variant
    ' Return array of colors for direct tours
    Dim colors(13) As Long
    
    colors(0) = RGB(220, 230, 255)   ' Light blue
    colors(1) = RGB(255, 230, 220)   ' Light orange
    colors(2) = RGB(220, 255, 220)   ' Light green
    colors(3) = RGB(255, 220, 255)   ' Light magenta
    colors(4) = RGB(255, 255, 220)   ' Light yellow
    colors(5) = RGB(230, 230, 255)   ' Light purple
    colors(6) = RGB(230, 255, 230)   ' Light mint
    colors(7) = RGB(200, 220, 255)   ' Slightly darker blue
    colors(8) = RGB(255, 210, 200)   ' Slightly darker orange
    colors(9) = RGB(200, 255, 200)   ' Slightly darker green
    colors(10) = RGB(255, 200, 245)  ' Slightly darker magenta
    colors(11) = RGB(255, 255, 200)  ' Slightly darker yellow
    colors(12) = RGB(210, 210, 255)  ' Slightly darker purple
    colors(13) = RGB(210, 255, 210)  ' Slightly darker mint
    
    GetDirectTourColorArray = colors
End Function

Function GetLagerContainerColorArray() As Variant
    ' Return array of colors for lager containers
    Dim colors(3) As Long
    
    colors(0) = RGB(220, 245, 255)  ' Soft Sky Blue
    colors(1) = RGB(245, 220, 255)  ' Gentle Lavender
    colors(2) = RGB(225, 255, 220)  ' Mint Mist
    colors(3) = RGB(255, 245, 215)  ' Warm Peach Glow
    
    GetLagerContainerColorArray = colors
End Function

Sub ClearDestinationAreas(ws As Worksheet)
    ' Clear all destination areas
    ws.Range("B3:S33").ClearContents
    ws.Range("B3:S33").Interior.colorIndex = xlNone
    ws.Range("B3:S33").Font.Bold = False
    
    ws.Range("B55:S72").ClearContents  ' Direct container area
    ws.Range("B55:S72").Interior.colorIndex = xlNone
    ws.Range("B55:S72").Font.Bold = False
    
    ws.Range("B74:S81").ClearContents  ' Lager container area
    ws.Range("B74:S81").Interior.colorIndex = xlNone
    ws.Range("B74:S81").Font.Bold = False
End Sub

Function FormatWABNumber(day As Integer, startWab As Integer, wabCount As Integer) As String
    ' Format WAB numbers based on count rules
    If wabCount = 1 Then
        FormatWABNumber = "12_" & (day + 1) & "_" & Format(startWab, "00")
    ElseIf wabCount = 2 Then
        FormatWABNumber = "12_" & (day + 1) & "_" & Format(startWab, "00") & "|" & _
                         Format(startWab + 1, "00")
    Else
        FormatWABNumber = "12_" & (day + 1) & "_" & Format(startWab, "00") & "/" & _
                         Format(startWab + wabCount - 1, "00")
    End If
End Function

Sub TransferWABs()
    ' Placeholder for WAB transfer function
    MsgBox "WAB-Transfer Funktion wird implementiert."
End Sub
