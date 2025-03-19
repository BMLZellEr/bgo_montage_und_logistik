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
    On Error GoTo 0
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
        ActiveSheet.Range("W2").Value = "Tourenplan für KW" & kwNumber & " (" & _
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
    ' Save the workbook to the specified folder with timestamp
    On Error GoTo ErrorHandler
    
    ' Setup the target folder path
    Dim targetFolder As String
    targetFolder = "C:\Users\bmlzeller\OneDrive - BGO Holding GmbH\BML_Dispo - Planung NOS - Planung NOS\10_Excel_Wocheneinteilung_Intern_NOS\KW_Autosave"
    
    ' Create the folder if it doesn't exist
    On Error Resume Next
    If Dir(targetFolder, vbDirectory) = "" Then
        MkDir targetFolder
    End If
    On Error GoTo ErrorHandler
    
    ' Generate timestamp
    Dim timeStamp As String
    timeStamp = Format(Now, "MM_dd_HH_mm")
    
    ' Generate filename
    Dim newFileName As String
    newFileName = targetFolder & "\AS_KW" & currentKW & "_NOS_Tourenkonzept_" & timeStamp & ".xlsm"
    
    ' Save a copy of the workbook without macros
    ThisWorkbook.SaveCopyAs newFileName
    
    SaveWorkbookToSpecialFolder = True
    Exit Function
    
ErrorHandler:
    MsgBox "Fehler beim Speichern: " & Err.Description, vbExclamation
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
        cell.Interior.ColorIndex = xlNone
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
        cell.Interior.ColorIndex = xlNone
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
    On Error GoTo 0
    
    ' Make sure we're back on the main sheet
    mainSheet.Activate
    
    ' Turn screen updating back on
    Application.ScreenUpdating = True
    
    MsgBox "Daten mit aktualisierten Datumsangaben übertragen!"
End Sub

Sub TransferTours()
    ' Placeholder for future Tour transfer functionality
    MsgBox "Tour-Transfer Funktion wird implementiert."
End Sub

Sub TransferWABs()
    ' Placeholder for future WAB transfer functionality
    MsgBox "WAB-Transfer Funktion wird implementiert."
End Sub
