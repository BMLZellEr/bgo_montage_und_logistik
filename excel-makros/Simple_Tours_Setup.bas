Sub SetupSimple()
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
    
    ' Create button
    On Error Resume Next
    ActiveSheet.Buttons.Delete
    Dim btn As Button
    Set btn = ActiveSheet.Buttons.Add(ActiveSheet.Range("V1").Left, _
                                     ActiveSheet.Range("V1").Top + 30, 120, 20)
    With btn
        .Caption = "Daten laden"
        .OnAction = "SimpleTransferWithDates"
    End With
    On Error GoTo 0
    
    MsgBox "Dropdown erstellt. Wählen Sie ein Gebiet und klicken Sie 'Daten laden'."
End Sub

Sub SimpleTransferWithDates()
    ' This is the simplest possible approach with date synchronization
    
    ' Store active sheet reference
    Dim mainSheet As Worksheet
    Set mainSheet = ActiveSheet
    
    ' Get NOS_Tourenkonzept sheet for date reference
    Dim nosSheet As Worksheet
    On Error Resume Next
    Set nosSheet = ThisWorkbook.Worksheets("NOS_Tourenkonzept")
    If Err.Number <> 0 Or nosSheet Is Nothing Then
        ' If NOS_Tourenkonzept sheet not found, use active sheet
        Set nosSheet = mainSheet
    End If
    On Error GoTo 0
    
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
    
    ' Clear old data
    mainSheet.Range("W2:AP80").Clear
    
    ' Copy source data
    sourceWs.Range("A1:S80").Copy
    
    ' Paste to target
    mainSheet.Range("W2").PasteSpecial xlPasteAll
    
    ' Clear clipboard
    Application.CutCopyMode = False
    
    ' Set font size to 10
    mainSheet.Range("W2:AP80").Font.Size = 10
    
    ' Update date headers (X3, AA3, AD3, AG3, AJ3, AM3) from NOS_Tourenkonzept
    ' First check if B2 exists in NOS_Tourenkonzept and contains date
    Dim startDate As Date
    Dim hasValidDate As Boolean
    hasValidDate = False
    
    On Error Resume Next
    If IsDate(nosSheet.Range("B2").Value) Then
        startDate = nosSheet.Range("B2").Value
        hasValidDate = True
    ElseIf IsDate(nosSheet.Range("B1").Value) Then
        startDate = nosSheet.Range("B1").Value
        hasValidDate = True
    End If
    On Error GoTo 0
    
    ' If we have a valid date, update the date headers
    If hasValidDate Then
        ' Update the 6 date headers (typically Monday through Saturday)
        For i = 0 To 5
            ' Calculate the date for this day (startDate + i days)
            Dim dayDate As Date
            dayDate = startDate + i
            
            ' Format the date as "Weekday, DD.MM.YYYY"
            Dim dateStr As String
            dateStr = Format(dayDate, "dddd, dd.mm.yyyy")
            
            ' Calculate the column for this day's header
            ' Starting at X(24) with steps of 3 columns: X(24), AA(27), AD(30), etc.
            Dim colIndex As Integer
            colIndex = 24 + (i * 3)
            
            ' Update the header if it exists
            If Not IsEmpty(mainSheet.Cells(3, colIndex).Value) Then
                mainSheet.Cells(3, colIndex).Value = dateStr
                
                ' Ensure it's merged across 3 columns if needed
                On Error Resume Next
                If Not mainSheet.Cells(3, colIndex).MergeCells Then
                    mainSheet.Range(mainSheet.Cells(3, colIndex), mainSheet.Cells(3, colIndex + 2)).Merge
                End If
                On Error GoTo 0
            End If
        Next i
        
        ' Update the title with current KW
        Dim kwInfo As String
        kwInfo = "KW" & WorksheetFunction.weekNum(startDate) & " (" & _
                Format(startDate, "dd.mm.yyyy") & " - " & _
                Format(startDate + 5, "dd.mm.yyyy") & ")"
        
        ' If title cell exists, update it with KW info
        If Not IsEmpty(mainSheet.Range("W2").Value) Then
            Dim origTitle As String
            origTitle = mainSheet.Range("W2").Value
            
            ' Remove any existing KW info if present
            If InStr(origTitle, "KW") > 0 Then
                origTitle = Left(origTitle, InStr(origTitle, "KW") - 2)
            End If
            
            ' Add the new KW info
            mainSheet.Range("W2").Value = origTitle & " - " & kwInfo
        End If
    End If
    
    ' Make sure we're back on the main sheet
    mainSheet.Activate
    
    ' Turn screen updating back on
    Application.ScreenUpdating = True
    
    MsgBox "Daten mit aktualisierten Datumsangaben übertragen!"
End Sub
