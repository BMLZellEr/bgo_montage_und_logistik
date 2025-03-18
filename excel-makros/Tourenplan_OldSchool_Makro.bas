' ------------------------------------------------------------------------
' Excel VBA Solution with 5x5 Grid of Fields
' 
' This solution provides:
' 1. A 5x5 grid of fields (25 total) that can be edited simultaneously
' 2. Simplified controls at the top (Load All, Save All, Load Single Field)
' 3. Support for both Tour Fields (3x3) and WAB Fields (2x3)
' 4. Easy selection of week and location
' ------------------------------------------------------------------------

Option Explicit

' Constants for field locations
Const MAIN_SHEET As String = "NOS_Tourenkonzept"
Const WEEK_CELL As String = "AI2" ' Pink field for week selection
Const LOCATION_CELL As String = "AJ2" ' Pink field for location selection
Const SINGLE_FIELD_SELECTOR As String = "AL2" ' Field for selecting a single field to load

' Grid dimensions
Const GRID_ROWS As Integer = 5
Const GRID_COLS As Integer = 5
Const TOTAL_FIELDS As Integer = GRID_ROWS * GRID_COLS

' Field starting positions and spacing on the main sheet
Const FIELD_START_ROW As Integer = 20
Const FIELD_START_COL As Integer = 10 ' Column J
Const FIELD_ROW_SPACING As Integer = 5 ' Space between fields vertically
Const FIELD_COL_SPACING As Integer = 5 ' Space between fields horizontally
Const FIELD_WIDTH As Integer = 3 ' Width of each field (3 columns)
Const FIELD_HEIGHT As Integer = 3 ' Height of fields (3 rows for Tour, 2 for WAB)

' Setup the entire grid system
Sub SetupGridFieldSystem()
    Dim ws As Worksheet
    
    ' Get main worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(MAIN_SHEET)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Worksheet '" & MAIN_SHEET & "' not found.", vbExclamation
        Exit Sub
    End If
    
    ' Clear the area where we'll place our grid system
    ClearGridArea ws
    
    ' Setup the top controls
    SetupTopControls ws
    
    ' Setup the grid of fields
    SetupFieldGrid ws
    
    MsgBox "Grid field system created successfully!" & vbCrLf & _
           "You can now work with up to 25 fields simultaneously.", vbInformation
End Sub

' Clear the area where the grid will be placed
Private Sub ClearGridArea(ws As Worksheet)
    ' Calculate the area to clear based on our grid dimensions
    Dim lastRow As Integer, lastCol As Integer
    
    lastRow = FIELD_START_ROW + (GRID_ROWS * FIELD_ROW_SPACING) + 10
    lastCol = FIELD_START_COL + (GRID_COLS * FIELD_COL_SPACING) + 10
    
    ' Clear the content and formatting
    ws.Range(ws.Cells(FIELD_START_ROW - 10, FIELD_START_COL - 2), ws.Cells(lastRow, lastCol)).Clear
    
    ' Remove any shapes (buttons, etc.)
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.TopLeftCell.Row >= FIELD_START_ROW - 10 And shp.TopLeftCell.Row <= lastRow And _
           shp.TopLeftCell.Column >= FIELD_START_COL - 2 And shp.TopLeftCell.Column <= lastCol Then
            shp.Delete
        End If
    Next shp
End Sub

' Setup the controls at the top of the grid
Private Sub SetupTopControls(ws As Worksheet)
    ' Create a title
    ws.Cells(FIELD_START_ROW - 9, FIELD_START_COL).Value = "TOUR PLANNING GRID SYSTEM"
    ws.Cells(FIELD_START_ROW - 9, FIELD_START_COL).Font.Bold = True
    ws.Cells(FIELD_START_ROW - 9, FIELD_START_COL).Font.Size = 14
    
    ' Create week selection dropdown
    ws.Cells(FIELD_START_ROW - 7, FIELD_START_COL).Value = "Week:"
    ws.Cells(FIELD_START_ROW - 7, FIELD_START_COL).Font.Bold = True
    
    With ws.Range(WEEK_CELL)
        .Value = "KW Selection"
        .Font.Bold = True
        .Interior.Color = RGB(255, 192, 203) ' Pink
        
        ' Create validation list with KW values
        With .Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:="KW 1/2025,KW 2/2025,KW 3/2025,KW 4/2025,KW 5/2025"
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
    End With
    
    ' Create location selection dropdown
    ws.Cells(FIELD_START_ROW - 7, FIELD_START_COL + 5).Value = "Location:"
    ws.Cells(FIELD_START_ROW - 7, FIELD_START_COL + 5).Font.Bold = True
    
    With ws.Range(LOCATION_CELL)
        .Value = "Location"
        .Font.Bold = True
        .Interior.Color = RGB(255, 192, 203) ' Pink
        
        ' Create validation list with location values
        With .Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:="SC Innsbruck,SC Dornbirn,SC Graz,SC Klagenfurt,SC Linz,SC WNeudorf,SC Deutschland"
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
    End With
    
    ' Create single field selector
    ws.Cells(FIELD_START_ROW - 7, FIELD_START_COL + 10).Value = "Single Field:"
    ws.Cells(FIELD_START_ROW - 7, FIELD_START_COL + 10).Font.Bold = True
    
    With ws.Range(SINGLE_FIELD_SELECTOR)
        ' Create a list of values from 1 to TOTAL_FIELDS
        Dim fieldList As String, i As Integer
        fieldList = ""
        For i = 1 To TOTAL_FIELDS
            fieldList = fieldList & i & ","
        Next i
        fieldList = Left(fieldList, Len(fieldList) - 1) ' Remove trailing comma
        
        ' Create validation list with field numbers
        With .Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:=fieldList
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
        .Value = "1"
    End With
    
    ' Create main control buttons
    AddButton ws, "Load_All_Fields", "Load All Fields", _
              ws.Cells(FIELD_START_ROW - 5, FIELD_START_COL).Address, _
              ws.Cells(FIELD_START_ROW - 5, FIELD_START_COL + 3).Address, _
              "LoadAllFields"
              
    AddButton ws, "Save_All_Fields", "Save All Fields", _
              ws.Cells(FIELD_START_ROW - 5, FIELD_START_COL + 5).Address, _
              ws.Cells(FIELD_START_ROW - 5, FIELD_START_COL + 8).Address, _
              "SaveAllFields"
              
    AddButton ws, "Load_Single_Field", "Load Single Field", _
              ws.Cells(FIELD_START_ROW - 5, FIELD_START_COL + 10).Address, _
              ws.Cells(FIELD_START_ROW - 5, FIELD_START_COL + 13).Address, _
              "LoadSingleField"
    
    ' Create status indicator
    ws.Cells(FIELD_START_ROW - 3, FIELD_START_COL).Value = "Status:"
    ws.Cells(FIELD_START_ROW - 3, FIELD_START_COL).Font.Bold = True
    
    ws.Cells(FIELD_START_ROW - 3, FIELD_START_COL + 1).Value = "Ready to load fields"
    ws.Range(ws.Cells(FIELD_START_ROW - 3, FIELD_START_COL + 1), _
             ws.Cells(FIELD_START_ROW - 3, FIELD_START_COL + 13)).Merge
End Sub

' Add a button with the given parameters
Private Sub AddButton(ws As Worksheet, btnName As String, btnCaption As String, _
                     topLeftCell As String, bottomRightCell As String, macroName As String)
    Dim btn As Button
    
    ' Add the button spanning the specified range
    Set btn = ws.Buttons.Add( _
        ws.Range(topLeftCell).Left, _
        ws.Range(topLeftCell).Top, _
        ws.Range(bottomRightCell).Left + ws.Range(bottomRightCell).Width - ws.Range(topLeftCell).Left, _
        ws.Range(bottomRightCell).Top + ws.Range(bottomRightCell).Height - ws.Range(topLeftCell).Top)
    
    ' Set button properties
    With btn
        .Name = btnName
        .Caption = btnCaption
        .OnAction = macroName
    End With
End Sub

' Setup the grid of fields
Private Sub SetupFieldGrid(ws As Worksheet)
    Dim row As Integer, col As Integer
    Dim fieldNum As Integer
    
    fieldNum = 1
    
    ' Create the field grid headers - column headers are letters, row headers are numbers
    For col = 1 To GRID_COLS
        ws.Cells(FIELD_START_ROW - 1, FIELD_START_COL + ((col - 1) * FIELD_COL_SPACING)).Value = Chr(64 + col) ' A, B, C...
        ws.Cells(FIELD_START_ROW - 1, FIELD_START_COL + ((col - 1) * FIELD_COL_SPACING)).Font.Bold = True
        ws.Cells(FIELD_START_ROW - 1, FIELD_START_COL + ((col - 1) * FIELD_COL_SPACING)).HorizontalAlignment = xlCenter
    Next col
    
    For row = 1 To GRID_ROWS
        ws.Cells(FIELD_START_ROW + ((row - 1) * FIELD_ROW_SPACING), FIELD_START_COL - 1).Value = row
        ws.Cells(FIELD_START_ROW + ((row - 1) * FIELD_ROW_SPACING), FIELD_START_COL - 1).Font.Bold = True
        ws.Cells(FIELD_START_ROW + ((row - 1) * FIELD_ROW_SPACING), FIELD_START_COL - 1).HorizontalAlignment = xlCenter
    Next row
    
    ' Create each field in the grid
    For row = 1 To GRID_ROWS
        For col = 1 To GRID_COLS
            CreateFieldInGrid ws, row, col, fieldNum
            fieldNum = fieldNum + 1
        Next col
    Next row
End Sub

' Create a single field within the grid
Private Sub CreateFieldInGrid(ws As Worksheet, gridRow As Integer, gridCol As Integer, fieldNum As Integer)
    Dim startRow As Integer, startCol As Integer
    Dim i As Integer, j As Integer
    
    ' Calculate the starting position for this field
    startRow = FIELD_START_ROW + ((gridRow - 1) * FIELD_ROW_SPACING)
    startCol = FIELD_START_COL + ((gridCol - 1) * FIELD_COL_SPACING)
    
    ' Create field number and type selector
    ws.Cells(startRow - 1, startCol).Value = "Field #" & fieldNum
    ws.Cells(startRow - 1, startCol).Font.Bold = True
    
    With ws.Cells(startRow - 1, startCol + 1)
        ' Create type dropdown
        With .Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:="Tour,WAB"
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
        .Value = "Tour" ' Default to Tour field
    End With
    
    ' Create the position selector
    With ws.Cells(startRow - 1, startCol + 2)
        .Value = "B" & (3 + ((fieldNum - 1) * 3)) ' Default position
    End With
    
    ' Create the field grid
    For i = 0 To FIELD_HEIGHT - 1
        For j = 0 To FIELD_WIDTH - 1
            With ws.Cells(startRow + i, startCol + j)
                .Value = ""
                .Interior.Color = RGB(220, 230, 241) ' Light blue background
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Size = 9
            End With
        Next j
    Next i
    
    ' Create a border around the field
    With ws.Range(ws.Cells(startRow, startCol), _
                  ws.Cells(startRow + FIELD_HEIGHT - 1, startCol + FIELD_WIDTH - 1)).Borders
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = 1
    End With
    
    ' Create a status indicator for this field
    ws.Cells(startRow + FIELD_HEIGHT, startCol).Value = "Status:"
    ws.Cells(startRow + FIELD_HEIGHT, startCol).Font.Size = 8
    
    ws.Cells(startRow + FIELD_HEIGHT, startCol + 1).Value = "Not loaded"
    ws.Cells(startRow + FIELD_HEIGHT, startCol + 1).Font.Size = 8
    ws.Cells(startRow + FIELD_HEIGHT, startCol + 1).Font.Italic = True
End Sub

' Load data for all fields
Public Sub LoadAllFields()
    Dim ws As Worksheet
    Dim targetSheet As String
    Dim row As Integer, col As Integer
    Dim fieldNum As Integer
    
    ' Get main worksheet
    Set ws = ThisWorkbook.Worksheets(MAIN_SHEET)
    
    ' Get the location selected in the dropdown
    targetSheet = ws.Range(LOCATION_CELL).Value
    If Trim(targetSheet) = "" Or Trim(targetSheet) = "Location" Then
        MsgBox "Please select a location first.", vbExclamation
        Exit Sub
    End If
    
    ' Convert location name to sheet name
    targetSheet = ConvertLocationToSheetName(targetSheet)
    
    ' Check if the sheet exists
    On Error Resume Next
    Dim targetWS As Worksheet
    Set targetWS = ThisWorkbook.Worksheets(targetSheet)
    On Error GoTo 0
    
    If targetWS Is Nothing Then
        MsgBox "Worksheet '" & targetSheet & "' not found. Please check the location selection.", vbExclamation
        Exit Sub
    End If
    
    ' Load each field in the grid
    fieldNum = 1
    For row = 1 To GRID_ROWS
        For col = 1 To GRID_COLS
            LoadFieldInGrid ws, targetWS, row, col, fieldNum
            fieldNum = fieldNum + 1
        Next col
    Next row
    
    ' Update status
    ws.Cells(FIELD_START_ROW - 3, FIELD_START_COL + 1).Value = "All fields loaded from " & targetSheet
    
    MsgBox "All fields loaded successfully from " & targetSheet & "!", vbInformation
End Sub

' Load a single field
Public Sub LoadSingleField()
    Dim ws As Worksheet
    Dim targetSheet As String
    Dim fieldNum As Integer
    Dim row As Integer, col As Integer
    
    ' Get main worksheet
    Set ws = ThisWorkbook.Worksheets(MAIN_SHEET)
    
    ' Get the field number to load
    fieldNum = Val(ws.Range(SINGLE_FIELD_SELECTOR).Value)
    If fieldNum < 1 Or fieldNum > TOTAL_FIELDS Then
        MsgBox "Please select a valid field number (1-" & TOTAL_FIELDS & ").", vbExclamation
        Exit Sub
    End If
    
    ' Get the location selected in the dropdown
    targetSheet = ws.Range(LOCATION_CELL).Value
    If Trim(targetSheet) = "" Or Trim(targetSheet) = "Location" Then
        MsgBox "Please select a location first.", vbExclamation
        Exit Sub
    End If
    
    ' Convert location name to sheet name
    targetSheet = ConvertLocationToSheetName(targetSheet)
    
    ' Check if the sheet exists
    On Error Resume Next
    Dim targetWS As Worksheet
    Set targetWS = ThisWorkbook.Worksheets(targetSheet)
    On Error GoTo 0
    
    If targetWS Is Nothing Then
        MsgBox "Worksheet '" & targetSheet & "' not found. Please check the location selection.", vbExclamation
        Exit Sub
    End If
    
    ' Convert field number to grid row and column
    row = ((fieldNum - 1) \ GRID_COLS) + 1
    col = ((fieldNum - 1) Mod GRID_COLS) + 1
    
    ' Load the field
    LoadFieldInGrid ws, targetWS, row, col, fieldNum
    
    ' Update status
    ws.Cells(FIELD_START_ROW - 3, FIELD_START_COL + 1).Value = "Field #" & fieldNum & " loaded from " & targetSheet
    
    MsgBox "Field #" & fieldNum & " loaded successfully from " & targetSheet & "!", vbInformation
End Sub

' Load a specific field in the grid
Private Sub LoadFieldInGrid(ws As Worksheet, targetWS As Worksheet, gridRow As Integer, gridCol As Integer, fieldNum As Integer)
    Dim startRow As Integer, startCol As Integer
    Dim fieldType As String, position As String
    Dim targetRow As Integer, targetCol As Integer
    Dim i As Integer, j As Integer
    
    ' Calculate the starting position for this field
    startRow = FIELD_START_ROW + ((gridRow - 1) * FIELD_ROW_SPACING)
    startCol = FIELD_START_COL + ((gridCol - 1) * FIELD_COL_SPACING)
    
    ' Get the field type and position
    fieldType = ws.Cells(startRow - 1, startCol + 1).Value
    position = ws.Cells(startRow - 1, startCol + 2).Value
    
    ' Skip if position is empty
    If Trim(position) = "" Then
        ws.Cells(startRow + FIELD_HEIGHT, startCol + 1).Value = "No position specified"
        Exit Sub
    End If
    
    ' Parse the position (e.g., "B3" -> col=2, row=3)
    targetCol = Asc(UCase(Left(position, 1))) - 64 ' Convert A->1, B->2, etc.
    targetRow = Val(Mid(position, 2))
    
    ' Load the data from the target sheet
    Dim rowCount As Integer
    rowCount = IIf(fieldType = "Tour", 3, 2) ' Tours are 3x3, WABs are 2x3
    
    For i = 0 To rowCount - 1
        For j = 0 To 2
            ws.Cells(startRow + i, startCol + j).Value = targetWS.Cells(targetRow + i, targetCol + j).Value
        Next j
    Next i
    
    ' If it's a WAB field, gray out the third row
    If fieldType = "WAB" Then
        For j = 0 To 2
            ws.Cells(startRow + 2, startCol + j).Value = ""
            ws.Cells(startRow + 2, startCol + j).Interior.Color = RGB(200, 200, 200) ' Gray
        Next j
    End If
    
    ' Update status for this field
    ws.Cells(startRow + FIELD_HEIGHT, startCol + 1).Value = "Loaded from " & targetWS.Name
End Sub

' Save data for all fields
Public Sub SaveAllFields()
    Dim ws As Worksheet
    Dim targetSheet As String
    Dim row As Integer, col As Integer
    Dim fieldNum As Integer
    
    ' Get main worksheet
    Set ws = ThisWorkbook.Worksheets(MAIN_SHEET)
    
    ' Get the location selected in the dropdown
    targetSheet = ws.Range(LOCATION_CELL).Value
    If Trim(targetSheet) = "" Or Trim(targetSheet) = "Location" Then
        MsgBox "Please select a location first.", vbExclamation
        Exit Sub
    End If
    
    ' Convert location name to sheet name
    targetSheet = ConvertLocationToSheetName(targetSheet)
    
    ' Check if the sheet exists
    On Error Resume Next
    Dim targetWS As Worksheet
    Set targetWS = ThisWorkbook.Worksheets(targetSheet)
    On Error GoTo 0
    
    If targetWS Is Nothing Then
        MsgBox "Worksheet '" & targetSheet & "' not found. Please check the location selection.", vbExclamation
        Exit Sub
    End If
    
    ' Save each field in the grid
    fieldNum = 1
    For row = 1 To GRID_ROWS
        For col = 1 To GRID_COLS
            SaveFieldInGrid ws, targetWS, row, col, fieldNum
            fieldNum = fieldNum + 1
        Next col
    Next row
    
    ' Update the header information with week and location
    With targetWS
        .Range("B1").Value = "Tourenplan für"
        .Range("G1").Value = "BML - NOS (I)"
        .Range("N1").Value = ws.Range(WEEK_CELL).Value
        .Range("AC1").Value = ws.Range(LOCATION_CELL).Value
    End With
    
    ' Update status
    ws.Cells(FIELD_START_ROW - 3, FIELD_START_COL + 1).Value = "All fields saved to " & targetSheet
    
    MsgBox "All fields saved successfully to " & targetSheet & "!", vbInformation
End Sub

' Save a specific field in the grid
Private Sub SaveFieldInGrid(ws As Worksheet, targetWS As Worksheet, gridRow As Integer, gridCol As Integer, fieldNum As Integer)
    Dim startRow As Integer, startCol As Integer
    Dim fieldType As String, position As String
    Dim targetRow As Integer, targetCol As Integer
    Dim i As Integer, j As Integer
    
    ' Calculate the starting position for this field
    startRow = FIELD_START_ROW + ((gridRow - 1) * FIELD_ROW_SPACING)
    startCol = FIELD_START_COL + ((gridCol - 1) * FIELD_COL_SPACING)
    
    ' Get the field type and position
    fieldType = ws.Cells(startRow - 1, startCol + 1).Value
    position = ws.Cells(startRow - 1, startCol + 2).Value
    
    ' Skip if position is empty
    If Trim(position) = "" Then
        Exit Sub
    End If
    
    ' Check if the field has been loaded
    If ws.Cells(startRow + FIELD_HEIGHT, startCol + 1).Value = "Not loaded" Then
        Exit Sub
    End If
    
    ' Parse the position (e.g., "B3" -> col=2, row=3)
    targetCol = Asc(UCase(Left(position, 1))) - 64 ' Convert A->1, B->2, etc.
    targetRow = Val(Mid(position, 2))
    
    ' Save the data to the target sheet
    Dim rowCount As Integer
    rowCount = IIf(fieldType = "Tour", 3, 2) ' Tours are 3x3, WABs are 2x3
    
    For i = 0 To rowCount - 1
        For j = 0 To 2
            targetWS.Cells(targetRow + i, targetCol + j).Value = ws.Cells(startRow + i, startCol + j).Value
        Next j
    Next i
    
    ' Update status for this field
    ws.Cells(startRow + FIELD_HEIGHT, startCol + 1).Value = "Saved to " & targetWS.Name
End Sub

' Helper function to convert location names to worksheet names
Private Function ConvertLocationToSheetName(locationName As String) As String
    Select Case locationName
        Case "SC Innsbruck"
            ConvertLocationToSheetName = "Tourenplan_BML_Innsbruck"
        Case "SC Dornbirn"
            ConvertLocationToSheetName = "Tourenplan_BML_Dornbirn"
        Case "SC Graz"
            ConvertLocationToSheetName = "Tourenplan_BML_Graz"
        Case "SC Klagenfurt"
            ConvertLocationToSheetName = "Tourenplan_BML_Klagenfurt"
        Case "SC Linz"
            ConvertLocationToSheetName = "Tourenplan_BML_Linz"
        Case "SC WNeudorf"
            ConvertLocationToSheetName = "Tourenplan_BML_WNeudorf"
        Case "SC Deutschland"
            ConvertLocationToSheetName = "Tourenplan_BML_Deutschland"
        Case Else
            ConvertLocationToSheetName = locationName
    End Select
End Function

' Main entry point to setup the grid system
Public Sub InstallGridSystem()
    SetupGridFieldSystem
End Sub
