Sub ImportCargoSupportData()
    ' Open a file dialog to select the Cargo-Support Excel file
    Dim filePath As String
    Dim cargoWB As Workbook
    
    ' Get the file path using a dialog
    filePath = GetCargoSupportFilePath()
    If filePath = "" Then Exit Sub  ' User canceled
    
    ' Open the Cargo-Support workbook
    Set cargoWB = Workbooks.Open(filePath, ReadOnly:=True)
    
    ' Process the data
    ProcessCargoSupportData cargoWB
    
    ' Close the workbook
    cargoWB.Close SaveChanges:=False
    
    MsgBox "Cargo-Support Daten wurden erfolgreich importiert.", vbInformation
End Sub

Function GetCargoSupportFilePath() As String
    ' Opens a file dialog to get the path to the Cargo-Support Excel file
    Dim fd As Office.FileDialog
    
    ' Create a FileDialog object as a File Picker
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "Bitte wählen Sie die Cargo-Support Excel-Datei"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel-Dateien", "*.xlsx;*.xls"
        
        ' Show the dialog box
        If .Show = True Then
            GetCargoSupportFilePath = .SelectedItems(1)
        Else
            GetCargoSupportFilePath = ""
        End If
    End With
End Function

Sub ProcessCargoSupportData(cargoWB As Workbook)
    ' Dictionary to store tour data
    ' Key = TourName|Date, Value = Collection of tour properties
    Dim tourDict As Object
    Set tourDict = CreateObject("Scripting.Dictionary")
    
    ' Reference to the worksheet
    Dim ws As Worksheet
    Set ws = cargoWB.Worksheets(1)  ' Assuming data is in the first sheet
    
    ' Find last row with data
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    
    ' Variables for current tour processing
    Dim currentTourName As String
    Dim currentTourNumber As String
    Dim currentABNumbers As New Collection
    Dim currentStops As Integer
    Dim currentWeight As Double
    Dim currentVolume As Double
    Dim tourDate As Date
    Dim isHeadline As Boolean
    Dim tourData As Collection
    
    ' Process each row
    Dim i As Long
    
    ' Initialize currentTourName (move this after declaring i)
    currentTourName = ""
    
    For i = 2 To lastRow  ' Assuming row 1 has headers
        ' Try to detect if this is a headline row
        ' In Cargo-Support exports, headlines are typically rows where StopReihenfolge (C) = 1 or empty and contain tour totals
        isHeadline = False
        
        ' Check if this is the start of a new tour (StopReihenfolge = 1 or empty and has a tour number)
        If (ws.Cells(i, "C").Value = 1 Or IsEmpty(ws.Cells(i, "C").Value)) And _
           Not IsEmpty(ws.Cells(i, "A").Value) And Not IsEmpty(ws.Cells(i, "B").Value) Then
            
            ' Save previous tour data if we have a complete set
            If currentTourName <> "" And Not IsEmpty(tourDate) Then
                ' Create a new collection for this tour
                Set tourData = New Collection
                
                ' Add the data to the collection
                tourData.Add currentTourNumber, "TourNumber"
                tourData.Add currentStops, "Stops"
                tourData.Add currentWeight, "Weight"
                tourData.Add currentVolume, "Volume"
                
                ' Add AB Numbers if any
                If currentABNumbers.count > 0 Then
                    ' Convert AB Numbers collection to a comma-separated string
                    Dim abString As String, ab As Variant
                    abString = ""
                    For Each ab In currentABNumbers
                        If abString <> "" Then abString = abString & ", "
                        abString = abString & CStr(ab)
                    Next ab
                    
                    tourData.Add abString, "ABNumbers"
                End If
                
                ' Add to dictionary with TourName|Date as key
                Dim dictKey As String
                dictKey = currentTourName & "|" & Format(tourDate, "yyyy-mm-dd")
                
                ' Add to dictionary if key doesn't exist
                If Not tourDict.Exists(dictKey) Then
                    tourDict.Add dictKey, tourData
                End If
            End If
            
            ' Start processing a new tour
            isHeadline = True
            currentTourName = NormalizeTourName(Trim(ws.Cells(i, "B").Value))
            currentTourNumber = Trim(ws.Cells(i, "A").Value)
            
            ' Reset counters for the new tour
            Set currentABNumbers = New Collection
            currentStops = 0
            currentWeight = 0
            currentVolume = 0
            
            ' Try to extract date from columns P (Abfahrt), Q (Ankunft) or R
            On Error Resume Next
            If IsDate(ws.Cells(i, "P").Value) Then
                tourDate = CDate(ws.Cells(i, "P").Value)
            ElseIf IsDate(ws.Cells(i, "Q").Value) Then
                tourDate = CDate(ws.Cells(i, "Q").Value)
            ElseIf IsDate(ws.Cells(i, "R").Value) Then
                tourDate = CDate(ws.Cells(i, "R").Value)
            Else
                ' If no date found, try to extract date from tour name or leave empty
                tourDate = Empty
            End If
            On Error GoTo 0
            
            ' Get total values from this headline row
            If Not IsEmpty(ws.Cells(i, "D").Value) Then
                currentWeight = CDbl(Replace(ws.Cells(i, "D").Value, ",", "."))
            End If
            
            If Not IsEmpty(ws.Cells(i, "E").Value) Then
                currentVolume = CDbl(Replace(ws.Cells(i, "E").Value, ",", "."))
            End If
        End If
        
        ' Count stops for this tour
        If currentTourName <> "" And Not isHeadline And Not IsEmpty(ws.Cells(i, "C").Value) Then
            currentStops = currentStops + 1
            
            ' Collect AB Numbers if present
            If Not IsEmpty(ws.Cells(i, "L").Value) Then
                Dim abNumber As String
                abNumber = Trim(ws.Cells(i, "L").Value)
                
                ' Check if this AB number is not already in the collection
                Dim abExists As Boolean, existingAB As Variant
                abExists = False
                
                For Each existingAB In currentABNumbers
                    If existingAB = abNumber Then
                        abExists = True
                        Exit For
                    End If
                Next existingAB
                
                If Not abExists And abNumber <> "" Then
                    currentABNumbers.Add abNumber
                End If
            End If
        End If
    Next i
    
    Debug.Print "Processed " & tourDict.count & " tours from Cargo Support data"
    Dim dbgKey As Variant
    For Each dbgKey In tourDict.keys
        Debug.Print "Cargo tour: " & dbgKey
    Next dbgKey
    
    ' Save the last tour if we have one
    If currentTourName <> "" And Not IsEmpty(tourDate) Then
        ' Create collection for the last tour
        Set tourData = New Collection
        tourData.Add currentTourNumber, "TourNumber"
        tourData.Add currentStops, "Stops"
        tourData.Add currentWeight, "Weight"
        tourData.Add currentVolume, "Volume"
        
        ' Add AB Numbers if any
        If currentABNumbers.count > 0 Then
            abString = ""
            For Each ab In currentABNumbers
                If abString <> "" Then abString = abString & ", "
                abString = abString & CStr(ab)
            Next ab
            
            tourData.Add abString, "ABNumbers"
        End If
        
        ' Add to dictionary
        dictKey = currentTourName & "|" & Format(tourDate, "yyyy-mm-dd")
        
        If Not tourDict.Exists(dictKey) Then
            tourDict.Add dictKey, tourData
        End If
    End If
    
    ' Store the cargo data dictionary
    StoreCargoSupportData tourDict
End Sub

Function NormalizeTourName(tourName As String) As String
    Dim result As String
    result = tourName
    
    ' Remove special characters that might cause matching issues
    result = Replace(result, ".", "")
    result = Replace(result, "-", " ")
    result = Replace(result, ",", " ")
    
    ' Remove multiple spaces
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    
    NormalizeTourName = Trim(result)
End Function

Sub StoreCargoSupportData(tourDict As Object)
    ' Store the cargo data in a hidden sheet for later use
    Dim ws As Worksheet
    
    ' Check if the hidden data sheet exists, create if not
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("CargoData")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = "CargoData"
        ws.Visible = xlSheetVeryHidden  ' Hide the sheet
    End If
    On Error GoTo 0
    
    ' Clear any existing data
    ws.Cells.Clear
    
    ' Add headers
    ws.Range("A1").Value = "TourName"
    ws.Range("B1").Value = "Date"
    ws.Range("C1").Value = "TourNumber"
    ws.Range("D1").Value = "Stops"
    ws.Range("E1").Value = "Weight"
    ws.Range("F1").Value = "Volume"
    ws.Range("G1").Value = "ABNumbers"
    
    ' Write the dictionary data to the sheet
    Dim row As Long, key As Variant, tourData As Collection
    row = 2
    
    For Each key In tourDict.keys
        ' Split the key to get tour name and date
        Dim keyParts As Variant
        keyParts = Split(key, "|")
        
        If UBound(keyParts) >= 1 Then
            ' Get the tour data collection
            Set tourData = tourDict(key)
            
            ' Write tour name and date
            ws.Cells(row, 1).Value = keyParts(0)  ' Tour name
            ws.Cells(row, 2).Value = keyParts(1)  ' Date
            
            ' Write tour data
            On Error Resume Next
            ws.Cells(row, 3).Value = tourData("TourNumber")
            ws.Cells(row, 4).Value = tourData("Stops")
            ws.Cells(row, 5).Value = tourData("Weight")
            ws.Cells(row, 6).Value = tourData("Volume")
            
            If ItemExists(tourData, "ABNumbers") Then
                ws.Cells(row, 7).Value = tourData("ABNumbers")
            End If
            On Error GoTo 0
            
            row = row + 1
        End If
    Next key
    
    ' Save the workbook to preserve the data
    ThisWorkbook.Save
End Sub

Function ItemExists(col As Collection, key As String) As Boolean
    ' Check if an item exists in a collection by key
    On Error Resume Next
    Dim temp As Variant
    temp = col(key)
    ItemExists = (Err.Number = 0)
    On Error GoTo 0
End Function

Function GetCargoSupportData() As Object
    ' Retrieve the cargo support data from the hidden sheet
    Dim tourDict As Object
    Set tourDict = CreateObject("Scripting.Dictionary")
    
    ' Check if the hidden data sheet exists
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("CargoData")
    If ws Is Nothing Then
        Set GetCargoSupportData = Nothing
        Exit Function
    End If
    On Error GoTo 0
    
    ' Find the last row with data
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    
    ' If there's only a header row, no data exists
    If lastRow <= 1 Then
        Set GetCargoSupportData = Nothing
        Exit Function
    End If
    
    ' Load the data into the dictionary
    Dim i As Long
    For i = 2 To lastRow
        ' Create a key from tour name and date
        Dim dictKey As String
        dictKey = ws.Cells(i, 1).Value & "|" & ws.Cells(i, 2).Value
        
        ' Create a collection for this tour
        Dim tourData As Collection
        Set tourData = New Collection
        
        ' Add the data to the collection
        tourData.Add ws.Cells(i, 3).Value, "TourNumber"
        
        On Error Resume Next
        If Not IsEmpty(ws.Cells(i, 4).Value) Then
            tourData.Add CInt(ws.Cells(i, 4).Value), "Stops"
        End If
        
        If Not IsEmpty(ws.Cells(i, 5).Value) Then
            tourData.Add CDbl(ws.Cells(i, 5).Value), "Weight"
        End If
        
        If Not IsEmpty(ws.Cells(i, 6).Value) Then
            tourData.Add CDbl(ws.Cells(i, 6).Value), "Volume"
        End If
        
        If Not IsEmpty(ws.Cells(i, 7).Value) Then
            tourData.Add ws.Cells(i, 7).Value, "ABNumbers"
        End If
        On Error GoTo 0
        
        ' Add to dictionary
        If Not tourDict.Exists(dictKey) Then
            tourDict.Add dictKey, tourData
        End If
    Next i
    
    ' Return the dictionary
    Set GetCargoSupportData = tourDict
End Function

Sub AddCargoSupportButton()
    ' Add Cargo-Support Import button
    On Error Resume Next
    Dim btnCargo As Button
    Set btnCargo = ActiveSheet.Buttons.Add(ActiveSheet.Range("V1").Left + 260, _
                                         ActiveSheet.Range("V1").Top + 60, 120, 20)
    With btnCargo
        .Caption = "Cargo-Daten"
        .OnAction = "ImportCargoSupportData"
        .Name = "btnCargoImport"
    End With
    
    ' Add TransferToursWithCargo button
    Dim btnTransferCargo As Button
    Set btnTransferCargo = ActiveSheet.Buttons.Add(ActiveSheet.Range("V1").Left + 260, _
                                                 ActiveSheet.Range("V1").Top + 90, 120, 20)
    With btnTransferCargo
        .Caption = "Transfer mit Cargo"
        .OnAction = "TransferToursWithCargo"
        .Name = "btnTransferCargo"
    End With
    On Error GoTo 0
End Sub

Sub TransferToursWithCargo()
    ' Get the cargo data dictionary
    Dim cargoData As Object
    Set cargoData = GetCargoSupportData()
    
    ' If no cargo data available, prompt to import it
    If cargoData Is Nothing Then
        Dim response As VbMsgBoxResult
        response = MsgBox("Keine Cargo-Support Daten verfügbar. Möchten Sie jetzt Daten importieren?", _
                         vbYesNo + vbQuestion)
        
        If response = vbYes Then
            ImportCargoSupportData
            Set cargoData = GetCargoSupportData()
            
            ' If still no data, exit
            If cargoData Is Nothing Then
                MsgBox "Keine Cargo-Support Daten verfügbar. Transfer abgebrochen.", vbExclamation
                Exit Sub
            End If
        Else
            ' User declined to import data, run standard transfer instead
            MsgBox "Transfer wird ohne Cargo-Daten durchgeführt.", vbInformation
            TransferTours
            Exit Sub
        End If
    End If
    
    ' Now call the enhanced transfer function with the cargo data
    TransferToursEnhanced cargoData
End Sub

Sub TransferToursEnhanced(cargoData As Object)
    ' Transfer tours using row mapping pattern with direct container logic
    ' and color coding as requested, enhanced with cargo data
    
    Dim mainWs As Worksheet
    Dim neudorfWs As Worksheet
    Dim mainCarRow As Integer, subCarRow As Integer
    Dim day As Integer, sourceCol As Integer, destCol As Integer
    Dim tourName As String, tourNumber As String, tourTime As String
    Dim workers As String, wabCount As Integer
    Dim hasNochPlatz As Boolean, isDirectContainer As Boolean
    Dim carsProcessed As Integer, toursAdded As Integer
    Dim directContainersUsed(6) As Integer ' Counter for direct containers for each day
    Dim startDate As Date ' Start date for the week (from B1)
    
    ' Get start date from B1
    On Error Resume Next
    If IsDate(ThisWorkbook.Sheets("NOS_Tourenkonzept").Range("B1").Value) Then
        startDate = ThisWorkbook.Sheets("NOS_Tourenkonzept").Range("B1").Value
    Else
        startDate = Date
    End If
    On Error GoTo 0
    
    ' Dictionary to track LLnr numbers (5-digit numbers) and their counts
    ' Key = 5-digit number, Value = count of appearances
    Dim llnrCountDict As Object
    Set llnrCountDict = CreateObject("Scripting.Dictionary")
    
    ' Dictionary to track LLnr numbers and their assigned colors
    ' Key = 5-digit number, Value = color
    Dim llnrColorDict As Object
    Set llnrColorDict = CreateObject("Scripting.Dictionary")
    
    ' Dictionary to track matched cargo tours and their real tour numbers
    ' Key = row|day, Value = TourInfo collection
    Dim matchedCargoDict As Object
    Set matchedCargoDict = CreateObject("Scripting.Dictionary")
    
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
    ' Clear comments from the range
    ClearCommentsFromRange neudorfWs.Range("B3:S33")
    
    neudorfWs.Range("B55:S72").ClearContents ' Clear direct container area
    neudorfWs.Range("B55:S72").Interior.colorIndex = xlNone
    neudorfWs.Range("B55:S72").Font.Bold = False
    ' Clear comments from the range
    ClearCommentsFromRange neudorfWs.Range("B55:S72")
    
    ' Also clear Lager container area
    neudorfWs.Range("B74:S81").ClearContents
    neudorfWs.Range("B74:S81").Interior.colorIndex = xlNone
    neudorfWs.Range("B74:S81").Font.Bold = False
    ' Clear comments from the range
    ClearCommentsFromRange neudorfWs.Range("B74:S81")
    
    ' Initialize counters
    carsProcessed = 0
    toursAdded = 0
    
    ' Initialize direct container counters
    For day = 0 To 6
        directContainersUsed(day) = 0
    Next day
    
    ' First pass - scan all tour numbers to identify 5-digit numbers
    ' and count how many times each appears,
    ' Also match tours with cargo data and store the matches
    mainCarRow = 5
    
    Do While mainCarRow <= 22
        ' Process all days for this car (0-4 for Mon-Fri)
        For day = 0 To 4
            ' Calculate column positions
            sourceCol = 2 + (day * 4)  ' B, F, J, N, R, V, Z
            
            ' Get tour number and name
            tourNumber = Trim(mainWs.Cells(mainCarRow, sourceCol).Value)
            tourName = Trim(mainWs.Cells(mainCarRow + 1, sourceCol).Value)
            
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
            
            ' Try to match with cargo data if there's a tour name
            If tourName <> "" Then
                ' Calculate the date for this day
                Dim tourDate As Date
                tourDate = startDate + day
                
                ' Try to find matching cargo data by TourName and Date
                Dim dictKey As String
                dictKey = tourName & "|" & Format(tourDate, "yyyy-mm-dd")
                
                ' If we find a match in cargo data
                If cargoData.Exists(dictKey) Then
                    ' Store the match in our dictionary for use in the final pass
                    Dim matchKey As String
                    matchKey = mainCarRow & "|" & day
                    
                    ' Create a collection to store all the matched info
                    Dim tourInfo As Collection
                    Set tourInfo = New Collection
                    
                    ' Add basic tour info
                    tourInfo.Add tourName, "TourName"
                    tourInfo.Add tourDate, "TourDate"
                    
                    ' Add cargo data
                    Dim cargoTourData As Collection
                    Set cargoTourData = cargoData(dictKey)
                    
                    On Error Resume Next
                    tourInfo.Add cargoTourData("TourNumber"), "RealTourNumber"
                    tourInfo.Add cargoTourData("Stops"), "Stops"
                    tourInfo.Add cargoTourData("Weight"), "Weight"
                    tourInfo.Add cargoTourData("Volume"), "Volume"
                    
                    If ItemExists(cargoTourData, "ABNumbers") Then
                        tourInfo.Add cargoTourData("ABNumbers"), "ABNumbers"
                    End If
                    On Error GoTo 0
                    
                    ' Store in matched dictionary
                    matchedCargoDict.Add matchKey, tourInfo
                End If
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
                
                ' Check if we have cargo data for this tour
                Dim cargoComment As String
                cargoComment = ""
                Dim realTourNumber As String
                realTourNumber = ""
                
                ' Create the lookup key for matched cargo data
                Dim cargoMatchKey As String
                cargoMatchKey = mainCarRow & "|" & day
                
                If matchedCargoDict.Exists(cargoMatchKey) Then
                    ' Get the cargo data
                    Dim matchedTourInfo As Collection
                    Set matchedTourInfo = matchedCargoDict(cargoMatchKey)
                    
                    ' Get the real tour number from cargo data
                    On Error Resume Next
                    realTourNumber = matchedTourInfo("RealTourNumber")
                    
                    ' Build cargo comment with additional information
                    cargoComment = "Cargo-Daten:" & vbCrLf
                    
                    If ItemExists(matchedTourInfo, "RealTourNumber") Then
                        cargoComment = cargoComment & "Tour-Nr.: " & matchedTourInfo("RealTourNumber") & vbCrLf
                    End If
                    
                    If ItemExists(matchedTourInfo, "Stops") Then
                        cargoComment = cargoComment & "Stops: " & matchedTourInfo("Stops") & vbCrLf
                    End If
                    
                    If ItemExists(matchedTourInfo, "Weight") Then
                        cargoComment = cargoComment & "Gewicht: " & Format(matchedTourInfo("Weight"), "#,##0.00") & " kg" & vbCrLf
                    End If
                    
                    If ItemExists(matchedTourInfo, "Volume") Then
                        cargoComment = cargoComment & "Volumen: " & Format(matchedTourInfo("Volume"), "#,##0.00") & " m³" & vbCrLf
                    End If
                    
                    If ItemExists(matchedTourInfo, "ABNumbers") Then
                        cargoComment = cargoComment & "AB-Nummern: " & matchedTourInfo("ABNumbers") & vbCrLf
                    End If
                    On Error GoTo 0
                    
                    ' Use the real tour number if available
                    If realTourNumber <> "" Then
                        ' Modify the tour number by adding the real tour number in parentheses
                        If InStr(tourNumber, realTourNumber) = 0 Then
                            tourNumber = realTourNumber & " (" & tourNumber & ")"
                        End If
                    End If
                End If
                
                ' Check if this should go to direct container area
                If isDirectContainer And directContainersUsed(day) < 6 Then
                    ' Add to direct container area
                    Dim directRow As Integer
                    directRow = 55 + (directContainersUsed(day) * 3)
                    
                    ' Copy tour data to direct container section with the unique direct tour color
                    CopyTourToDestinationEnhanced neudorfWs, directRow, destCol, tourName, tourNumber, _
                                         tourTime, workers, wabCount, hasNochPlatz, _
                                         isDirectContainer, llnrColor, matchingLLnr, wabNumber, directTourColor, _
                                         cargoComment
                    
                    ' Regular tour - copy to car's row with the same unique direct tour color
                    CopyTourToDestinationEnhanced neudorfWs, subCarRow, destCol, tourName, tourNumber, _
                                         tourTime, workers, wabCount, hasNochPlatz, _
                                         isDirectContainer, llnrColor, matchingLLnr, wabNumber, directTourColor, _
                                         cargoComment
                    
                    ' Increment direct container counter for this day
                    directContainersUsed(day) = directContainersUsed(day) + 1
                Else
                    ' Regular tour - copy to car's row
                    CopyTourToDestinationEnhanced neudorfWs, subCarRow, destCol, tourName, tourNumber, _
                                         tourTime, workers, wabCount, hasNochPlatz, _
                                         isDirectContainer, llnrColor, matchingLLnr, wabNumber, directTourColor, _
                                         cargoComment
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
                    Dim lgrIdx As Integer
                    lgrIdx = lgrWabCount(lgrDay)
                    
                    With lgrWabCell
                        .Interior.Color = lgrColors(lgrIdx Mod 4)
                        .Font.Bold = True
                        .HorizontalAlignment = xlCenter
                        .Borders.LineStyle = xlContinuous
                        .Borders.weight = xlThin
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
                    .Borders.weight = xlThin
                    .Borders.Color = RGB(0, 0, 0) ' Black border
                End With
            End If
        Next lgrIdx
    Next lgrDay
    ' END OF NEW PART
    
    ' Show results with cargo info
    Dim cargoMatches As Integer
    cargoMatches = matchedCargoDict.count
    
    MsgBox "Transfer mit Cargo-Daten abgeschlossen:" & vbCrLf & _
           "Fahrzeuge verarbeitet: " & carsProcessed & vbCrLf & _
           "Touren hinzugefügt: " & toursAdded & vbCrLf & _
           "Cargo-Daten integriert: " & cargoMatches, vbInformation
End Sub

Sub CopyTourToDestinationEnhanced(ws As Worksheet, row As Integer, col As Integer, tourName As String, _
                        tourNumber As String, tourTime As String, workers As String, _
                        wabCount As Integer, hasNochPlatz As Boolean, isDirectContainer As Boolean, _
                        Optional llnrColor As Long = -1, Optional llnrMatch As String = "", _
                        Optional wabNumbers As String = "", Optional directTourColor As Long = -1, _
                        Optional cargoComment As String = "")
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
        .weight = xlThin
        .colorIndex = xlAutomatic
    End With
    
    With tourFieldRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .weight = xlThin
        .colorIndex = xlAutomatic
    End With
    
    With tourFieldRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .weight = xlThin
        .colorIndex = xlAutomatic
    End With
    
    With tourFieldRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .weight = xlThin
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
            .Borders.weight = xlMedium
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
            .Borders.weight = xlThin
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
    
    ' Add cargo comment if provided
    If cargoComment <> "" Then
        ' Add the cargo data as a comment to the tour number cell
        Dim commentShape As Comment
        
        ' Check if the cell already has a comment
        On Error Resume Next
        If Not ws.Cells(row + 1, col).Comment Is Nothing Then
            ws.Cells(row + 1, col).Comment.Delete
        End If
        
        ' Add the comment
        Set commentShape = ws.Cells(row + 1, col).AddComment
        commentShape.Visible = False
        commentShape.Text Text:=cargoComment
        commentShape.Shape.TextFrame.AutoSize = True
        On Error GoTo 0
        
        ' Add a small visual indicator that this tour has cargo data (green triangle in corner)
        ws.Cells(row + 1, col).Interior.pattern = xlPatternDown
        ws.Cells(row + 1, col).Interior.PatternColorIndex = 4 ' Green
    End If
End Sub

' Add this GetNextColor function to your module
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

Sub ClearCommentsFromRange(rng As Range)
    ' Clear all comments from a range
    On Error Resume Next
    Dim cell As Range
    For Each cell In rng
        If Not cell.Comment Is Nothing Then
            cell.Comment.Delete
        End If
    Next cell
    On Error GoTo 0
End Sub
