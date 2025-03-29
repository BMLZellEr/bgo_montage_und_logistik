Sub ProcessToursAndCreatePDFs()
    ' Tour Processing Macro - Template-based version with PDF generation
    ' Version 4.0 - Uses predefined Excel templates for PDF generation with multi-page capability
    ' Creates tour-specific subfolders and a single log file
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim wsSummary As Worksheet
    Dim lastRow As Long, i As Long
    Dim tourNumber As String, lastTourNumber As String
    Dim tourName As String, tourDate As String
    Dim isServiceCenter As Boolean
    Dim pdfFolderPath As String
    Dim fso As Object
    Dim updateLog As Object
    
    ' Create FileSystemObject for folder operations
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Ask user for PDF output folder
    pdfFolderPath = BrowseForFolder("Select folder to save PDF files")
    If pdfFolderPath = "" Then
        MsgBox "Operation cancelled.", vbInformation
        Exit Sub
    End If
    
    ' Ensure folder path ends with backslash
    If Right(pdfFolderPath, 1) <> "\" Then
        pdfFolderPath = pdfFolderPath & "\"
    End If
    
    ' Create update log file
    Set updateLog = fso.CreateTextFile(pdfFolderPath & "Update_Log_" & Format(Now(), "yyyy-mm-dd_hh-mm-ss") & ".txt", True)
    updateLog.WriteLine "=== TOUR PROCESSING LOG ====================================="
    updateLog.WriteLine "Started: " & Now()
    updateLog.WriteLine "Output folder: " & pdfFolderPath
    updateLog.WriteLine "=========================================================="
    
    ' Show status message
    Application.StatusBar = "Processing tours. Please wait..."
    Application.ScreenUpdating = False
    
    ' Set reference to the active worksheet
    Set ws = ActiveSheet
    
    ' Check if summary sheet already exists, if not create it
    On Error Resume Next
    Set wsSummary = Worksheets("TourSummary")
    On Error GoTo ErrorHandler
    
    If wsSummary Is Nothing Then
        Set wsSummary = Worksheets.Add(After:=Worksheets(Worksheets.count))
        wsSummary.Name = "TourSummary"
    Else
        wsSummary.Cells.Clear
    End If
    
    ' Create headers in summary sheet with simplified columns
    With wsSummary
        .Cells(1, 1).Value = "Tour_Name"
        .Cells(1, 2).Value = "Tour_Date"
        .Cells(1, 3).Value = "Tour_Type" ' Direct or Service Center
        .Cells(1, 4).Value = "Total_Weight (kg)"
        .Cells(1, 5).Value = "Total_Volume (m³)"
        .Cells(1, 6).Value = "AB_Numbers"
        .Cells(1, 7).Value = "Items_Per_Stop" ' Combined column
        
        ' Format headers
        .Range(.Cells(1, 1), .Cells(1, 7)).Font.Bold = True
        .Range(.Cells(1, 1), .Cells(1, 7)).Interior.Color = RGB(200, 200, 200)
    End With
    
    ' Find the last row with data
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    
    ' First pass: Create the summary for each tour
    Dim summaryRow As Long
    summaryRow = 2
    
    ' Tour tracking collection
    Dim tourSummary As Object
    Set tourSummary = CreateObject("Scripting.Dictionary")
    
    updateLog.WriteLine "Building tour summary dictionary..."
    
    ' Process each row
    For i = 2 To lastRow
        ' Skip empty rows
        If Not IsEmpty(ws.Cells(i, 1).Value) Then
            tourNumber = Trim(CStr(ws.Cells(i, 1).Value))
            
            ' Only process rows with stop numbers (regular data rows)
            If IsNumeric(ws.Cells(i, 3).Value) Then
                ' Get tour data if it doesn't exist yet
                If Not tourSummary.Exists(tourNumber) Then
                    ' Create a new tour entry
                    Dim newTourData(7) As Variant ' Expanded to include the original tourText and date for folder naming
                    newTourData(0) = "" ' Tour name (filled below)
                    newTourData(1) = "" ' Tour date (filled below)
                    newTourData(2) = "" ' Tour type (filled below)
                    newTourData(3) = 0 ' Total weight - we'll calculate this
                    newTourData(4) = 0 ' Total volume - we'll calculate this
                    newTourData(5) = "" ' AB Numbers (filled below)
                    newTourData(6) = "" ' Original tour text from column B
                    newTourData(7) = "" ' System delivery date (for folder name)
                    
                    tourSummary.Add tourNumber, newTourData
                End If
                
                ' Get tour data array
                Dim tourDataArray As Variant
                tourDataArray = tourSummary(tourNumber)
                
                ' Store the original tour text from Column B if not already filled
                If tourDataArray(6) = "" And Not IsEmpty(ws.Cells(i, 2).Value) Then
                    tourDataArray(6) = CStr(ws.Cells(i, 2).Value)
                End If
                
                ' Get system date for naming if not already filled
                If tourDataArray(7) = "" And Not IsEmpty(ws.Cells(i, 18).Value) Then ' Column R - System Zustelldatum
                    If IsDate(ws.Cells(i, 18).Value) Then
                        tourDataArray(7) = CStr(ws.Cells(i, 18).Value)
                    End If
                End If
                
                ' Parse tour name and date if not already filled
                If tourDataArray(0) = "" And Not IsEmpty(ws.Cells(i, 2).Value) Then
                    Dim parsedTourName As String, parsedTourDate As String
                    ParseTourNameAndDate CStr(ws.Cells(i, 2).Value), parsedTourName, parsedTourDate
                    
                    tourDataArray(0) = parsedTourName
                    tourDataArray(1) = parsedTourDate
                    
                    ' Check if it's a Service Center tour
                    isServiceCenter = InStr(CStr(ws.Cells(i, 2).Value), "SC ") > 0
                    tourDataArray(2) = IIf(isServiceCenter, "Service Center", "Direct Tour")
                End If
                
                ' Add AB Number if not already included
                Dim abNumber As String
                abNumber = ws.Cells(i, 12).Value ' Column L
                
                If Len(abNumber) > 0 And InStr(tourDataArray(5), abNumber) = 0 Then
                    If Len(tourDataArray(5)) > 0 Then
                        tourDataArray(5) = tourDataArray(5) & ", "
                    End If
                    tourDataArray(5) = tourDataArray(5) & abNumber
                End If
                
                ' Add to the weight and volume totals
                ' Try to get weight and volume from columns D and E
                If IsNumeric(ws.Cells(i, 4).Value) Then
                    tourDataArray(3) = tourDataArray(3) + CDbl(ws.Cells(i, 4).Value)
                End If
                
                If IsNumeric(ws.Cells(i, 5).Value) Then
                    tourDataArray(4) = tourDataArray(4) + CDbl(ws.Cells(i, 5).Value)
                End If
                
                ' Update the tour data in the dictionary
                tourSummary(tourNumber) = tourDataArray
            End If
        End If
    Next i
    
    updateLog.WriteLine "Found " & tourSummary.count & " tours to process"
    
    ' Second pass: Fill the summary sheet with the gathered data
    Dim tourKey As Variant
    summaryRow = 2
    
    For Each tourKey In tourSummary.Keys
        Dim currentTourData As Variant
        currentTourData = tourSummary(tourKey)
        
        ' Fill basic tour info
        wsSummary.Cells(summaryRow, 1).Value = currentTourData(0) ' Tour name
        wsSummary.Cells(summaryRow, 2).Value = currentTourData(1) ' Tour date
        wsSummary.Cells(summaryRow, 3).Value = currentTourData(2) ' Tour type
        wsSummary.Cells(summaryRow, 4).Value = currentTourData(3) ' Total weight
        wsSummary.Cells(summaryRow, 5).Value = currentTourData(4) ' Total volume
        wsSummary.Cells(summaryRow, 6).Value = currentTourData(5) ' AB Numbers
        
        ' Now process the stops and their items for this tour
        Dim combinedItems As String
        combinedItems = ""

        ' Go through all rows to find stops for this tour
        For i = 2 To lastRow
            If Not IsEmpty(ws.Cells(i, 1).Value) And Trim(CStr(ws.Cells(i, 1).Value)) = tourKey Then
                ' Only process rows with stop numbers
                If IsNumeric(ws.Cells(i, 3).Value) Then
                    Dim stopNum As Long
                    stopNum = ws.Cells(i, 3).Value
            
                    ' Get "Packstück Artikeltypen" (column AU)
                    Dim artikelTypen As String
                    If Not IsEmpty(ws.Cells(i, 47).Value) Then ' Column AU
                        artikelTypen = CStr(ws.Cells(i, 47).Value)
                    Else
                        artikelTypen = ""
                    End If
            
                    ' Get "Warenbeschreibung" (column AV)
                    Dim warenText As String
                    If Not IsEmpty(ws.Cells(i, 48).Value) Then ' Column AV
                        warenText = CStr(ws.Cells(i, 48).Value)
                
                        ' Process and format the items with both inputs
                        Dim formattedItems As String
                        formattedItems = FormatItemsList(warenText, artikelTypen)
                
                        ' Add to combined items
                        combinedItems = combinedItems & "Stop " & stopNum & ":" & vbCrLf & formattedItems & vbCrLf & vbCrLf
                    End If
                End If
            End If
        Next i
        
        ' Add the processed item data to the summary
        wsSummary.Cells(summaryRow, 7).Value = combinedItems
        
        summaryRow = summaryRow + 1
    Next tourKey
    
    ' Format the summary sheet
    With wsSummary
        ' Format numbers
        For i = 2 To .Cells(.Rows.count, 1).End(xlUp).row
            .Cells(i, 4).NumberFormat = "#,##0.00"
            .Cells(i, 5).NumberFormat = "#,##0.00"
        Next i
        
        ' Format columns
        .Columns.AutoFit
        .Columns(7).ColumnWidth = 120 ' Make Items column wider
        
        ' Add borders
        .UsedRange.Borders.LineStyle = xlContinuous
        .UsedRange.Borders.weight = xlThin
    End With
    
    updateLog.WriteLine "Summary sheet created. Starting PDF generation..."
    
    ' Now create PDF for each tour using templates
    Dim tourCount As Long, successCount As Long, failCount As Long
    tourCount = 0
    successCount = 0
    failCount = 0
    
    ' PHASE 1: Create only Tour Summaries
    For Each tourKey In tourSummary.Keys
        tourCount = tourCount + 1
        updateLog.WriteLine "------------------------------------------------"
        updateLog.WriteLine "Processing tour " & tourCount & " of " & tourSummary.count & ": " & tourKey
        
        ' Get the tour data
        currentTourData = tourSummary(tourKey)
        
        ' Convert both tourKey and tourName to string to avoid ByRef error
        Dim tourKeyStr As String, tourNameStr As String, originalTourText As String, systemDate As String
        tourKeyStr = CStr(tourKey)
        tourNameStr = CStr(currentTourData(0))
        originalTourText = CStr(currentTourData(6))
        systemDate = CStr(currentTourData(7))
        
        ' Create subfolder name based on the pattern: mm_dd_$Tournummer_$Tourname
        Dim subFolderName As String, tourFolder As String
        
        ' Format date part (mm_dd)
        Dim datePart As String
        datePart = ""
        
        ' Try to get the date from system delivery date
        If IsDate(systemDate) Then
            datePart = Format(CDate(systemDate), "mm_dd")
        End If
        
        ' If no date found, use current date
        If datePart = "" Then
            datePart = Format(Date, "mm_dd")
        End If
        
        ' Clean the tour name for folder name (remove invalid characters)
        Dim cleanTourName As String
        cleanTourName = CleanForFileName(originalTourText)
        
        ' Create the folder name
        subFolderName = datePart & "_" & tourKeyStr & "_" & cleanTourName
        tourFolder = pdfFolderPath & subFolderName
        
        ' Try to create folder with full name first
        On Error Resume Next
        If Not fso.FolderExists(tourFolder) Then
            fso.CreateFolder (tourFolder)
            
            ' If error occurs, use simpler name
            If Err.Number <> 0 Then
                updateLog.WriteLine "  Using simpler folder name (error creating full name folder)"
                Err.Clear
                ' Fall back to simpler name: mm_dd_TourNumber
                subFolderName = datePart & "_" & tourKeyStr
                tourFolder = pdfFolderPath & subFolderName
                
                If Not fso.FolderExists(tourFolder) Then
                    fso.CreateFolder (tourFolder)
                End If
            End If
        End If
        
        ' Ensure folder path ends with backslash
        If Right(tourFolder, 1) <> "\" Then
            tourFolder = tourFolder & "\"
        End If
        On Error GoTo ErrorHandler
        
        ' Calculate the totalWeight and totalVolume for this specific tour
        Dim tourTotalWeight As Double, tourTotalVolume As Double
        tourTotalWeight = CDbl(currentTourData(3))
        tourTotalVolume = CDbl(currentTourData(4))
        
        updateLog.WriteLine "  Creating tour folder: " & subFolderName
        updateLog.WriteLine "  Weight: " & Format(tourTotalWeight, "#,##0.00") & " kg, Volume: " & Format(tourTotalVolume, "#,##0.00") & " m³"
        
        ' Create a collection for tourStops
        Dim tourStopCollection As New Collection
        For i = 2 To lastRow
            If ws.Cells(i, 1).Value = tourKeyStr And IsNumeric(ws.Cells(i, 3).Value) Then
                tourStopCollection.Add i
            End If
        Next i
        
        updateLog.WriteLine "  Found " & tourStopCollection.count & " stops"
        
        ' Now call ONLY the summary function
        On Error Resume Next
        CreateTourSummaryPDFFromTemplate ws, tourKeyStr, originalTourText, tourFolder, tourTotalWeight, tourTotalVolume, tourStopCollection, updateLog
        
        ' Check if an error occurred with PDF creation
        If Err.Number <> 0 Then
            updateLog.WriteLine "  ERROR creating Tour Summary PDF: " & Err.description
            Err.Clear
            failCount = failCount + 1
        Else
            successCount = successCount + 1
        End If
        
        On Error GoTo ErrorHandler
    Next tourKey
    
    ' PHASE 2: Create only Frachtbriefe - completely separate process
    updateLog.WriteLine "=========================================================="
    updateLog.WriteLine "Tour Summaries completed. Starting Frachtbrief creation..."
    updateLog.WriteLine "=========================================================="
    
    ' Force Excel to fully release resources before starting Frachtbrief creation
    Application.ScreenUpdating = True
    Application.Wait Now + TimeValue("00:00:01")
    Application.ScreenUpdating = False
    
    For Each tourKey In tourSummary.Keys
        ' Get the tour data again
        Dim frachtTourData As Variant
        frachtTourData = tourSummary(tourKey)
        
        ' Convert both tourKey and tourName to string again
        Dim frachtTourKeyStr As String, frachtTourNameStr As String, frachtOriginalTourText As String
        frachtTourKeyStr = CStr(tourKey)
        frachtTourNameStr = CStr(frachtTourData(0))
        frachtOriginalTourText = CStr(frachtTourData(6))
        
        ' Use the same folder path
        Dim frachtTourFolder As String
        Dim frachtDatePart As String, frachtCleanTourName As String, frachtSubFolderName As String
        
        ' Get same date part
        frachtDatePart = ""
        If IsDate(frachtTourData(7)) Then
            frachtDatePart = Format(CDate(frachtTourData(7)), "mm_dd")
        End If
        If frachtDatePart = "" Then
            frachtDatePart = Format(Date, "mm_dd")
        End If
        
        ' Prepare folder path
        frachtCleanTourName = CleanForFileName(frachtOriginalTourText)
        frachtSubFolderName = frachtDatePart & "_" & frachtTourKeyStr & "_" & frachtCleanTourName
        frachtTourFolder = pdfFolderPath & frachtSubFolderName
        
        ' Check if folder exists first (should be created from phase 1)
        On Error Resume Next
        If Not fso.FolderExists(frachtTourFolder) Then
            ' Try simpler name
            frachtSubFolderName = frachtDatePart & "_" & frachtTourKeyStr
            frachtTourFolder = pdfFolderPath & frachtSubFolderName
            
            ' If still not found, create it
            If Not fso.FolderExists(frachtTourFolder) Then
                fso.CreateFolder (frachtTourFolder)
            End If
        End If
        Err.Clear
        On Error GoTo ErrorHandler
        
        If Right(frachtTourFolder, 1) <> "\" Then
            frachtTourFolder = frachtTourFolder & "\"
        End If
        
        ' Recollect all stops for this tour
        Dim frachtTourStopCollection As New Collection
        For i = 2 To lastRow
            If ws.Cells(i, 1).Value = frachtTourKeyStr And IsNumeric(ws.Cells(i, 3).Value) Then
                frachtTourStopCollection.Add i
            End If
        Next i
        
        ' Create Frachtbrief PDF only
        updateLog.WriteLine "Creating Frachtbrief for tour " & frachtTourKeyStr & "..."
        On Error Resume Next
        CreateCombinedFreightPDFFromTemplate ws, frachtTourKeyStr, frachtOriginalTourText, frachtTourFolder, frachtTourStopCollection
        
        If Err.Number <> 0 Then
            updateLog.WriteLine "  ERROR creating Frachtbrief PDF: " & Err.description
            Err.Clear
        Else
            updateLog.WriteLine "  Successfully created Frachtbrief PDF"
        End If
        
        ' Force resource cleanup after each tour
        On Error Resume Next
        Application.Wait Now + TimeValue("00:00:01")
        Err.Clear
        
        On Error GoTo ErrorHandler
    Next tourKey
    
    ' Write summary to log
    updateLog.WriteLine "=========================================================="
    updateLog.WriteLine "Processing completed at: " & Now()
    updateLog.WriteLine "Total tours: " & tourCount
    updateLog.WriteLine "Successful PDFs: " & successCount
    updateLog.WriteLine "=========================================================="
    
    ' Close the log file
    updateLog.Close
    Set updateLog = Nothing
    
CleanExit:
    ' Reset status bar and screen updating
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "Tour summary created in the TourSummary worksheet." & vbCrLf & _
           "PDF files saved to subfolders in: " & pdfFolderPath & vbCrLf & vbCrLf & _
           "Summary: Total=" & tourCount & ", Success=" & successCount & ", Failed=" & failCount & vbCrLf & _
           "See Update_Log file for details.", vbInformation
    
    Exit Sub
    
ErrorHandler:
    ' Write to log if possible
    If Not updateLog Is Nothing Then
        updateLog.WriteLine "CRITICAL ERROR: " & Err.description & " (Line: " & Erl & ")"
        updateLog.Close
    End If
    
    MsgBox "An error occurred: " & Err.description & " (Line: " & Erl & ")", vbCritical
    Resume CleanExit
End Sub

Sub CreateTourSummaryPDFFromTemplate(ws As Worksheet, tourNumber As String, tourName As String, pdfPath As String, totalWeight As Double, totalVolume As Double, tourStops As Collection, ByRef updateLog As Object)
    ' Optimized solution that creates a single PDF with multiple pages for each tour
    ' Uses a similar approach to the Frachtbriefe PDF generation
    ' Creates a log file instead of showing messages
    
    On Error GoTo ErrorHandler
    
    ' Log start of processing this tour
    updateLog.WriteLine "  Processing tour summary PDF..."
    
    ' Create a new workbook for all the pages
    Dim tempWb As Workbook
    Set tempWb = Workbooks.Add
    
    ' Reference the template worksheet
    Dim templateWs As Worksheet
    On Error Resume Next
    Set templateWs = ws.Parent.Worksheets("Tour_Summary_Template")
    
    If templateWs Is Nothing Then
        ' Try Sheet1 as a fallback
        Set templateWs = ws.Parent.Worksheets("Sheet1")
        
        If templateWs Is Nothing Then
            updateLog.WriteLine "  ERROR: Template not found - cannot create summary PDF"
            Err.Clear
            On Error GoTo ErrorHandler
            Exit Sub
        End If
    End If
    On Error GoTo ErrorHandler
    
    ' Get tour information
    Dim firstStopRow As Long, fullTourName As String
    Dim vehicleInfo As String, deliveryDate As String
    Dim tourRoute As String, totalDistance As String
    Dim totalMontagezeit As Double
    
    firstStopRow = tourStops(1)
    
    ' Get full tour name
    fullTourName = ""
    If Not IsEmpty(ws.Cells(firstStopRow, 2).Value) Then
        fullTourName = CStr(ws.Cells(firstStopRow, 2).Value)
        While Right(fullTourName, 1) = "."
            fullTourName = Left(fullTourName, Len(fullTourName) - 1)
        Wend
    End If
    
    ' Get vehicle info
    vehicleInfo = ""
    If Not IsEmpty(ws.Cells(firstStopRow, 25).Value) Then
        vehicleInfo = ws.Cells(firstStopRow, 25).Value
    End If
    
    ' Get delivery date
    deliveryDate = ""
    If Not IsEmpty(ws.Cells(firstStopRow, 18).Value) Then
        If IsDate(ws.Cells(firstStopRow, 18).Value) Then
            deliveryDate = Format(ws.Cells(firstStopRow, 18).Value, "dddd, dd.MM.yyyy")
        Else
            deliveryDate = CStr(ws.Cells(firstStopRow, 18).Value)
        End If
    End If
    
    ' Get route info
    tourRoute = ""
    If Not IsEmpty(ws.Cells(firstStopRow, 52).Value) Then
        tourRoute = ws.Cells(firstStopRow, 52).Value
    End If
    
    ' Get distance
    totalDistance = ""
    If Not IsEmpty(ws.Cells(firstStopRow, 53).Value) Then
        If IsNumeric(ws.Cells(firstStopRow, 53).Value) Then
            totalDistance = Format(ws.Cells(firstStopRow, 53).Value, "#,##0.00") & " km"
        Else
            totalDistance = ws.Cells(firstStopRow, 53).Value
        End If
    End If
    
    ' Calculate total montagezeit
    totalMontagezeit = 0
    Dim j As Long
    For j = 1 To tourStops.count
        Dim currentRow As Long, montagezeit As Double
        currentRow = tourStops(j)
        montagezeit = 0
        If Not IsEmpty(ws.Cells(currentRow, 54).Value) And IsNumeric(ws.Cells(currentRow, 54).Value) Then
            montagezeit = CDbl(ws.Cells(currentRow, 54).Value)
        End If
        totalMontagezeit = totalMontagezeit + montagezeit
    Next j
    
    ' Maximum stops per page - carefully limited to avoid merging issues
    Dim maxStopsPerPage As Long
    maxStopsPerPage = 12 ' Conservative limit
    
    ' Calculate number of pages needed
    Dim pageCount As Long
    pageCount = Application.WorksheetFunction.Ceiling(tourStops.count / maxStopsPerPage, 1)
    
    ' Log number of pages
    updateLog.WriteLine "  Creating " & pageCount & " pages for " & tourStops.count & " stops"
    
    ' Process each page as a separate worksheet
    Dim i As Long
    For i = 1 To pageCount
        ' Calculate stops for this page
        Dim startStop As Long, endStop As Long
        startStop = ((i - 1) * maxStopsPerPage) + 1
        endStop = Application.WorksheetFunction.Min(i * maxStopsPerPage, tourStops.count)
        
        ' Copy the template to the new workbook
        On Error Resume Next
        templateWs.Copy After:=tempWb.Sheets(tempWb.Sheets.count)
        
        If Err.Number <> 0 Then
            ' If template copy fails, try simpler method
            updateLog.WriteLine "  Warning: Template copy failed for page " & i & " - using alternative method"
            Err.Clear
            tempWb.Sheets.Add After:=tempWb.Sheets(tempWb.Sheets.count)
        End If
        
        ' Rename the sheet
        On Error Resume Next
        tempWb.Sheets(tempWb.Sheets.count).Name = "Page_" & i
        If Err.Number <> 0 Then
            updateLog.WriteLine "  Warning: Could not rename sheet for page " & i
            Err.Clear
        End If
        On Error GoTo ErrorHandler
        
        ' Get the newly created sheet
        Dim tempWs As Worksheet
        Set tempWs = tempWb.Sheets(tempWb.Sheets.count)
        
        ' Replace header placeholders
        On Error Resume Next
        tempWs.UsedRange.Replace "[[TOUR_NAME]]", fullTourName, xlWhole
        tempWs.UsedRange.Replace "[[TOUR_NUMBER]]", tourNumber & " (Page " & i & " of " & pageCount & ")", xlWhole
        tempWs.UsedRange.Replace "[[Tour_DATE]]", deliveryDate, xlWhole
        tempWs.UsedRange.Replace "[[VEHICLE]]", vehicleInfo, xlWhole
        tempWs.UsedRange.Replace "[[Tourstrecke]]", tourRoute, xlWhole
        tempWs.UsedRange.Replace "[[Tour_Ges_Kilometer]]", totalDistance, xlWhole
        Err.Clear
        On Error GoTo ErrorHandler
        
        ' Find where to add stop data
        Dim dataStartRow As Long
        dataStartRow = 0
        For j = 1 To tempWs.UsedRange.Rows.count
            If InStr(1, tempWs.Cells(j, 1).Value, "[[Stop") > 0 Or _
               InStr(1, tempWs.Cells(j, 1).Value, "Stopps") > 0 Then
                dataStartRow = j + 1
                Exit For
            End If
        Next j
        
        If dataStartRow = 0 Then
            dataStartRow = 9 ' Default if not found
        End If
        
        ' Find column indices
        Dim colStop As Long, colKunde As Long, colVolumen As Long
        Dim colGewicht As Long, colMontZeit As Long, colStopZeit As Long
        
        colStop = 1    ' Default column positions
        colKunde = 2
        colVolumen = 3
        colGewicht = 4
        colMontZeit = 5
        colStopZeit = 6
        
        ' Locate column headers
        For j = 1 To tempWs.Cells(dataStartRow - 1, tempWs.Columns.count).End(xlToLeft).Column
            Select Case Trim(tempWs.Cells(dataStartRow - 1, j).Value)
                Case "Stopps"
                    colStop = j
                Case "Kunden"
                    colKunde = j
                Case "Volumen"
                    colVolumen = j
                Case "Gewicht"
                    colGewicht = j
                Case "Mont-Zeit"
                    colMontZeit = j
                Case "Stop-Zeit", "Avis-Zeit"
                    colStopZeit = j
            End Select
        Next j
        
        ' Change column header from "Stop-Zeit" to "Avis-Zeit" if needed
        On Error Resume Next
        For j = 1 To tempWs.Cells(dataStartRow - 1, tempWs.Columns.count).End(xlToLeft).Column
            If Trim(tempWs.Cells(dataStartRow - 1, j).Value) = "Stop-Zeit" Then
                tempWs.Cells(dataStartRow - 1, j).Value = "Avis-Zeit"
                Exit For
            End If
        Next j
        Err.Clear
        On Error GoTo ErrorHandler
        
        ' Clear existing stop data
        On Error Resume Next
        Dim lastRow As Long
        lastRow = tempWs.Cells(tempWs.Rows.count, colStop).End(xlUp).row
        If lastRow > dataStartRow Then
            tempWs.Rows(dataStartRow & ":" & lastRow).Delete
        End If
        Err.Clear
        On Error GoTo ErrorHandler
        
        ' Determine if Kunden column spans multiple columns
        Dim kundenEndCol As Long
        kundenEndCol = colKunde
        
        If colKunde = 2 Then
            kundenEndCol = 4
            If kundenEndCol > tempWs.Columns.count Then
                kundenEndCol = tempWs.Columns.count
            End If
        End If
        
        ' Process stops for this page
        Dim stopRow As Long
        stopRow = dataStartRow
        
        For j = startStop To endStop
            Dim stopIndex As Long
            stopIndex = j
            currentRow = tourStops(stopIndex)
            
            ' Extract stop data
            Dim stopNum As Long, recipientName As String, fullAddress As String
            Dim weight As Double, volume As Double, stopZeitRange As String
            
            ' Stop number
            stopNum = 0
            If IsNumeric(ws.Cells(currentRow, 3).Value) Then
                stopNum = ws.Cells(currentRow, 3).Value
            Else
                stopNum = stopIndex ' Use index as fallback
            End If
            
            ' Customer name
            recipientName = ""
            If Not IsEmpty(ws.Cells(currentRow, 36).Value) Then
                recipientName = CStr(ws.Cells(currentRow, 36).Value)
            End If
            
            ' Address
            Dim recipientAddress As String, recipientCity As String, recipientPostcode As String
            recipientAddress = ""
            recipientCity = ""
            recipientPostcode = ""
            
            If Not IsEmpty(ws.Cells(currentRow, 37).Value) Then
                recipientAddress = CStr(ws.Cells(currentRow, 37).Value)
            End If
            
            If Not IsEmpty(ws.Cells(currentRow, 38).Value) Then
                recipientCity = CStr(ws.Cells(currentRow, 38).Value)
            End If
            
            If Not IsEmpty(ws.Cells(currentRow, 39).Value) Then
                recipientPostcode = CStr(ws.Cells(currentRow, 39).Value)
            End If
            
            fullAddress = recipientAddress
            If Len(recipientPostcode) > 0 Or Len(recipientCity) > 0 Then
                If Len(fullAddress) > 0 Then
                    fullAddress = fullAddress & ", "
                End If
                fullAddress = fullAddress & recipientPostcode & " " & recipientCity
            End If
            
            ' Measurements
            weight = 0
            If IsNumeric(ws.Cells(currentRow, 4).Value) Then
                weight = CDbl(ws.Cells(currentRow, 4).Value)
            End If
            
            volume = 0
            If IsNumeric(ws.Cells(currentRow, 5).Value) Then
                volume = CDbl(ws.Cells(currentRow, 5).Value)
            End If
            
            montagezeit = 0
            If Not IsEmpty(ws.Cells(currentRow, 54).Value) And IsNumeric(ws.Cells(currentRow, 54).Value) Then
                montagezeit = CDbl(ws.Cells(currentRow, 54).Value)
            End If
            
            ' Time window
            stopZeitRange = ""
            If Not IsEmpty(ws.Cells(currentRow, 18).Value) Then
                If IsDate(ws.Cells(currentRow, 18).Value) Then
        Dim startTimeValue As Date, endTimeValue As Date
                    On Error Resume Next
                    startTimeValue = CDate(Format(ws.Cells(currentRow, 18).Value, "hh:mm"))
                    
                    If Err.Number = 0 Then
                        ' Round to nearest hour
                        Dim startHour As Integer
                        startHour = Hour(startTimeValue)
                        startTimeValue = DateSerial(Year(Date), Month(Date), Day(Date)) + TimeSerial(startHour, 0, 0)
                        
                        ' Add 3 hours
                        endTimeValue = DateAdd("h", 3, startTimeValue)
                        
                        ' Format
                        stopZeitRange = Format(startTimeValue, "hh:mm") & " - " & Format(endTimeValue, "hh:mm")
                    End If
                    Err.Clear
                End If
            End If
            On Error GoTo ErrorHandler
            
            ' Insert rows for this stop if not the first one
            If j > startStop Then
                On Error Resume Next
                tempWs.Rows(stopRow & ":" & (stopRow + 1)).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                
                If Err.Number <> 0 Then
                    ' Try simpler insert
                    Err.Clear
                    tempWs.Rows(stopRow & ":" & (stopRow + 1)).Insert Shift:=xlDown
                End If
                Err.Clear
                On Error GoTo ErrorHandler
            End If
            
            ' Fill in cell values for first row
            tempWs.Cells(stopRow, colStop).Value = "Stop " & stopNum
            tempWs.Cells(stopRow, colKunde).Value = recipientName
            
            ' Format values with validation
            If volume > 0 Then
                tempWs.Cells(stopRow, colVolumen).Value = Format(volume, "#,##0.00") & " m³"
            Else
                tempWs.Cells(stopRow, colVolumen).Value = "0,00 m³"
            End If
            
            If weight > 0 Then
                tempWs.Cells(stopRow, colGewicht).Value = Format(weight, "#,##0.00") & " kg"
            Else
                tempWs.Cells(stopRow, colGewicht).Value = "0,00 kg"
            End If
            
            If montagezeit > 0 Then
                tempWs.Cells(stopRow, colMontZeit).Value = Format(montagezeit, "#,##0.00") & " h"
            Else
                tempWs.Cells(stopRow, colMontZeit).Value = "0,00 h"
            End If
            
            If colStopZeit > 0 Then
                tempWs.Cells(stopRow, colStopZeit).Value = stopZeitRange
            End If
            
            ' Fill in second row
            tempWs.Cells(stopRow + 1, colStop).Value = ""
            tempWs.Cells(stopRow + 1, colKunde).Value = fullAddress
            tempWs.Cells(stopRow + 1, colVolumen).Value = ""
            tempWs.Cells(stopRow + 1, colGewicht).Value = ""
            tempWs.Cells(stopRow + 1, colMontZeit).Value = ""
            If colStopZeit > 0 Then
                tempWs.Cells(stopRow + 1, colStopZeit).Value = ""
            End If
            
            ' Try merging cells with independent error handling for each merge
            ' 1. Merge stop column
            On Error Resume Next
            tempWs.Range(tempWs.Cells(stopRow, colStop), tempWs.Cells(stopRow + 1, colStop)).Merge
            Err.Clear
            
            ' 2. Merge customer name column
            On Error Resume Next
            If kundenEndCol > colKunde Then
                tempWs.Range(tempWs.Cells(stopRow, colKunde), tempWs.Cells(stopRow, kundenEndCol)).Merge
            End If
            Err.Clear
            
            ' 3. Merge customer address column
            On Error Resume Next
            If kundenEndCol > colKunde Then
                tempWs.Range(tempWs.Cells(stopRow + 1, colKunde), tempWs.Cells(stopRow + 1, kundenEndCol)).Merge
            End If
            Err.Clear
            
            ' 4. Merge measurement columns vertically
            On Error Resume Next
            tempWs.Range(tempWs.Cells(stopRow, colVolumen), tempWs.Cells(stopRow + 1, colVolumen)).Merge
            Err.Clear
            
            On Error Resume Next
            tempWs.Range(tempWs.Cells(stopRow, colGewicht), tempWs.Cells(stopRow + 1, colGewicht)).Merge
            Err.Clear
            
            On Error Resume Next
            tempWs.Range(tempWs.Cells(stopRow, colMontZeit), tempWs.Cells(stopRow + 1, colMontZeit)).Merge
            Err.Clear
            
            On Error Resume Next
            If colStopZeit > 0 Then
                tempWs.Range(tempWs.Cells(stopRow, colStopZeit), tempWs.Cells(stopRow + 1, colStopZeit)).Merge
            End If
            Err.Clear
            
            ' Apply borders
            On Error Resume Next
            With tempWs.Range(tempWs.Cells(stopRow, colStop), tempWs.Cells(stopRow + 1, IIf(colStopZeit > 0, colStopZeit, colMontZeit)))
                .Borders.LineStyle = xlContinuous
                .Borders.weight = xlThin
            End With
            Err.Clear
            
            ' Remove inner borders
            On Error Resume Next
            If kundenEndCol > colKunde Then
                tempWs.Range(tempWs.Cells(stopRow, colKunde), tempWs.Cells(stopRow, kundenEndCol)).Borders(xlEdgeBottom).LineStyle = xlNone
                tempWs.Range(tempWs.Cells(stopRow + 1, colKunde), tempWs.Cells(stopRow + 1, kundenEndCol)).Borders(xlEdgeTop).LineStyle = xlNone
            End If
            Err.Clear
            
            ' Set alignment - stop column
            On Error Resume Next
            With tempWs.Range(tempWs.Cells(stopRow, colStop), tempWs.Cells(stopRow + 1, colStop))
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            Err.Clear
            
            ' Set alignment - measurements
            On Error Resume Next
            With tempWs.Range(tempWs.Cells(stopRow, colVolumen), tempWs.Cells(stopRow + 1, IIf(colStopZeit > 0, colStopZeit, colMontZeit)))
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            Err.Clear
            
            ' Set alignment - customer name and address
            On Error Resume Next
            With tempWs.Range(tempWs.Cells(stopRow, colKunde), tempWs.Cells(stopRow, IIf(kundenEndCol > colKunde, kundenEndCol, colKunde)))
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlTop
            End With
            
            With tempWs.Range(tempWs.Cells(stopRow + 1, colKunde), tempWs.Cells(stopRow + 1, IIf(kundenEndCol > colKunde, kundenEndCol, colKunde)))
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlTop
            End With
            Err.Clear
            
            On Error GoTo ErrorHandler
            
            ' Move to next row
            stopRow = stopRow + 2
        Next j
        
        ' Add final row - totals on the last page, "continued" on others
        If i = pageCount Then
            ' Add totals row on last page
            On Error Resume Next
            tempWs.Cells(stopRow, colStop).Value = "Total"
            tempWs.Cells(stopRow, colKunde).Value = ""
            tempWs.Cells(stopRow + 1, colKunde).Value = ""
            tempWs.Cells(stopRow, colVolumen).Value = Format(totalVolume, "#,##0.00") & " m³"
            tempWs.Cells(stopRow, colGewicht).Value = Format(totalWeight, "#,##0.00") & " kg"
            tempWs.Cells(stopRow, colMontZeit).Value = Format(totalMontagezeit, "#,##0.00") & " h"
            Err.Clear
            
            ' Merge cells for total row
            On Error Resume Next
            tempWs.Range(tempWs.Cells(stopRow, colStop), tempWs.Cells(stopRow + 1, colStop)).Merge
            tempWs.Range(tempWs.Cells(stopRow, colKunde), tempWs.Cells(stopRow + 1, colKunde)).Merge
            tempWs.Range(tempWs.Cells(stopRow, colVolumen), tempWs.Cells(stopRow + 1, colVolumen)).Merge
            tempWs.Range(tempWs.Cells(stopRow, colGewicht), tempWs.Cells(stopRow + 1, colGewicht)).Merge
            tempWs.Range(tempWs.Cells(stopRow, colMontZeit), tempWs.Cells(stopRow + 1, colMontZeit)).Merge
            
            If colStopZeit > 0 Then
                tempWs.Range(tempWs.Cells(stopRow, colStopZeit), tempWs.Cells(stopRow + 1, colStopZeit)).Merge
            End If
            Err.Clear
            
            ' Format total row
            On Error Resume Next
            With tempWs.Range(tempWs.Cells(stopRow, colStop), tempWs.Cells(stopRow + 1, IIf(colStopZeit > 0, colStopZeit, colMontZeit)))
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
                .Borders.weight = xlThick
                .Interior.ColorIndex = xlNone
            End With
            Err.Clear
        Else
            ' Add "continued" note on non-final pages
            On Error Resume Next
            tempWs.Cells(stopRow, colStop).Value = "Continued on next page..."
            tempWs.Cells(stopRow, colKunde).Value = ""
            
            tempWs.Range(tempWs.Cells(stopRow, colStop), tempWs.Cells(stopRow, IIf(colStopZeit > 0, colStopZeit, colMontZeit))).Merge
            
            With tempWs.Cells(stopRow, colStop)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Italic = True
                .Font.Bold = True
            End With
            Err.Clear
        End If
        
        On Error GoTo ErrorHandler
        
        ' Set print settings
        With tempWs.PageSetup
            .Orientation = xlPortrait
            .PaperSize = xlPaperA4
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
            .CenterHorizontally = True
            .CenterVertically = False
        End With
    Next i
    
    ' Delete the first blank sheet
    Application.DisplayAlerts = False
    If tempWb.Sheets.count > pageCount Then
        tempWb.Sheets(1).Delete
    End If
    Application.DisplayAlerts = True
    
    ' Save as a single PDF file
    Dim pdfFileName As String
    pdfFileName = pdfPath & "Tour_" & tourNumber & "_Summary.pdf"
    
    ' Export as a single PDF - similar to Frachtbriefe approach
    On Error Resume Next
    tempWb.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfFileName, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        
    If Err.Number <> 0 Then
        ' Log error and try alternate method
        updateLog.WriteLine "  Warning: PDF export failed with error " & Err.Number & " - " & Err.description
        updateLog.WriteLine "  Attempting alternative PDF export method..."
        Err.Clear
        
        ' Try PrintOut method instead
        tempWb.PrintOut Copies:=1, Preview:=False, ActivePrinter:="Microsoft Print to PDF", _
            PrintToFile:=True, Collate:=True, PrToFileName:=pdfFileName
            
        If Err.Number <> 0 Then
            ' If all PDF methods fail, try saving as Excel
            updateLog.WriteLine "  Warning: All PDF methods failed - saving as Excel file"
            Err.Clear
            
            Dim xlsFileName As String
            xlsFileName = pdfPath & "Tour_" & tourNumber & "_Summary.xlsx"
            
            tempWb.SaveAs Filename:=xlsFileName, FileFormat:=xlOpenXMLWorkbook
            
            If Err.Number <> 0 Then
                updateLog.WriteLine "  ERROR: Could not save even as Excel - using text fallback"
                Err.Clear
                SpecialHandlingForTour ws, pdfPath, tourNumber
            Else
                updateLog.WriteLine "  Created Excel file: " & xlsFileName
            End If
        Else
            updateLog.WriteLine "  Created PDF using alternative method: " & pdfFileName
        End If
    Else
        updateLog.WriteLine "  Successfully created PDF: " & pdfFileName
    End If
    Err.Clear
    
    ' Close the workbook
    On Error Resume Next
    tempWb.Close SaveChanges:=False
    If Err.Number <> 0 Then
        updateLog.WriteLine "  Warning: Error closing workbook: " & Err.description
        Err.Clear
    End If
    
    updateLog.WriteLine "  Completed tour summary PDF processing"
    Exit Sub
    
ErrorHandler:
    ' Log the error in the update file
    Dim errMsg As String
    errMsg = "  ERROR in tour " & tourNumber & ": " & Err.description
    
    ' Add line number if available
    If Erl <> 0 Then
        errMsg = errMsg & " (Line: " & Erl & ")"
    End If
    
    updateLog.WriteLine errMsg
    
    ' Clean up any open workbooks
    On Error Resume Next
    If Not tempWb Is Nothing Then
        If tempWb.Name <> "" Then
            tempWb.Close SaveChanges:=False
        End If
    End If
    Err.Clear
    
    ' Use special handling as fallback
    updateLog.WriteLine "  Using text fallback method for tour " & tourNumber
    SpecialHandlingForTour ws, pdfPath, tourNumber
End Sub

Sub CreateCombinedFreightPDFFromTemplate(ws As Worksheet, tourNumber As String, tourName As String, pdfPath As String, tourStops As Collection)
    ' Create a combined PDF for all stops in a tour using the template
    On Error GoTo ErrorHandler
    
    Dim templateWs As Worksheet
    Dim tempWb As Workbook
    Dim stopCount As Long
    Dim i As Long, j As Long
    Dim pdfFileNameFracht As String
    
    ' Get stop count
    stopCount = tourStops.count
    
    ' Initialize randomizer for unique IDs
    Randomize
    
    ' Generate a unique identifier for this run - adding randomness helps prevent collisions
    Dim uniqueID As String
    uniqueID = Format(Now(), "yyyymmdd_hhmmss") & "_" & CStr(Int((10000 * Rnd()) + 1))
    
    ' Reference the template worksheet - make sure to use the SAME workbook where the data is
    On Error Resume Next
    Set templateWs = ws.Parent.Worksheets("Frachtbrief_Template")
    
    If templateWs Is Nothing Then
        ' Try Sheet2 or Sheet3 as a fallback
        Set templateWs = ws.Parent.Worksheets("Sheet2")
        
        If templateWs Is Nothing Then
            Set templateWs = ws.Parent.Worksheets("Sheet3")
            
            If templateWs Is Nothing Then
                MsgBox "Stop Freight template sheet not found in this workbook! Please add a sheet named 'Frachtbrief_Template'.", vbCritical
                Exit Sub
            End If
        End If
    End If
    On Error GoTo ErrorHandler
    
    ' Create a new temporary workbook name
    Dim tempWbName As String
    tempWbName = "TempWB_" & tourNumber & "_" & uniqueID
    
    ' Create a new workbook with a unique name
    Set tempWb = Workbooks.Add
    tempWb.Title = tempWbName
    
    ' Process each stop and add to the workbook
    For j = 1 To tourStops.count
        Dim currentRow As Long
        currentRow = tourStops(j)
        Dim stopNum As Long
        stopNum = ws.Cells(currentRow, 3).Value ' Column C
        
        ' Copy the template to the new workbook
        templateWs.Copy After:=tempWb.Sheets(tempWb.Sheets.count)
        
        ' Rename the newly added sheet
        tempWb.Sheets(tempWb.Sheets.count).Name = "Stop_" & stopNum
        
        ' Get the newly created sheet
        Dim tempWs As Worksheet
        Set tempWs = tempWb.Sheets("Stop_" & stopNum)
        
        ' Get all the data for this stop
        Dim abNumber As String, artikelTypen As String, warenText As String
        Dim recipientName As String, recipientAddress As String, recipientCity As String, recipientPostcode As String
        Dim nosContact As String, nosPhone As String, nosEmail As String
        Dim deliveryDate As String, serviceType As String
        Dim abInfo As String, buildingInfo As String, importantInfo As String, deliveryInfo As String
        Dim stopWeight As Double, stopVolume As Double, montagezeit As Double
        Dim totalQuantity As String

        ' Extract all the data for this stop from the source worksheet
        abNumber = ws.Cells(currentRow, 12).Value ' Column L
        
        ' Get the total quantity from column AI (35)
        totalQuantity = ""
        If Not IsEmpty(ws.Cells(currentRow, 35).Value) Then
            totalQuantity = CStr(ws.Cells(currentRow, 35).Value)
        End If
        
        ' Check if columns exist before trying to access
        artikelTypen = ""
        If currentRow > 0 And ws.Columns.count >= 47 Then
            If Not IsEmpty(ws.Cells(currentRow, 47).Value) Then
                artikelTypen = CStr(ws.Cells(currentRow, 47).Value)
            End If
        End If
        
        warenText = ""
        If currentRow > 0 And ws.Columns.count >= 48 Then
            If Not IsEmpty(ws.Cells(currentRow, 48).Value) Then
                warenText = CStr(ws.Cells(currentRow, 48).Value)
            End If
        End If
        
        ' Basic stop info
        stopWeight = 0
        If IsNumeric(ws.Cells(currentRow, 4).Value) Then ' Column D - Weight
            stopWeight = ws.Cells(currentRow, 4).Value
        End If
        
        stopVolume = 0
        If IsNumeric(ws.Cells(currentRow, 5).Value) Then ' Column E - Volume
            stopVolume = ws.Cells(currentRow, 5).Value
        End If
        
        ' Customer info
        recipientName = ""
        If Not IsEmpty(ws.Cells(currentRow, 36).Value) Then ' Column AI - Empfänger
            recipientName = ws.Cells(currentRow, 36).Value
        End If
        
        recipientAddress = ""
        If Not IsEmpty(ws.Cells(currentRow, 37).Value) Then ' Column AJ - Empf. Str.
            recipientAddress = ws.Cells(currentRow, 37).Value
        End If
        
        recipientCity = ""
        If Not IsEmpty(ws.Cells(currentRow, 38).Value) Then ' Column AK - Empf. Ort
            recipientCity = ws.Cells(currentRow, 38).Value
        End If
        
        recipientPostcode = ""
        If Not IsEmpty(ws.Cells(currentRow, 39).Value) Then ' Column AL - Empf. Plz
            recipientPostcode = ws.Cells(currentRow, 39).Value
        End If
        
        ' NOS contact info (columns BK, BL, BM as per screenshot)
        nosContact = ""
        If currentRow > 0 And ws.Columns.count >= 63 Then
            If Not IsEmpty(ws.Cells(currentRow, 63).Value) Then
                nosContact = CStr(ws.Cells(currentRow, 63).Value)
            End If
        End If
        
        nosPhone = ""
        If currentRow > 0 And ws.Columns.count >= 64 Then
            If Not IsEmpty(ws.Cells(currentRow, 64).Value) Then
                nosPhone = CStr(ws.Cells(currentRow, 64).Value)
            End If
        End If
        
        nosEmail = ""
        If currentRow > 0 And ws.Columns.count >= 65 Then
            If Not IsEmpty(ws.Cells(currentRow, 65).Value) Then
                nosEmail = CStr(ws.Cells(currentRow, 65).Value)
            End If
        End If
        
        ' Service and delivery info
        serviceType = ""
        If Not IsEmpty(ws.Cells(currentRow, 10).Value) Then ' Column J - ServiceTyp
            serviceType = ws.Cells(currentRow, 10).Value
        End If
        
        ' Montagezeit info - check if column exists
        montagezeit = 0
        If currentRow > 0 And ws.Columns.count >= 54 Then
            If Not IsEmpty(ws.Cells(currentRow, 54).Value) And IsNumeric(ws.Cells(currentRow, 54).Value) Then
                montagezeit = CDbl(ws.Cells(currentRow, 54).Value)
            End If
        End If
        
        ' Get delivery date and time from System Zustelldatum (Column R)
        deliveryDate = ""
        If Not IsEmpty(ws.Cells(currentRow, 18).Value) Then ' Column R - System Zustelldatum
            If IsDate(ws.Cells(currentRow, 18).Value) Then
                deliveryDate = Format(ws.Cells(currentRow, 18).Value, "dd.MM.yyyy HH:mm")
            Else
                deliveryDate = CStr(ws.Cells(currentRow, 18).Value)
            End If
        End If
        
        ' Additional information fields - check if columns exist
        abInfo = ""
        If currentRow > 0 And ws.Columns.count >= 42 Then
            If Not IsEmpty(ws.Cells(currentRow, 42).Value) Then
                abInfo = CStr(ws.Cells(currentRow, 42).Value)
            End If
        End If
        
        buildingInfo = ""
        If currentRow > 0 And ws.Columns.count >= 43 Then
            If Not IsEmpty(ws.Cells(currentRow, 43).Value) Then
                buildingInfo = CStr(ws.Cells(currentRow, 43).Value)
            End If
        End If
        
        importantInfo = ""
        If currentRow > 0 And ws.Columns.count >= 44 Then
            If Not IsEmpty(ws.Cells(currentRow, 44).Value) Then
                importantInfo = CStr(ws.Cells(currentRow, 44).Value)
            End If
        End If
        
        deliveryInfo = ""
        If currentRow > 0 And ws.Columns.count >= 45 Then
            If Not IsEmpty(ws.Cells(currentRow, 45).Value) Then
                deliveryInfo = CStr(ws.Cells(currentRow, 45).Value)
            End If
        End If
        
        ' Replace all placeholders in the template
        tempWs.UsedRange.Replace "[[TOUR_NUMBER]]", tourNumber, xlWhole
        tempWs.UsedRange.Replace "[[TOUR_NAME]]", tourName, xlWhole
        tempWs.UsedRange.Replace "[[Stop 1]]", "Stop " & stopNum, xlWhole
        tempWs.UsedRange.Replace "[[Kunde_Name 1]]", recipientName, xlWhole
        
        ' Neuer Platzhalter für Tournummer + Stop
        tempWs.UsedRange.Replace "[[Tournummer + Stop]]", tourNumber & " - Stop " & stopNum, xlWhole
        
        ' Correct the placeholder names for volume, weight, and time
        tempWs.UsedRange.Replace "[[Stop_1_V]]", Format(stopVolume, "#,##0.00") & " m³", xlWhole
        tempWs.UsedRange.Replace "[[Stop_1_W]]", Format(stopWeight, "#,##0.00") & " kg", xlWhole
        tempWs.UsedRange.Replace "[[Stop_1_Z]]", Format(montagezeit, "#,##0.00") & " h", xlWhole
        
        tempWs.UsedRange.Replace "[[AB-Nummer]]", abNumber, xlWhole
        tempWs.UsedRange.Replace "[[Tour_DATE & Time]]", deliveryDate, xlWhole
        tempWs.UsedRange.Replace "[[ServiceTyp]]", serviceType, xlWhole
        tempWs.UsedRange.Replace "[[Kunde_Str 1]]", recipientAddress, xlWhole
        tempWs.UsedRange.Replace "[[Kunde_Ort 1]]", recipientCity, xlWhole
        tempWs.UsedRange.Replace "[[Kunde_Plz 1]]", recipientPostcode, xlWhole
        
        ' Replace NOS contact info placeholders
        tempWs.UsedRange.Replace "[[Nos_Ansprechpartner]]", nosContact, xlWhole
        tempWs.UsedRange.Replace "[[Nos_Tel]]", nosPhone, xlWhole
        tempWs.UsedRange.Replace "[[Nos_Mail]]", nosEmail, xlWhole
        
        ' Remove old placeholders for customer contact info
        tempWs.UsedRange.Replace "[[Kunde_Tel]]", "", xlWhole
        tempWs.UsedRange.Replace "[[Kunde_Mail]]", "", xlWhole
        tempWs.UsedRange.Replace "[[Kunde_Ansprechpartner]]", "", xlWhole
        
        tempWs.UsedRange.Replace "[[AB Info]]", abInfo, xlWhole
        tempWs.UsedRange.Replace "[[Gebäudeinfo]]", buildingInfo, xlWhole
        tempWs.UsedRange.Replace "[[Wichtiger Hinweis Auftrag]]", importantInfo, xlWhole
        tempWs.UsedRange.Replace "[[Anlieferinfo]]", deliveryInfo, xlWhole
        
        ' Find any remaining placeholder in the sheet and clear it
        tempWs.UsedRange.Replace "[[Stop_1]]", "", xlWhole
        
        ' Find the item table in the Frachtbrief template
        Dim itemTableRow As Long
        itemTableRow = 0
        
        ' Look for the item table header row
        For i = 1 To tempWs.UsedRange.Rows.count
            If InStr(1, tempWs.Cells(i, 1).Value, "Position") > 0 Then
                itemTableRow = i + 1 ' Start one row below the header
                Exit For
            End If
        Next i
        
        ' If not found, look for the placeholder
        If itemTableRow = 0 Then
            For i = 1 To tempWs.UsedRange.Rows.count
                If InStr(1, tempWs.Cells(i, 1).Value, "[[Stop_1]]") > 0 Then
                    itemTableRow = i
                    Exit For
                End If
            Next i
        End If
        
        ' Last resort fallback
        If itemTableRow = 0 Then
            itemTableRow = 25
        End If
        
        ' Process Artikel data - using the artikelTypen and warenText from the data
        If Len(warenText) > 0 Then
            ' Split and prepare artikelTypen if available
            Dim positionArr() As String
            Dim positionCount As Long
            positionCount = 0
            
            If Len(artikelTypen) > 0 Then
                If InStr(artikelTypen, ",") > 0 Then
                    positionArr = Split(artikelTypen, ",")
                    positionCount = UBound(positionArr) + 1
                Else
                    ReDim positionArr(0)
                    positionArr(0) = artikelTypen
                    positionCount = 1
                End If
            End If
            
            ' Split the items by the separator
            Dim items() As String
            items = Split(warenText, "----------")
            
            ' Current row for item insertion
            Dim row As Long
            row = itemTableRow
            
            ' Get column indices for the required fields
            Dim colPosition As Long, colQuantity As Long, colItemNumbers As Long, colItemName As Long, colIcon As Long
            
            ' Updated column positions (start at 1)
            colPosition = 1     ' B: Position
            colItemNumbers = 2  ' C-D are merged for item numbers
            colItemName = 5     ' E-F are merged for item name
            colQuantity = 7     ' A: Stückzahl Gesamt
            colIcon = 8         ' G is for Icon
            
            ' First, count how many items we'll have
            Dim itemCount As Long
            itemCount = 0
            
            For i = 0 To UBound(items)
                If Len(Trim(items(i))) > 0 Then
                    itemCount = itemCount + 1
                End If
            Next i
            
            ' Process each item
            For i = 0 To UBound(items)
                Dim itemText As String
                itemText = Trim(items(i))
                
                ' Clean the item text - remove any excessive whitespace, tabs, or line breaks
                itemText = CleanText(itemText)
                
                If Len(itemText) > 0 Then
                    ' Get position for this item if available
                    Dim position As String
                    position = ""
                    
                    If i < positionCount Then
                        position = Trim(positionArr(i))
                    End If
                    
                    ' Parse the item data - split by pipe character
                    If InStr(itemText, "|") > 0 Then
                        Dim parts() As String
                        parts = Split(itemText, "|")
                        
                        ' Clean each part
                        For k = 0 To UBound(parts)
                            parts(k) = Trim(parts(k))
                        Next k
                        
                        ' Format for the specific column layout
                        Dim itemNumbers As String
                        Dim itemName As String
                        
                        ' Default values
                        itemNumbers = ""
                        itemName = ""
                        
                        If UBound(parts) >= 3 Then
                            ' Get the first 3 item numbers
                            itemNumbers = parts(0) & " | " & parts(1) & " | " & parts(2)
                            
                            ' Special handling for 4th part (parts(3))
                            Dim part4 As String
                            part4 = parts(3)
                            
                            ' Look for product codes in the 4th part
                            If IsNumeric(Left(part4, 1)) Then
                                ' If it starts with a number, it's likely a product code
                                itemNumbers = itemNumbers & " | " & part4
                                
                                ' Check if there are additional parts for item name
                                If UBound(parts) >= 4 Then
                                    itemName = parts(4)
                                    ' Add any additional parts to item name
                                    For k = 5 To UBound(parts)
                                        itemName = itemName & " " & parts(k)
                                    Next k
                                End If
                            Else
                                ' If it doesn't start with a number, it's likely the start of the item name
                                itemName = part4
                                ' Add any additional parts to item name
                                For k = 4 To UBound(parts)
                                    itemName = itemName & " " & parts(k)
                                Next k
                            End If
                        Else
                            ' Not enough parts, just distribute what we have
                            If UBound(parts) >= 1 Then
                                ' At least 2 parts - use first part for numbers, second for name
                                itemNumbers = parts(0)
                                itemName = parts(1)
                                If UBound(parts) >= 2 Then
                                    itemName = itemName & " " & parts(2)
                                End If
                            Else
                                ' Only one part - use it as numbers
                                itemNumbers = parts(0)
                            End If
                        End If
                        
                        ' Remove any remaining pipe characters in the item name
                        itemName = Replace(itemName, "|", " ")
                    Else
                        ' Simple case - no pipe separator
                        itemNumbers = ""
                        itemName = itemText
                    End If
                    
                    ' Fill in the template cells according to your specific layout
                    tempWs.Cells(row, colPosition).Value = position
                    
                    ' Only put total quantity in the first item row
                    If i = 0 Then
                        tempWs.Cells(row, colQuantity).Value = totalQuantity
                    Else
                        tempWs.Cells(row, colQuantity).Value = ""
                    End If
                    
                    tempWs.Cells(row, colItemNumbers).Value = itemNumbers   ' Goes in column C (merged C-D)
                    tempWs.Cells(row, colItemName).Value = itemName         ' Goes in column E (merged E-F)
                    
                    ' Make sure to delete any existing merged cells before merging
                    On Error Resume Next
                    If tempWs.Cells(row, colItemNumbers).MergeCells Then
                        tempWs.Cells(row, colItemNumbers).MergeArea.UnMerge
                    End If
                    If tempWs.Cells(row, colItemName).MergeCells Then
                        tempWs.Cells(row, colItemName).MergeArea.UnMerge
                    End If
                    On Error GoTo ErrorHandler
                    
                    ' Merge the cells for item numbers (C-D)
                    tempWs.Range(tempWs.Cells(row, colItemNumbers), tempWs.Cells(row, colItemNumbers + 1)).Merge
                    
                    ' Merge the cells for item name (E-F)
                    tempWs.Range(tempWs.Cells(row, colItemName), tempWs.Cells(row, colItemName + 1)).Merge
                    
                    ' Set the alignment and formatting for item numbers (left-aligned, vertically centered)
                    With tempWs.Cells(row, colItemNumbers)
                        .HorizontalAlignment = xlLeft
                        .VerticalAlignment = xlCenter
                        .WrapText = False ' Ensure text doesn't wrap
                    End With
                    
                    ' Set the alignment and formatting for item name (left-aligned, vertically centered, bold)
                    With tempWs.Cells(row, colItemName)
                        .HorizontalAlignment = xlLeft
                        .VerticalAlignment = xlCenter
                        .Font.Bold = True
                        .WrapText = True ' Allow wrapping for longer item names
                    End With
                    
                    ' Insert a new row for the next item if there are more
                    If i < UBound(items) And i < itemCount - 1 Then
                        tempWs.Rows(row + 1).Insert Shift:=xlDown
                        
                        ' Copy formatting from the current row
                        tempWs.Rows(row).Copy
                        tempWs.Rows(row + 1).PasteSpecial xlPasteFormats
                        Application.CutCopyMode = False
                    End If
                    
                    row = row + 1
                End If
            Next i
            
            ' After all items are added, merge the Stückzahl column
            If itemCount > 1 Then
                Dim firstRow As Long, lastRow As Long
                firstRow = itemTableRow
                lastRow = firstRow + itemCount - 1
                
                ' Merge the Stückzahl cells vertically
                tempWs.Range(tempWs.Cells(firstRow, colQuantity), tempWs.Cells(lastRow, colQuantity)).Merge
                
                ' Center the quantity both horizontally and vertically
                With tempWs.Cells(firstRow, colQuantity)
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .Font.Bold = True
                End With
            End If
        Else
            ' No item data, just clear the placeholders in the first item row
            tempWs.Cells(itemTableRow, colPosition).Value = ""
            tempWs.Cells(itemTableRow, colQuantity).Value = ""
            tempWs.Cells(itemTableRow, colItemNumbers).Value = ""
            tempWs.Cells(itemTableRow, colItemName).Value = ""
        End If
    Next j
    
    ' Delete the first blank sheet that was created with the workbook
    Application.DisplayAlerts = False
    If tempWb.Sheets.count > 1 Then
        tempWb.Sheets(1).Delete
    End If
    Application.DisplayAlerts = True
    
    ' Set the print settings for all sheets
    For Each tempWs In tempWb.Worksheets
        With tempWs.PageSetup
            .Orientation = xlPortrait
            .PaperSize = xlPaperA4  ' Explicitly set to A4 paper size
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
            .CenterHorizontally = True
            .CenterVertically = False
        End With
    Next tempWs
    
    ' Create a truly unique filename using timestamp, tour number, and random ID
    pdfFileNameFracht = pdfPath & "Tour_" & tourNumber & "_Frachtbrief_" & uniqueID & ".pdf"
    
    ' Export as PDF with unique name
    On Error Resume Next
    tempWb.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfFileNameFracht, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    
    ' If PDF creation failed, try an alternative approach
    If Err.Number <> 0 Then
        Err.Clear
        
        ' Save the workbook as an intermediary file with unique name
        Dim tempXlsxPath As String
        tempXlsxPath = pdfPath & "Temp_" & tourNumber & "_" & uniqueID & ".xlsx"
        
        ' Save as Excel first
        tempWb.SaveAs Filename:=tempXlsxPath, FileFormat:=xlOpenXMLWorkbook
        
        ' Then try to create PDF from the saved file
        tempWb.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfFileNameFracht, Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
            
        ' Clean up the temporary Excel file
        On Error Resume Next
        tempWb.Close SaveChanges:=False
        
        ' Give Excel a moment to close the file
        Application.Wait Now + TimeValue("00:00:01")
        
        If Dir(tempXlsxPath) <> "" Then
            Kill tempXlsxPath
        End If
    Else
        ' Close normally if PDF export was successful
        tempWb.Close SaveChanges:=False
    End If
    
    ' Make sure the PDF file exists
    If Dir(pdfFileNameFracht) = "" Then
        MsgBox "Warning: PDF file was not created for tour " & tourNumber & ". Using backup method.", vbExclamation
        
        ' Create a very simple text summary as fallback
        Dim textFilePath As String
        textFilePath = pdfPath & "Tour_" & tourNumber & "_Frachtbrief_" & uniqueID & ".txt"
        
        Dim textFile As Object
        Set textFile = CreateObject("Scripting.FileSystemObject").CreateTextFile(textFilePath, True)
        
        textFile.WriteLine "FRACHTBRIEF (TEXT BACKUP)"
        textFile.WriteLine "========================"
        textFile.WriteLine ""
        textFile.WriteLine "Tour: " & tourNumber & " - " & tourName
        textFile.WriteLine "Stops: " & stopCount
        textFile.WriteLine "Created: " & Now()
        
        textFile.Close
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating combined PDF for tour " & tourNumber & ": " & Err.description, vbCritical
    On Error Resume Next
    
    ' Clean up any open workbooks
    If Not tempWb Is Nothing Then
        If tempWb.Name <> "" Then  ' Check if workbook is still open
            tempWb.Close SaveChanges:=False
        End If
    End If
End Sub

' Helper function to clean file name
Function CleanForFileName(textValue As String) As String
    ' Clean a string to make it valid for a file or folder name
    Dim invalidChars As String
    Dim result As String
    Dim i As Long, char As String
    
    ' Define characters that are invalid in file/folder names
    invalidChars = "\/:*?""<>|"
    
    result = textValue
    
    ' Replace invalid characters with underscores
    For i = 1 To Len(invalidChars)
        char = Mid(invalidChars, i, 1)
        result = Replace(result, char, "_")
    Next i
    
    ' Replace spaces with underscores for better folder names
    result = Replace(result, " ", "_")
    
    ' Handle periods - replace all except the last one (for file extension)
    Dim lastDotPos As Long
    lastDotPos = InStrRev(result, ".")
    
    ' If there's a period in the name
    If lastDotPos > 0 Then
        ' If it's at the end of the string, remove it
        If lastDotPos = Len(result) Then
            result = Left(result, lastDotPos - 1)
        ElseIf lastDotPos = Len(result) - 1 Then
            ' If it's the second-last character, remove it too
            result = Left(result, lastDotPos - 1)
        Else
            ' Replace all periods except those that appear to be part of a file extension
            Dim beforeLastDot As String, afterLastDot As String
            beforeLastDot = Left(result, lastDotPos - 1)
            afterLastDot = Mid(result, lastDotPos)
            
            ' Replace periods in the first part
            beforeLastDot = Replace(beforeLastDot, ".", "_")
            
            ' Reassemble the string
            result = beforeLastDot & afterLastDot
        End If
    End If
    
    ' Remove any multiple underscores
    Do While InStr(result, "__") > 0
        result = Replace(result, "__", "_")
    Loop
    
    ' Trim leading and trailing underscores
    If Left(result, 1) = "_" Then result = Mid(result, 2)
    If Right(result, 1) = "_" Then result = Left(result, Len(result) - 1)
    
    ' Limit length to avoid path too long errors (adjust as needed)
    If Len(result) > 50 Then
        result = Left(result, 50)
    End If
    
    CleanForFileName = result
End Function

' Helper function to check if a cell is part of a merged range
Function IsMerged(cell As Range) As Boolean
    On Error Resume Next
    IsMerged = cell.MergeCells
    On Error GoTo 0
End Function

' Helper function to clean text and remove unwanted characters
Function CleanText(text As String) As String
    Dim cleanedText As String
    cleanedText = text
    
    ' Remove line breaks and tabs
    cleanedText = Replace(cleanedText, vbCr, " ")
    cleanedText = Replace(cleanedText, vbLf, " ")
    cleanedText = Replace(cleanedText, vbCrLf, " ")
    cleanedText = Replace(cleanedText, vbTab, " ")
    
    ' Replace multiple spaces with a single space
    Do While InStr(cleanedText, "  ") > 0
        cleanedText = Replace(cleanedText, "  ", " ")
    Loop
    
    ' Trim leading and trailing spaces
    cleanedText = Trim(cleanedText)
    
    CleanText = cleanedText
End Function


Function FormatItemsList(itemsText As String, artikelTypen As String) As String
    ' Format the items list combining Packstück Artikeltypen and Warenbeschreibung
    On Error Resume Next
    
    ' Default for empty input
    If Len(itemsText) = 0 Then
        FormatItemsList = "No items"
        Exit Function
    End If
    
    Dim formattedList As String
    formattedList = ""
    
    ' Split the items by the "----------" separator
    Dim items() As String
    items = Split(itemsText, "----------")
    
    ' Parse Artikeltypen into an array for positions
    Dim positions() As String
    Dim positionsCount As Long
    
    positionsCount = 0
    
    If Len(artikelTypen) > 0 Then
        ' Try different separators
        If InStr(artikelTypen, ",") > 0 Then
            positions = Split(artikelTypen, ",")
            positionsCount = UBound(positions) + 1
        Else
            ' If no commas, try to split by numbers
            positions = Split(artikelTypen, ".")
            If UBound(positions) > 0 Then
                positionsCount = UBound(positions) + 1
            Else
                ' If still no luck, assume each number is an item position
                positions = Array(artikelTypen)
                positionsCount = 1
            End If
        End If
    End If
    
    Dim i As Long
    For i = 0 To UBound(items)
        If Len(Trim(items(i))) > 0 Then
            ' Clean up the item text
            Dim itemLine As String
            itemLine = Trim(Replace(Replace(Replace(items(i), vbCrLf, ""), vbCr, ""), vbLf, ""))
            
            ' Add formatted item to the list
            If Len(formattedList) > 0 Then formattedList = formattedList & vbCrLf
            
            ' Split by pipe character to get the item parts
            Dim parts() As String
            If InStr(itemLine, "|") > 0 Then
                parts = Split(itemLine, "|")
                
                If UBound(parts) >= 3 Then
                    ' Get position for this item if available
                    Dim position As String
                    position = ""
                    
                    If i < positionsCount Then
                        position = Trim(positions(i))
                    End If
                    
                    ' Format all 4 item numbers and description
                    Dim itemNum1 As String, itemNum2 As String, itemNum3 As String, itemNum4 As String
                    Dim itemDesc As String
                    
                    itemNum1 = Trim(parts(0))
                    itemNum2 = Trim(parts(1))
                    itemNum3 = Trim(parts(2))
                    itemNum4 = Trim(parts(3))
                    
                    ' Extract item description from part 4
                    itemDesc = ""
                    
                    ' Find where the description starts (after any spaces in part 4)
                    Dim startPos As Long
                    startPos = 1
                    While startPos <= Len(itemNum4) And Mid(itemNum4, startPos, 1) = " "
                        startPos = startPos + 1
                    Wend
                    
                    ' Separate item code from description
                    Dim itemCode As String
                    Dim descStartPos As Long
                    
                    ' Find first space sequence after initial non-spaces
                    descStartPos = startPos
                    Dim inSpace As Boolean
                    inSpace = False
                    
                    For j = startPos To Len(itemNum4)
                        If Mid(itemNum4, j, 1) = " " Then
                            If Not inSpace Then
                                inSpace = True
                            End If
                        ElseIf inSpace Then
                            descStartPos = j
                            Exit For
                        End If
                    Next j
                    
                    If descStartPos > startPos And descStartPos < Len(itemNum4) Then
                        itemCode = Trim(Left(itemNum4, descStartPos - 1))
                        itemDesc = Trim(Mid(itemNum4, descStartPos))
                    Else
                        itemCode = itemNum4
                        itemDesc = ""
                    End If
                    
                    ' Format as "• Position | Item1 | Item2 | Item3 | ItemCode - ItemDesc"
                    If Len(position) > 0 Then
                        formattedList = formattedList & "• " & position & " | " & itemNum1 & " | " & itemNum2 & " | " & itemNum3 & " | " & itemCode & " - " & itemDesc
                    Else
                        formattedList = formattedList & "• " & itemNum1 & " | " & itemNum2 & " | " & itemNum3 & " | " & itemCode & " - " & itemDesc
                    End If
                Else
                    ' Fallback if parts are not as expected
                    formattedList = formattedList & "• " & itemLine
                End If
            Else
                ' No pipe separator, use the whole line
                formattedList = formattedList & "• " & itemLine
            End If
        End If
    Next i
    
    FormatItemsList = formattedList
End Function

Sub ParseTourNameAndDate(tourNameText As String, ByRef tourName As String, ByRef tourDate As String)
    ' Parse tour name and date from format like "Wien 8 - 07.04." or "SC Wr. Neudorf 07.04."
    On Error Resume Next
    
    Dim pos As Long
    
    ' Check if it's a Service Center tour
    If Left(tourNameText, 3) = "SC " Then
        ' Remove "SC " prefix
        tourNameText = Mid(tourNameText, 4)
        
        ' Find the date which is typically the last part
        pos = InStrRev(tourNameText, " ")
        If pos > 0 Then
            tourName = Left(tourNameText, pos - 1)
            tourDate = Mid(tourNameText, pos + 1)
        Else
            tourName = tourNameText
            tourDate = ""
        End If
    Else
        ' Regular tour, look for the hyphen
        pos = InStr(tourNameText, " - ")
        If pos > 0 Then
            tourName = Left(tourNameText, pos - 1)
            tourDate = Mid(tourNameText, pos + 3)
        Else
            ' Alternative format without hyphen
            pos = InStrRev(tourNameText, " ")
            If pos > 0 Then
                tourName = Left(tourNameText, pos - 1)
                tourDate = Mid(tourNameText, pos + 1)
            Else
                tourName = tourNameText
                tourDate = ""
            End If
        End If
    End If
End Sub

Function BrowseForFolder(Optional prompt As String = "Select a folder") As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = prompt
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Function
        BrowseForFolder = .SelectedItems(1)
    End With
End Function

Function GetItemCategory(itemDescription As String) As String
    ' Extract category from item description
    ' This is just a placeholder - you may need to adjust based on your actual data
    
    On Error Resume Next
    
    Dim category As String
    category = ""
    
    ' Look for common category indicators - modify this based on your actual categories
    If InStr(1, itemDescription, "Tisch", vbTextCompare) > 0 Then
        category = "Tisch"
    ElseIf InStr(1, itemDescription, "Schrank", vbTextCompare) > 0 Then
        category = "Schrank"
    ElseIf InStr(1, itemDescription, "Stuhl", vbTextCompare) > 0 Then
        category = "Stuhl"
    ElseIf InStr(1, itemDescription, "Container", vbTextCompare) > 0 Then
        category = "Container"
    ElseIf InStr(1, itemDescription, "Regal", vbTextCompare) > 0 Then
        category = "Regal"
    End If
    
    GetItemCategory = category
End Function

Function CountItems(warenText As String) As Long
    If Len(Trim(warenText)) = 0 Then
        CountItems = 0
        Exit Function
    End If
    
    Dim items() As String
    items = Split(warenText, "----------")
    
    ' Count non-empty items
    Dim count As Long, i As Long
    count = 0
    
    For i = 0 To UBound(items)
        If Len(Trim(items(i))) > 0 Then
            count = count + 1
        End If
    Next i
    
    CountItems = count
End Function

Function FindFormulaRow(ws As Worksheet, tourNumber As String) As Long
    Dim lastRow As Long, i As Long
    
    FindFormulaRow = 0 ' Default return value if not found
    
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    
    For i = 2 To lastRow
        ' Look for rows with tour number but no stop number (likely a formula/total row)
        If ws.Cells(i, 1).Value = tourNumber And Not IsNumeric(ws.Cells(i, 3).Value) Then
            FindFormulaRow = i
            Exit Function
        End If
    Next i
End Function

Function ExtractNumericValue(cellValue As Variant) As Double
    On Error Resume Next
    
    ExtractNumericValue = 0 ' Default
    
    If IsNumeric(cellValue) Then
        ExtractNumericValue = CDbl(cellValue)
    Else
        ' Try to extract number from text
        Dim textValue As String, i As Long
        Dim numStr As String, foundDigit As Boolean
        
        textValue = CStr(cellValue)
        numStr = ""
        foundDigit = False
        
        For i = 1 To Len(textValue)
            Dim ch As String
            ch = Mid(textValue, i, 1)
            
            ' Accept digits, decimal point, and minus sign
            If (ch >= "0" And ch <= "9") Or ch = "." Or (ch = "-" And numStr = "") Then
                numStr = numStr & ch
                If ch >= "0" And ch <= "9" Then foundDigit = True
            ElseIf foundDigit And (ch = " " Or ch = vbTab) Then
                ' Space after digits can end our number
                Exit For
            End If
        Next i
        
        If foundDigit And IsNumeric(numStr) Then
            ExtractNumericValue = CDbl(numStr)
        End If
    End If
    
    Err.Clear
End Function

Function ParseItems(warenText As String) As Variant
    Dim items() As String
    Dim result() As Variant
    Dim i As Long, count As Long
    
    ' Default empty result
    ReDim result(0, 0)
    result(0, 0) = ""
    
    If Len(Trim(warenText)) = 0 Then
        ParseItems = result
        Exit Function
    End If
    
    ' Split by separator
    items = Split(warenText, "----------")
    
    ' Count non-empty items
    count = 0
    For i = 0 To UBound(items)
        If Len(Trim(items(i))) > 0 Then
            count = count + 1
        End If
    Next i
    
    If count = 0 Then
        ParseItems = result
        Exit Function
    End If
    
    ' Resize result array
    ReDim result(count - 1, 4) ' 5 columns: itemNum1, itemNum2, itemNum3, itemNum4, category
    
    ' Process each item
    Dim itemIndex As Long
    itemIndex = 0
    
    For i = 0 To UBound(items)
        Dim itemText As String
        itemText = Trim(items(i))
        
        If Len(itemText) > 0 Then
            ' Default values
            result(itemIndex, 0) = ""
            result(itemIndex, 1) = ""
            result(itemIndex, 2) = ""
            result(itemIndex, 3) = ""
            result(itemIndex, 4) = "" ' Category
            
            ' Split by pipe
            If InStr(itemText, "|") > 0 Then
                Dim parts() As String
                parts = Split(itemText, "|")
                
                ' Assign parts to result
                Dim j As Long
                For j = 0 To UBound(parts)
                    If j <= 3 Then
                        result(itemIndex, j) = Trim(parts(j))
                    End If
                Next j
            Else
                ' No pipe separators, just use as is
                result(itemIndex, 3) = itemText
            End If
            
            itemIndex = itemIndex + 1
        End If
    Next i
    
    ParseItems = result
End Function

Sub SpecialHandlingForTour(ws As Worksheet, pdfPath As String, tourNumber As String)
    ' This function creates a simple text summary for any tour that fails PDF generation
    ' Use it as a fallback when PDF generation fails
    
    Dim lastRow As Long, i As Long
    Dim tourStops As New Collection
    Dim totalWeight As Double, totalVolume As Double
    
    ' Find all rows for this tour and collect info
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    
    ' First collect all stops and their information
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = tourNumber And IsNumeric(ws.Cells(i, 3).Value) Then
            tourStops.Add i ' Store the row number for each stop
            
            ' Add weight and volume for totals
            If IsNumeric(ws.Cells(i, 4).Value) Then
                totalWeight = totalWeight + CDbl(ws.Cells(i, 4).Value)
            End If
            
            If IsNumeric(ws.Cells(i, 5).Value) Then
                totalVolume = totalVolume + CDbl(ws.Cells(i, 5).Value)
            End If
        End If
    Next i
    
    ' Create a simple text file with tour information
    Dim fso As Object, textFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Create a fallback folder
    Dim fallbackFolder As String
    fallbackFolder = pdfPath & tourNumber & "_Backup"
    
    On Error Resume Next
    If Not fso.FolderExists(fallbackFolder) Then
        fso.CreateFolder fallbackFolder
    End If
    
    If Err.Number <> 0 Then
        Err.Clear
        fallbackFolder = pdfPath
    End If
    
    ' Ensure folder path ends with backslash
    If Right(fallbackFolder, 1) <> "\" Then
        fallbackFolder = fallbackFolder & "\"
    End If
    
    ' Get original tour name
    Dim tourFullName As String
    If tourStops.count > 0 Then
        Dim firstRow As Long
        firstRow = tourStops(1)
        If Not IsEmpty(ws.Cells(firstRow, 2).Value) Then
            tourFullName = ws.Cells(firstRow, 2).Value
        Else
            tourFullName = "Tour " & tourNumber
        End If
    Else
        tourFullName = "Tour " & tourNumber
    End If
    
    ' Create a simple text summary
    Set textFile = fso.CreateTextFile(fallbackFolder & "Tour_" & tourNumber & "_Summary.txt", True)
    
    textFile.WriteLine "TOUR SUMMARY"
    textFile.WriteLine "============"
    textFile.WriteLine ""
    textFile.WriteLine "Tour Number: " & tourNumber
    textFile.WriteLine "Tour Name: " & tourFullName
    textFile.WriteLine "Total Weight: " & Format(totalWeight, "#,##0.00") & " kg"
    textFile.WriteLine "Total Volume: " & Format(totalVolume, "#,##0.00") & " m³"
    textFile.WriteLine "Number of Stops: " & tourStops.count
    textFile.WriteLine "Generated: " & Now()
    textFile.WriteLine ""
    textFile.WriteLine "STOPS"
    textFile.WriteLine "====="
    
    ' Add individual stops
    For i = 1 To tourStops.count
        Dim currentRow As Long
        currentRow = tourStops(i)
        
        Dim stopNum As Long, recipientName As String, weight As Double, volume As Double
        Dim address As String, postcode As String, city As String, fullAddress As String
        
        stopNum = ws.Cells(currentRow, 3).Value ' Column C
        
        recipientName = ""
        If Not IsEmpty(ws.Cells(currentRow, 36).Value) Then ' Column AI - Empfänger
            recipientName = ws.Cells(currentRow, 36).Value
        End If
        
        ' Get address information
        address = ""
        If Not IsEmpty(ws.Cells(currentRow, 37).Value) Then ' Column AJ - Empf. Str.
            address = ws.Cells(currentRow, 37).Value
        End If
        
        city = ""
        If Not IsEmpty(ws.Cells(currentRow, 38).Value) Then ' Column AK - Empf. Ort
            city = ws.Cells(currentRow, 38).Value
        End If
        
        postcode = ""
        If Not IsEmpty(ws.Cells(currentRow, 39).Value) Then ' Column AL - Empf. Plz
            postcode = ws.Cells(currentRow, 39).Value
        End If
        
        ' Format the address line
        fullAddress = address
        If Len(postcode) > 0 Or Len(city) > 0 Then
            If Len(fullAddress) > 0 Then
                fullAddress = fullAddress & ", "
            End If
            fullAddress = fullAddress & postcode & " " & city
        End If
        
        If IsNumeric(ws.Cells(currentRow, 4).Value) Then ' Column D
            weight = ws.Cells(currentRow, 4).Value
        End If
        
        If IsNumeric(ws.Cells(currentRow, 5).Value) Then ' Column E
            volume = ws.Cells(currentRow, 5).Value
        End If
        
        textFile.WriteLine ""
        textFile.WriteLine "Stop " & stopNum
        textFile.WriteLine "Customer: " & recipientName
        textFile.WriteLine "Address: " & fullAddress
        textFile.WriteLine "Weight: " & Format(weight, "#,##0.00") & " kg"
        textFile.WriteLine "Volume: " & Format(volume, "#,##0.00") & " m³"
    Next i
    
    textFile.WriteLine ""
    textFile.WriteLine "NOTE: This text summary was generated because PDF creation encountered an error."
    
    textFile.Close
    
    MsgBox "A text summary has been created for tour " & tourNumber & " at: " & fallbackFolder, vbInformation
End Sub
