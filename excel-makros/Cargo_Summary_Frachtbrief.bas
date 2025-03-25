Sub ProcessToursAndCreatePDFs()
    ' Tour Processing Macro - Complete version with PDF generation
    ' Version 2.0 - Combines both Tour Summary and PDF functionality
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim wsSummary As Worksheet
    Dim lastRow As Long, i As Long
    Dim tourNumber As String, lastTourNumber As String
    Dim tourName As String, tourDate As String
    Dim isServiceCenter As Boolean
    Dim pdfFolderPath As String
    
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
                    Dim newTourData(5) As Variant
                    newTourData(0) = "" ' Tour name (filled below)
                    newTourData(1) = "" ' Tour date (filled below)
                    newTourData(2) = "" ' Tour type (filled below)
                    newTourData(3) = 0 ' Total weight - we'll calculate this
                    newTourData(4) = 0 ' Total volume - we'll calculate this
                    newTourData(5) = "" ' AB Numbers (filled below)
                    
                    tourSummary.Add tourNumber, newTourData
                End If
                
                ' Get tour data array
                Dim tourDataArray As Variant
                tourDataArray = tourSummary(tourNumber)
                
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
    
    ' Now create PDF for each tour
    For Each tourKey In tourSummary.Keys
        ' Get the tour data
        currentTourData = tourSummary(tourKey)
        
        ' Convert both tourKey and tourName to string to avoid ByRef error
        Dim tourKeyStr As String, tourNameStr As String
        tourKeyStr = CStr(tourKey)
        tourNameStr = CStr(currentTourData(0))
        
        Application.StatusBar = "Creating PDFs for Tour " & tourKeyStr & "..."
        CreateTourPDF ws, tourKeyStr, tourNameStr, pdfFolderPath
    Next tourKey
    
CleanExit:
    ' Reset status bar and screen updating
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "Tour summary created in the TourSummary worksheet. PDF files saved to: " & pdfFolderPath, vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description & " (Line: " & Erl & ")", vbCritical
    Resume CleanExit
End Sub

Sub CreateTourPDF(ws As Worksheet, tourNumber As String, tourName As String, pdfPath As String)
    ' Create a PDF file for the tour with stop information and freight details
    Dim lastRow As Long, i As Long
    Dim stopCount As Long
    Dim totalWeight As Double, totalVolume As Double
    
    ' Find all rows for this tour and count how many stops
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    stopCount = 0
    
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = tourNumber And IsNumeric(ws.Cells(i, 3).Value) Then
            stopCount = stopCount + 1
            
            ' Add weight and volume for totals
            If IsNumeric(ws.Cells(i, 4).Value) Then
                totalWeight = totalWeight + CDbl(ws.Cells(i, 4).Value)
            End If
            
            If IsNumeric(ws.Cells(i, 5).Value) Then
                totalVolume = totalVolume + CDbl(ws.Cells(i, 5).Value)
            End If
        End If
    Next i
    
    ' First create a summary PDF with all stops for this tour
    CreateTourSummaryPDF ws, tourNumber, tourName, pdfPath, totalWeight, totalVolume
    
    ' Process each stop for this tour
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = tourNumber And IsNumeric(ws.Cells(i, 3).Value) Then
            ' Create a PDF for this stop
            CreateStopFreightPDF ws, i, tourNumber, tourName, pdfPath
        End If
    Next i
End Sub

Sub CreateStopFreightPDF(ws As Worksheet, rowNum As Long, tourNumber As String, tourName As String, pdfPath As String)
    ' Create a PDF file for a specific stop on a tour
    Dim tempWs As Worksheet
    Dim stopNum As Long
    Dim abNumber As String
    Dim artikelTypen As String, warenText As String
    Dim pdfFileName As String
    Dim i As Long
    Dim recipientName As String, recipientAddress As String, recipientCity As String, recipientPostcode As String
    Dim deliveryDate As String, deliveryTimeWindow As String
    Dim buildingInfo As String, deliveryRemarks As String
    Dim stopWeight As Double, stopVolume As Double
    
    ' Get stop information
    stopNum = ws.Cells(rowNum, 3).Value ' Column C
    abNumber = ws.Cells(rowNum, 12).Value ' Column L
    artikelTypen = ws.Cells(rowNum, 47).Value ' Column AU - Packstück Artikeltypen
    warenText = ws.Cells(rowNum, 48).Value ' Column AV - Warenbeschreibung
    stopWeight = ws.Cells(rowNum, 4).Value ' Column D - Weight
    stopVolume = ws.Cells(rowNum, 5).Value ' Column E - Volume
    
    ' Get recipient info
    recipientName = ws.Cells(rowNum, 35).Value ' Column AI - Empfänger
    recipientAddress = ws.Cells(rowNum, 36).Value ' Column AJ - Empf. Str.
    recipientCity = ws.Cells(rowNum, 37).Value ' Column AK - Empf. Ort
    recipientPostcode = ws.Cells(rowNum, 38).Value ' Column AL - Empf. Plz
    
    ' Get delivery info
    If Not IsEmpty(ws.Cells(rowNum, 16).Value) Then ' Column P - Liefertag_System
        deliveryDate = Format(ws.Cells(rowNum, 16).Value, "dd.MM.yyyy")
    ElseIf Not IsEmpty(ws.Cells(rowNum, 17).Value) Then ' Column Q - Leistungsdatum
        deliveryDate = Format(ws.Cells(rowNum, 17).Value, "dd.MM.yyyy")
    Else
        deliveryDate = ""
    End If
    
    ' Get time window if present (often in column for arrival time)
    deliveryTimeWindow = ""
    If Not IsEmpty(ws.Cells(rowNum, 31).Value) Then ' Column AE - Entladestart
        deliveryTimeWindow = ws.Cells(rowNum, 31).Value
    End If
    
    buildingInfo = ws.Cells(rowNum, 42).Value ' Column AP - Gebäudeinfo
    deliveryRemarks = ws.Cells(rowNum, 44).Value ' Column AR - Anlieferinfo
    
    ' Create a temporary worksheet for the PDF
    Set tempWs = Worksheets.Add
    tempWs.Name = "TempPDF_Stop" & stopNum
    
    ' Create the header for the PDF
    With tempWs
        ' Set larger paper size
        .PageSetup.PaperSize = xlPaperA4
        .PageSetup.Orientation = xlPortrait
        
        ' Company logo position (top right)
        .Range("I1:K3").Merge
        .Range("I1").Value = "Erik"
        .Range("I1").Font.Size = 36
        .Range("I1").Font.Bold = True
        .Range("I1").HorizontalAlignment = xlRight
        
        ' AB Number and customer reference (top left)
        .Range("A1:E1").Merge
        .Range("A1").Value = "AB-Nummer: " & abNumber
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        
        ' Tour information
        .Range("A3").Value = "Tour: " & tourNumber
        .Range("A4").Value = "Tourname: " & tourName
        .Range("A5").Value = "Stop: " & stopNum
        .Range("A3:A5").Font.Bold = True
        
        ' Weight and volume
        .Range("I4").Value = "Weight (kg):"
        .Range("I5").Value = "Volume (m³):"
        .Range("J4").Value = Format(stopWeight, "#,##0.00")
        .Range("J5").Value = Format(stopVolume, "#,##0.00")
        .Range("I4:I5").Font.Bold = True
        .Range("J4:J5").HorizontalAlignment = xlRight
        
        ' Recipient information
        .Range("A7").Value = "Recipient:"
        .Range("B7").Value = recipientName
        .Range("A8").Value = "Address:"
        .Range("B8").Value = recipientAddress
        .Range("A9").Value = "City:"
        .Range("B9").Value = recipientPostcode & " " & recipientCity
        .Range("A7:A9").Font.Bold = True
        
        ' Delivery information
        .Range("E7").Value = "Delivery Date:"
        .Range("F7").Value = deliveryDate
        If Len(deliveryTimeWindow) > 0 Then
            .Range("F7").Value = deliveryDate & " " & deliveryTimeWindow
        End If
        
        .Range("E8").Value = "Building Info:"
        .Range("F8").Value = buildingInfo
        .Range("E9").Value = "Delivery Remarks:"
        .Range("F9").Value = deliveryRemarks
        .Range("E7:E9").Font.Bold = True
        
        ' Draw horizontal line
        .Range("A11:K11").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("A11:K11").Borders(xlEdgeBottom).weight = xlMedium
        
        ' Items header
        .Range("A13").Value = "Item Number"
        .Range("C13").Value = "Item Name / Description"
        .Range("G13").Value = "Category"
        .Range("A13:G13").Font.Bold = True
        .Range("A13:G13").Interior.Color = RGB(220, 220, 220)
        
        ' Parse and format the items
        Dim row As Long
        row = 14
        
        If Len(warenText) > 0 Then
            ' Parse items from Warenbeschreibung
            Dim items() As String
            items = Split(warenText, "----------")
            
            ' Parse artikelTypen into a dictionary for easier mapping
            Dim artikelDict As Object
            Set artikelDict = CreateObject("Scripting.Dictionary")
            
            If Len(artikelTypen) > 0 Then
                Dim artikelItems() As String
                
                ' Try split by commas first
                If InStr(artikelTypen, ",") > 0 Then
                    artikelItems = Split(artikelTypen, ",")
                Else
                    ' Fallback to straight assignment
                    ReDim artikelItems(0)
                    artikelItems(0) = artikelTypen
                End If
                
                ' Add to dictionary with position as key
                For i = 0 To UBound(artikelItems)
                    artikelDict.Add i, Trim(artikelItems(i))
                Next i
            End If
            
            ' Process each item
            For i = 0 To UBound(items)
                Dim itemText As String
                itemText = Trim(items(i))
                
                If Len(itemText) > 0 Then
                    ' Split by pipe
                    Dim parts() As String
                    If InStr(itemText, "|") > 0 Then
                        parts = Split(itemText, "|")
                        
                        If UBound(parts) >= 3 Then
                            ' Get item details
                            Dim itemNumber As String, itemDescription As String, itemCategory As String
                            
                            itemNumber = Trim(parts(0))
                            
                            ' Find the actual description part (usually at end)
                            itemDescription = Trim(parts(UBound(parts)))
                            
                            ' Get category if available
                            itemCategory = ""
                            If artikelDict.Exists(i) Then
                                itemCategory = artikelDict(i)
                            End If
                            
                            ' Add to worksheet
                            .Cells(row, 1).Value = itemNumber ' Item Number
                            .Cells(row, 3).Value = itemDescription ' Description
                            .Cells(row, 7).Value = itemCategory ' Category
                            
                            row = row + 1
                        End If
                    Else
                        ' Just add as is
                        .Cells(row, 3).Value = itemText
                        row = row + 1
                    End If
                End If
            Next i
        Else
            .Cells(row, 3).Value = "No item data available"
            row = row + 1
        End If
        
        ' Format the PDF worksheet
        .Columns("A").ColumnWidth = 12 ' Item number
        .Columns("B").ColumnWidth = 2 ' Spacer
        .Columns("C:F").ColumnWidth = 20 ' Item name - wider for longer descriptions
        .Columns("G").ColumnWidth = 10 ' Category
        
        ' Add page footer
        .Range("A50").Value = "Generated on: " & Format(Now, "dd.MM.yyyy HH:mm")
        .Range("I50").Value = "Page 1"
        
        ' Add borders to the item list
        If row > 14 Then
            .Range(.Cells(13, 1), .Cells(row - 1, 7)).Borders.LineStyle = xlContinuous
            .Range(.Cells(13, 1), .Cells(row - 1, 7)).Borders.weight = xlThin
        End If
    End With
    
    ' Save as PDF
    pdfFileName = pdfPath & "Tour_" & tourNumber & "_Stop_" & stopNum & "_Frachtbrief.pdf"
    
    ' Use larger paper size and better scaling
    tempWs.PageSetup.Orientation = xlPortrait
    tempWs.PageSetup.Zoom = False  ' Disable zoom
    tempWs.PageSetup.FitToPagesWide = 1
    tempWs.PageSetup.FitToPagesTall = 1
    
    tempWs.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfFileName, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    
    ' Delete the temporary worksheet
    Application.DisplayAlerts = False
    tempWs.Delete
    Application.DisplayAlerts = True
End Sub

Sub CreateTourSummaryPDF(ws As Worksheet, tourNumber As String, tourName As String, pdfPath As String, totalWeight As Double, totalVolume As Double)
    ' Create a summary PDF for all stops in a tour
    Dim tempWs As Worksheet
    Dim lastRow As Long, i As Long
    Dim pdfFileName As String
    Dim stopRow As Long
    Dim vehicleInfo As String, deliveryDate As String
    
    ' Get vehicle and date info from first row of this tour
    For i = 2 To ws.Cells(ws.Rows.count, "A").End(xlUp).row
        If ws.Cells(i, 1).Value = tourNumber And IsNumeric(ws.Cells(i, 3).Value) Then
            ' Get vehicle info
            If Not IsEmpty(ws.Cells(i, 25).Value) Then ' Column Y - Kennzeichen
                vehicleInfo = ws.Cells(i, 25).Value
            End If
            
            ' Get delivery date
            If Not IsEmpty(ws.Cells(i, 16).Value) Then ' Column P - Liefertag_System
                deliveryDate = Format(ws.Cells(i, 16).Value, "dd.MM.yyyy")
            ElseIf Not IsEmpty(ws.Cells(i, 17).Value) Then ' Column Q - Leistungsdatum
                deliveryDate = Format(ws.Cells(i, 17).Value, "dd.MM.yyyy")
            End If
            
            Exit For
        End If
    Next i
    
    ' Create a temporary worksheet for the PDF
    Set tempWs = Worksheets.Add
    tempWs.Name = "TempPDF_TourSummary"
    
    ' Create the header for the PDF
    With tempWs
        ' Title Section - "Tourenübersicht"
        .Range("A1:D2").Merge
        .Range("A1").Value = "Tourenübersicht"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 20
        
        ' Logo section (right side)
        .Range("F1:I2").Merge
        .Range("F1").Value = "bgo"
        .Range("F1").Font.Size = 24
        .Range("F1").Font.Bold = True
        .Range("F1").HorizontalAlignment = xlRight
        
        ' Tour detail section
        .Range("A4").Value = "Tour: " & tourNumber & ", " & tourName
        .Range("A4").Font.Bold = True
        .Range("A4").Font.Size = 12
        
        .Range("A5").Value = "Auslieferdatum: " & deliveryDate
        .Range("A6").Value = "Fahrzeug: " & vehicleInfo
        
        ' Create the table headers
        .Range("A8").Value = "Stopp"
        .Range("B8").Value = "Empfänger"
        .Range("C8").Value = "Ort"
        .Range("D8").Value = "cbm"
        .Range("E8").Value = "kg"
        .Range("F8").Value = "Mont./h"
        .Range("G8").Value = "Ankunft"
        .Range("A8:G8").Font.Bold = True
        
        ' Fill in the stops information
        stopRow = 9
        lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
        Dim totalMontagezeit As Double
        totalMontagezeit = 0
        
        For i = 2 To lastRow
            If ws.Cells(i, 1).Value = tourNumber And IsNumeric(ws.Cells(i, 3).Value) Then
                ' Get stop-specific values
                Dim stopNum As Long, recipientName As String, city As String
                Dim weight As Double, volume As Double, montagezeit As Double, arrivalTime As String
                
                stopNum = ws.Cells(i, 3).Value ' Column C
                recipientName = ws.Cells(i, 35).Value ' Column AI - Empfänger
                city = ws.Cells(i, 37).Value ' Column AK - Empf. Ort
                
                weight = 0
                If IsNumeric(ws.Cells(i, 4).Value) Then
                    weight = ws.Cells(i, 4).Value ' Column D
                End If
                
                volume = 0
                If IsNumeric(ws.Cells(i, 5).Value) Then
                    volume = ws.Cells(i, 5).Value ' Column E
                End If
                
                montagezeit = 0
                ' Try to find Montagezeit column (may be in AY, AZ, etc.)
                If Not IsEmpty(ws.Cells(i, 53).Value) And IsNumeric(ws.Cells(i, 53).Value) Then ' Column BA
                    montagezeit = ws.Cells(i, 53).Value
                End If
                
                arrivalTime = ""
                If Not IsEmpty(ws.Cells(i, 31).Value) Then ' Column AE - Entladestart
                    arrivalTime = ws.Cells(i, 31).Value
                End If
                
                ' Add to total montagezeit
                totalMontagezeit = totalMontagezeit + montagezeit
                
                ' Fill the row
                .Cells(stopRow, 1).Value = stopNum
                .Cells(stopRow, 2).Value = recipientName
                .Cells(stopRow, 3).Value = city
                .Cells(stopRow, 4).Value = volume
                .Cells(stopRow, 5).Value = weight
                .Cells(stopRow, 6).Value = montagezeit
                .Cells(stopRow, 7).Value = arrivalTime
                
                stopRow = stopRow + 1
            End If
        Next i
        
        ' Add row for totals
        .Cells(stopRow, 1).Value = ""
        .Cells(stopRow, 3).Value = ""
        .Cells(stopRow, 4).Value = totalVolume
        .Cells(stopRow, 5).Value = totalWeight
        .Cells(stopRow, 6).Value = totalMontagezeit
        
        ' Format numbers
        .Range("D9:F" & stopRow).NumberFormat = "#,##0.00"
        
        ' Add additional notes or delivery information if needed
        Dim maxRow As Long
        maxRow = stopRow + 3
        
        ' Add special instruction/note row
        .Cells(maxRow, 1).Value = "beide Fonds-Aufträge bis spätestens 15.00 Uhr ausliefern"
        .Range(.Cells(maxRow, 1), .Cells(maxRow, 7)).Merge
        
        ' Format the worksheet
        .Columns.AutoFit
        .Range("A8:G" & stopRow).Borders.LineStyle = xlContinuous
        .Range("A8:G" & stopRow).Borders.weight = xlThin
        
        ' Page number
        .Cells(maxRow + 2, 7).Value = "Seite 1 von 1"
    End With
    
    ' Save as PDF
    pdfFileName = pdfPath & "Tour_" & tourNumber & "_Summary.pdf"
    
    tempWs.PageSetup.Orientation = xlLandscape
    tempWs.PageSetup.FitToPagesWide = 1
    tempWs.PageSetup.FitToPagesTall = 1
    
    tempWs.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfFileName, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    
    ' Delete the temporary worksheet
    Application.DisplayAlerts = False
    tempWs.Delete
    Application.DisplayAlerts = True
End Sub

' Helper function to format the items list
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

' Helper function to parse tour name and date
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

' Helper function to browse for a folder
Function BrowseForFolder(Optional prompt As String = "Select a folder") As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = prompt
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Function
        BrowseForFolder = .SelectedItems(1)
    End With
End Function

' Helper function to get item category
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

' Function to count items in warenText
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

' Function to find the formula row for a tour (for extracting totals)
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

' Function to extract numeric value safely from mixed text/formulas
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

' Function to parse items from warenText
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
