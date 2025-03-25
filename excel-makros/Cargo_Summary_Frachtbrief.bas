Sub ProcessTours()
    ' Tour Processing Macro
    ' Processes tour data to create PDFs and a summary worksheet
    
    Dim ws As Worksheet
    Dim wsSummary As Worksheet
    Dim lastRow As Long, i As Long
    Dim tourNumber As String, lastTourNumber As String
    Dim tourName As String, tourDate As String
    Dim rng As Range
    Dim pdfPath As String
    Dim cell As Range
    Dim directTour As Boolean
    Dim tourWeight As Double, tourVolume As Double
    Dim tourABNumbers As String
    Dim tourItems As String
    Dim currentStopItems As String
    Dim lastStop As Long
    Dim currentStop As Long
    Dim isServiceCenter As Boolean
    
    ' Show status message
    Application.StatusBar = "Processing tours. Please wait..."
    
    ' Set reference to the active worksheet
    Set ws = ActiveSheet
    
    ' Check if summary sheet already exists, if not create it
    On Error Resume Next
    Set wsSummary = Worksheets("TourSummary")
    On Error GoTo 0
    
    If wsSummary Is Nothing Then
        Set wsSummary = Worksheets.Add(After:=Worksheets(Worksheets.count))
        wsSummary.Name = "TourSummary"
    Else
        wsSummary.Cells.Clear
    End If
    
    ' Create headers in summary sheet
    With wsSummary
        .Cells(1, 1).Value = "Tour_Name"
        .Cells(1, 2).Value = "Tour_Date"
        .Cells(1, 3).Value = "Tour_Type" ' Direct or Service Center
        .Cells(1, 4).Value = "Total_Weight (kg)"
        .Cells(1, 5).Value = "Total_Volume (m³)"
        .Cells(1, 6).Value = "AB_Numbers"
        .Cells(1, 7).Value = "Items_Per_Stop"
        
        ' Format headers
        .Range(.Cells(1, 1), .Cells(1, 7)).Font.Bold = True
        .Range(.Cells(1, 1), .Cells(1, 7)).Interior.Color = RGB(200, 200, 200)
    End With
    
    ' Find the last row with data
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    
    ' Create a folder for PDFs if it doesn't exist
    pdfPath = "C:\BML_Export\"
    If Dir(pdfPath, vbDirectory) = "" Then
        MkDir pdfPath
    End If
    
    lastTourNumber = ""
    Dim summaryRow As Long
    summaryRow = 2
    
    ' Process each row in the data
    For i = 2 To lastRow
        tourNumber = ws.Cells(i, 1).Value
        
        ' Skip rows that don't have a tour number or have formula cells with Max=
        If Not IsEmpty(tourNumber) And Not (IsNull(tourNumber)) And Not InStr(ws.Cells(i, "C").Value, "Max=") > 0 Then
            ' Check if we're at a new tour
            If tourNumber <> lastTourNumber Then
                ' Save previous tour data to summary sheet if not the first tour
                If lastTourNumber <> "" Then
                    WriteTourSummary wsSummary, summaryRow - 1
                    CreateTourPDF ws, lastTourNumber, pdfPath
                End If
                
                ' Start collecting data for the new tour
                lastTourNumber = tourNumber
                
                ' Parse the tour name and date
                ParseTourNameAndDate ws.Cells(i, 2).Value, tourName, tourDate
                
                ' Check if it's a Service Center tour
                isServiceCenter = InStr(ws.Cells(i, 2).Value, "SC ") > 0
                
                ' Add to summary sheet
                wsSummary.Cells(summaryRow, 1).Value = tourName
                wsSummary.Cells(summaryRow, 2).Value = tourDate
                wsSummary.Cells(summaryRow, 3).Value = IIf(isServiceCenter, "Service Center", "Direct Tour")
                
                ' Get total weight and volume from the formula cells
                ' These are usually in the row with Max=n in column C
                Dim formulaRow As Long
                formulaRow = FindFormulaRow(ws, tourNumber)
                
                If formulaRow > 0 Then
                    ' Extract values from the formula cells
                    wsSummary.Cells(summaryRow, 4).Value = ExtractNumericValue(ws.Cells(formulaRow, 4).Value) ' Weight
                    wsSummary.Cells(summaryRow, 5).Value = ExtractNumericValue(ws.Cells(formulaRow, 5).Value) ' Volume
                End If
                
                ' Initialize AB Numbers and Items
                tourABNumbers = ""
                tourItems = ""
                
                summaryRow = summaryRow + 1
            End If
            
            ' Process the stops and collect AB numbers and items for the current tour
            ' Only process rows with numeric stop order
            If IsNumeric(ws.Cells(i, 3).Value) Then
                currentStop = ws.Cells(i, 3).Value
                
                ' Add AB Number if not already added
                Dim abNumber As String
                abNumber = ws.Cells(i, 12).Value ' Column L
                
                If Len(abNumber) > 0 Then
                    If InStr(tourABNumbers, abNumber) = 0 Then
                        If Len(tourABNumbers) > 0 Then tourABNumbers = tourABNumbers & ", "
                        tourABNumbers = tourABNumbers & abNumber
                    End If
                End If
                
                ' Format the items for this stop
                currentStopItems = "Stop " & currentStop & ": " & vbCrLf
                currentStopItems = currentStopItems & FormatItemsList(ws.Cells(i, 47).Value) ' Column AV
                
                If Len(tourItems) > 0 Then tourItems = tourItems & vbCrLf & vbCrLf
                tourItems = tourItems & currentStopItems
                
                ' Update the summary sheet with the AB Numbers and Items
                Dim lastSummaryRow As Long
                lastSummaryRow = summaryRow - 1
                
                wsSummary.Cells(lastSummaryRow, 6).Value = tourABNumbers
                wsSummary.Cells(lastSummaryRow, 7).Value = tourItems
            End If
        End If
    Next i
    
    ' Process the last tour
    If lastTourNumber <> "" Then
        WriteTourSummary wsSummary, summaryRow - 1
        CreateTourPDF ws, lastTourNumber, pdfPath
    End If
    
    ' Format the summary sheet
    With wsSummary
        .Columns.AutoFit
        .Columns(7).ColumnWidth = 100 ' Make Items column wider
        
        ' Add borders
        .UsedRange.Borders.LineStyle = xlContinuous
        .UsedRange.Borders.weight = xlThin
    End With
    
    ' Reset status bar
    Application.StatusBar = False
    
    MsgBox "Tour processing complete. PDFs saved to " & pdfPath & " and summary created in the TourSummary worksheet.", vbInformation
End Sub

Function FindFormulaRow(ws As Worksheet, tourNumber As String) As Long
    ' Find the row with the formula cells for a tour (row with Max=n in column C)
    Dim lastRow As Long, i As Long
    
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = tourNumber And InStr(ws.Cells(i, 3).Value, "Max=") > 0 Then
            FindFormulaRow = i
            Exit Function
        End If
    Next i
    
    FindFormulaRow = 0
End Function

Function ExtractNumericValue(formulaText As String) As Double
    ' Extract numeric value from formula text like "?=2813,84"
    Dim numText As String
    
    If InStr(formulaText, "?=") > 0 Then
        numText = Replace(Mid(formulaText, InStr(formulaText, "?=") + 2), ",", ".")
        On Error Resume Next
        ExtractNumericValue = CDbl(numText)
        On Error GoTo 0
    ElseIf IsNumeric(Replace(formulaText, ",", ".")) Then
        ' Handle direct numeric values (like in your screenshot)
        ExtractNumericValue = CDbl(Replace(formulaText, ",", "."))
    Else
        ExtractNumericValue = 0
    End If
End Function

Sub ParseTourNameAndDate(tourNameText As String, ByRef tourName As String, ByRef tourDate As String)
    ' Parse tour name and date from format like "Wien 8 - 07.04." or "SC Wr. Neudorf 07.04."
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

Function ParseItems(itemsText As String) As Variant
    ' Parse items text into a 2D array with item details
    Dim items() As String
    Dim result() As String
    Dim i As Long, j As Long
    Dim parts() As String
    
    If Len(itemsText) = 0 Then
        ReDim result(0, 3)
        result(0, 0) = ""
        result(0, 1) = ""
        result(0, 2) = ""
        result(0, 3) = "No items"
        ParseItems = result
        Exit Function
    End If
    
    ' Split by the "----------" separator
    items = Split(itemsText, "----------")
    
    ' Count valid items
    Dim validItemsCount As Long
    validItemsCount = 0
    
    For i = 0 To UBound(items)
        If Len(Trim(items(i))) > 0 Then
            validItemsCount = validItemsCount + 1
        End If
    Next i
    
    ' Create result array
    ReDim result(validItemsCount - 1, 3)
    
    j = 0
    For i = 0 To UBound(items)
        If Len(Trim(items(i))) > 0 Then
            ' Remove extra spaces and line breaks
            items(i) = Trim(Replace(items(i), vbCrLf, ""))
            items(i) = Replace(items(i), vbCr, "")
            items(i) = Replace(items(i), vbLf, "")
            
            ' Parse the item parts (format: Number|OrderNumber|Reference|Description)
            parts = Split(items(i), "|")
            
            If UBound(parts) >= 3 Then
                result(j, 0) = parts(0) ' Item Number
                result(j, 1) = parts(1) ' Order Number
                result(j, 2) = parts(2) ' Reference
                result(j, 3) = parts(3) ' Description (contains both code and name)
            Else
                ' If not in expected format, just store the whole text in the Description
                result(j, 0) = ""
                result(j, 1) = ""
                result(j, 2) = ""
                result(j, 3) = items(i)
            End If
            
            j = j + 1
        End If
    Next i
    
    ParseItems = result
End Function

Function FormatItemsList(itemsText As String) As String
    ' Format the items list for better readability in the summary sheet
    Dim items As Variant
    Dim formattedList As String
    Dim i As Long
    
    If Len(itemsText) = 0 Then
        FormatItemsList = "No items"
        Exit Function
    End If
    
    ' Parse the items
    items = ParseItems(itemsText)
    
    formattedList = ""
    For i = LBound(items) To UBound(items)
        ' Add formatted item to the list
        If Len(formattedList) > 0 Then formattedList = formattedList & vbCrLf
        
        ' Extract just the item number and full description
        If Len(items(i, 0)) > 0 Then
            ' Get the full description part (after any spaces in the 4th column)
            Dim fullDesc As String
            fullDesc = Trim(items(i, 3))
            
            ' Some descriptions have code and name separated by spaces
            ' Find the position where the actual name starts (after spaces)
            Dim namePos As Long
            namePos = 1
            
            ' Look for the first non-space character after a series of spaces
            Dim inSpaces As Boolean
            inSpaces = False
            
            For j = 1 To Len(fullDesc)
                If Mid(fullDesc, j, 1) = " " Then
                    inSpaces = True
                ElseIf inSpaces Then
                    ' We found the first character after a series of spaces
                    namePos = j
                    Exit For
                End If
            Next j
            
            ' If we found a position where the name starts
            If namePos > 1 And namePos < Len(fullDesc) Then
                Dim itemCode As String
                Dim itemName As String
                
                itemCode = Trim(Left(fullDesc, namePos - 1))
                itemName = Trim(Mid(fullDesc, namePos))
                
                formattedList = formattedList & "• " & items(i, 0) & " - " & itemName
            Else
                ' Just use the full description if we can't split it
                formattedList = formattedList & "• " & items(i, 0) & " - " & fullDesc
            End If
        Else
            formattedList = formattedList & "• " & items(i, 3)
        End If
    Next i
    
    FormatItemsList = formattedList
End Function

Sub WriteTourSummary(wsSummary As Worksheet, rowNum As Long)
    ' Format the numbers in the summary row
    With wsSummary
        .Cells(rowNum, 4).NumberFormat = "#,##0.00"
        .Cells(rowNum, 5).NumberFormat = "#,##0.00"
    End With
End Sub

Sub CreateTourPDF(ws As Worksheet, tourNumber As String, pdfPath As String)
    ' Create a PDF file for the tour with stop information and freight details
    Dim tempWs As Worksheet
    Dim lastRow As Long, i As Long
    Dim startRow As Long, endRow As Long
    Dim tourName As String
    Dim pdfFileName As String
    Dim rng As Range
    Dim stopCount As Long
    
    ' Update status bar
    Application.StatusBar = "Creating PDF for Tour " & tourNumber & "..."
    
    ' Find all rows for this tour and count how many stops
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    stopCount = 0
    
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = tourNumber And IsNumeric(ws.Cells(i, 3).Value) Then
            stopCount = stopCount + 1
        End If
    Next i
    
    ' Find the tour name
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = tourNumber And IsNumeric(ws.Cells(i, 3).Value) Then
            tourName = ws.Cells(i, 2).Value
            Exit For
        End If
    Next i
    
    ' Process each stop for this tour
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = tourNumber And IsNumeric(ws.Cells(i, 3).Value) Then
            ' Create a PDF for this stop
            CreateStopFreightPDF ws, i, tourNumber, tourName, pdfPath
        End If
    Next i
    
    ' Also create a summary PDF with all stops for this tour
    CreateTourSummaryPDF ws, tourNumber, tourName, pdfPath
End Sub

Sub CreateStopFreightPDF(ws As Worksheet, rowNum As Long, tourNumber As String, tourName As String, pdfPath As String)
    ' Create a PDF file for a specific stop on a tour
    Dim tempWs As Worksheet
    Dim stopNum As Long
    Dim abNumber As String
    Dim itemsText As String
    Dim pdfFileName As String
    Dim i As Long
    Dim recipientName As String, recipientAddress As String, recipientCity As String
    Dim deliveryDate As String
    Dim buildingInfo As String, deliveryRemarks As String
    Dim stopWeight As Double, stopVolume As Double
    
    ' Get stop information
    stopNum = ws.Cells(rowNum, 3).Value
    abNumber = ws.Cells(rowNum, 12).Value ' Column L
    itemsText = ws.Cells(rowNum, 47).Value ' Column AV - Warenbeschreibung
    stopWeight = ws.Cells(rowNum, 4).Value ' Column D
    stopVolume = ws.Cells(rowNum, 5).Value ' Column E
    
    ' Get recipient info
    recipientName = ws.Cells(rowNum, 36).Value ' Column AJ
    recipientAddress = ws.Cells(rowNum, 37).Value ' Column AK
    recipientCity = ws.Cells(rowNum, 38).Value ' Column AL
    
    ' Get delivery info
    If Not IsEmpty(ws.Cells(rowNum, 16).Value) Then ' Column P
        deliveryDate = Format(ws.Cells(rowNum, 16).Value, "dd.MM.yyyy")
    Else
        deliveryDate = ""
    End If
    
    buildingInfo = ws.Cells(rowNum, 43).Value ' Column AQ
    deliveryRemarks = ws.Cells(rowNum, 44).Value ' Column AR
    
    ' Create a temporary worksheet for the PDF
    Set tempWs = Worksheets.Add
    tempWs.Name = "TempPDF_Stop" & stopNum
    
    ' Create the header for the PDF (similar to your example)
    With tempWs
        ' Set larger paper size (larger than A4 initially)
        .PageSetup.PaperSize = xlPaperA3
        
        ' Company logo position (top right)
        .Range("I1:K3").Merge
        .Range("I1").Value = "hali"
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
        .Range("B9").Value = recipientCity
        .Range("A7:A9").Font.Bold = True
        
        ' Delivery information
        .Range("E7").Value = "Delivery Date:"
        .Range("F7").Value = deliveryDate
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
        
        ' Replace the item parsing section with this:

        ' Parse and format the items
        If Len(itemsText) > 0 Then
            Dim items As Variant
            items = ParseItems(itemsText)
            
            Dim row As Long
            row = 14
            
            For i = LBound(items) To UBound(items)
                ' Item Number
                .Cells(row, 1).Value = items(i, 0)
                
                ' Item Name / Description - Split into code and name for better display
                Dim fullDesc As String
                fullDesc = Trim(items(i, 3))
                
                ' Find where the item name starts (after spaces)
                Dim namePos As Long, j As Long
                namePos = 1
                Dim inSpaces As Boolean
                inSpaces = False
                
                For j = 1 To Len(fullDesc)
                    If Mid(fullDesc, j, 1) = " " Then
                        inSpaces = True
                    ElseIf inSpaces Then
                        namePos = j
                        Exit For
                    End If
                Next j
                
                ' Display the full item name
                If namePos > 1 And namePos < Len(fullDesc) Then
                    ' Split into code and name
                    Dim itemCode As String
                    Dim itemName As String
                    
                    itemCode = Trim(Left(fullDesc, namePos - 1))
                    itemName = Trim(Mid(fullDesc, namePos))
                    
                    .Cells(row, 3).Value = itemName
                Else
                    ' Just display the whole thing
                    .Cells(row, 3).Value = fullDesc
                End If
                
                ' Get category
                Dim categoryCode As String
                categoryCode = GetItemCategory(fullDesc)
                .Cells(row, 7).Value = categoryCode
                
                row = row + 1
            Next i
        Else
            .Cells(14, 1).Value = "No items for this stop"
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

Sub CreateTourSummaryPDF(ws As Worksheet, tourNumber As String, tourName As String, pdfPath As String)
    ' Create a summary PDF for all stops in a tour
    Dim tempWs As Worksheet
    Dim lastRow As Long, i As Long
    Dim pdfFileName As String
    Dim stopRow As Long
    Dim totalWeight As Double, totalVolume As Double
    
    ' Find formula row to get totals
    Dim formulaRow As Long
    formulaRow = FindFormulaRow(ws, tourNumber)
    
    If formulaRow > 0 Then
        totalWeight = ExtractNumericValue(ws.Cells(formulaRow, 4).Value)
        totalVolume = ExtractNumericValue(ws.Cells(formulaRow, 5).Value)
    End If
    
    ' Create a temporary worksheet for the PDF
    Set tempWs = Worksheets.Add
    tempWs.Name = "TempPDF_TourSummary"
    
    ' Create the header for the PDF
    With tempWs
        .Range("A1:F1").Merge
        .Range("A1").Value = "Tour Summary: " & tourNumber & " - " & tourName
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 16
        .Range("A1").HorizontalAlignment = xlCenter
        
        .Range("A3").Value = "Total Weight (kg):"
        .Range("B3").Value = Format(totalWeight, "#,##0.00")
        .Range("A4").Value = "Total Volume (m³):"
        .Range("B4").Value = Format(totalVolume, "#,##0.00")
        .Range("A3:A4").Font.Bold = True
        
        ' Stops header
        .Range("A6").Value = "Stop"
        .Range("B6").Value = "AB Number"
        .Range("C6").Value = "Recipient"
        .Range("D6").Value = "Weight (kg)"
        .Range("E6").Value = "Volume (m³)"
        .Range("F6").Value = "Items Count"
        .Range("A6:F6").Font.Bold = True
        .Range("A6:F6").Interior.Color = RGB(220, 220, 220)
        
        ' Fill in the stops information
        stopRow = 7
        lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
        
        For i = 2 To lastRow
            If ws.Cells(i, 1).Value = tourNumber And IsNumeric(ws.Cells(i, 3).Value) Then
                .Cells(stopRow, 1).Value = ws.Cells(i, 3).Value ' Stop
                .Cells(stopRow, 2).Value = ws.Cells(i, 12).Value ' AB Number
                .Cells(stopRow, 3).Value = ws.Cells(i, 36).Value ' Recipient
                .Cells(stopRow, 4).Value = ws.Cells(i, 4).Value ' Weight
                .Cells(stopRow, 5).Value = ws.Cells(i, 5).Value ' Volume
                
                ' Count items
                Dim itemsCount As Long
                itemsCount = CountItems(ws.Cells(i, 47).Value) ' Column AV
                .Cells(stopRow, 6).Value = itemsCount
                
                stopRow = stopRow + 1
            End If
        Next i
        
        ' Format numbers
        .Range("D7:E" & stopRow - 1).NumberFormat = "#,##0.00"
        
        ' Add footer with total row
        .Cells(stopRow, 1).Value = "TOTAL:"
        .Cells(stopRow, 1).Font.Bold = True
        .Cells(stopRow, 4).Value = totalWeight
        .Cells(stopRow, 5).Value = totalVolume
        .Cells(stopRow, 6).Formula = "=SUM(F7:F" & stopRow - 1 & ")"
        .Range("D" & stopRow & ":F" & stopRow).Font.Bold = True
        .Range("D" & stopRow & ":E" & stopRow).NumberFormat = "#,##0.00"
        
        ' Format the worksheet
        .Columns.AutoFit
        .Range("A6:F" & stopRow).Borders.LineStyle = xlContinuous
        .Range("A6:F" & stopRow).Borders.weight = xlThin
    End With
    
    ' Save as PDF
    pdfFileName = pdfPath & "Tour_" & tourNumber & "_Summary.pdf"
    
    tempWs.PageSetup.Orientation = xlLandscape
    tempWs.PageSetup.FitToPagesWide = 1
    tempWs.PageSetup.FitToPagesTall = False
    
    tempWs.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfFileName, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    
    ' Delete the temporary worksheet
    Application.DisplayAlerts = False
    tempWs.Delete
    Application.DisplayAlerts = True
End Sub

Function CountItems(itemsText As String) As Long
    ' Count the number of items in an items list
    If Len(itemsText) = 0 Then
        CountItems = 0
        Exit Function
    End If
    
    Dim count As Long
    count = 0
    
    ' Count number of line breaks or separators
    count = (Len(itemsText) - Len(Replace(itemsText, "----------", ""))) / Len("----------") + 1
    
    CountItems = count
End Function

Function GetItemCategory(ByVal itemName As String) As String
    ' Returns a category symbol based on the item name/description
    ' Using text symbols instead of emoji for VBA compatibility
    
    Dim itemNameUpper As String
    itemNameUpper = UCase(itemName)
    
    ' Desks and tables
    If InStr(itemNameUpper, "TISCH") > 0 Or _
       InStr(itemNameUpper, "DESK") > 0 Or _
       InStr(itemNameUpper, "TABLE") > 0 Or _
       InStr(itemNameUpper, "ARBEITSTISCH") > 0 Or _
       InStr(itemNameUpper, "MXTI") > 0 Or _
       InStr(itemNameUpper, "MXTK") > 0 Or _
       InStr(itemNameUpper, "BESPRECHUNGSTISCH") > 0 Then
        GetItemCategory = "D"  ' Desk
        
    ' Chairs and seating
    ElseIf InStr(itemNameUpper, "STUHL") > 0 Or _
           InStr(itemNameUpper, "CHAIR") > 0 Or _
           InStr(itemNameUpper, "SESSEL") > 0 Or _
           InStr(itemNameUpper, "DREHSTUHL") > 0 Or _
           InStr(itemNameUpper, "SITZ") > 0 Then
        GetItemCategory = "C" ' Chair
        
    ' Storage and cabinets
    ElseIf InStr(itemNameUpper, "SCHRANK") > 0 Or _
           InStr(itemNameUpper, "CABINET") > 0 Or _
           InStr(itemNameUpper, "STORAGE") > 0 Or _
           InStr(itemNameUpper, "REGAL") > 0 Or _
           InStr(itemNameUpper, "CONT-RC") > 0 Or _
           InStr(itemNameUpper, "ROLLCONT") > 0 Then
        GetItemCategory = "S" ' Storage
        
    ' Panels and dividers
    ElseIf InStr(itemNameUpper, "PANEEL") > 0 Or _
           InStr(itemNameUpper, "PANEL") > 0 Or _
           InStr(itemNameUpper, "NOOVA") > 0 Or _
           InStr(itemNameUpper, "DIVIDER") > 0 Or _
           InStr(itemNameUpper, "TRENNWAND") > 0 Then
        GetItemCategory = "P" ' Panel
        
    ' Electronic components
    ElseIf InStr(itemNameUpper, "ELEKTR") > 0 Or _
           InStr(itemNameUpper, "ELECTRIC") > 0 Or _
           InStr(itemNameUpper, "KABEL") > 0 Or _
           InStr(itemNameUpper, "CABLE") > 0 Or _
           InStr(itemNameUpper, "STECKDOSE") > 0 Then
        GetItemCategory = "E" ' Electric
        
    ' Height adjustable desks
    ElseIf InStr(itemNameUpper, "HEBETISCH") > 0 Or _
           InStr(itemNameUpper, "FLUX") > 0 Or _
           InStr(itemNameUpper, "HEIGHT ADJUST") > 0 Then
        GetItemCategory = "H" ' Height adjustable
        
    ' Accessories
    ElseIf InStr(itemNameUpper, "POLSTER") > 0 Or _
           InStr(itemNameUpper, "CUSHION") > 0 Or _
           InStr(itemNameUpper, "PAD") > 0 Or _
           InStr(itemNameUpper, "ZUBEHÖR") > 0 Or _
           InStr(itemNameUpper, "ACCESSORY") > 0 Then
        GetItemCategory = "A" ' Accessories
        
    ' Services
    ElseIf InStr(itemNameUpper, "MONTAGE") > 0 Or _
           InStr(itemNameUpper, "SERVICE") > 0 Or _
           InStr(itemNameUpper, "INSTALLATION") > 0 Then
        GetItemCategory = "M" ' Montage/Service
        
    ' Default for unknown items
    Else
        GetItemCategory = "X" ' Default
    End If
End Function
