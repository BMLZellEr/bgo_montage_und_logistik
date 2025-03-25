Sub ProcessTours()
    ' Tour Processing Macro - Fixed version
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim wsSummary As Worksheet
    Dim lastRow As Long, i As Long
    Dim tourNumber As String, lastTourNumber As String
    Dim tourName As String, tourDate As String
    Dim isServiceCenter As Boolean
    
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
    
    ' Process each row to build tour information
    For i = 2 To lastRow
        ' Skip empty rows
        If Not IsEmpty(ws.Cells(i, 1).Value) Then
            tourNumber = Trim(CStr(ws.Cells(i, 1).Value))
            
            ' Process formula rows (with Max= in column C)
            If InStr(CStr(ws.Cells(i, 3).Value), "Max=") > 0 Then
                ' Store the weight and volume for this tour
                If Not tourSummary.Exists(tourNumber) Then
                    ' Weight and volume from formula cells
                    Dim weightStr As String, volumeStr As String
                    weightStr = CStr(ws.Cells(i, 4).Value)
                    volumeStr = CStr(ws.Cells(i, 5).Value)
                    
                    ' Extract numeric values
                    Dim weight As Double, volume As Double
                    weight = ExtractNumericValue(weightStr)
                    volume = ExtractNumericValue(volumeStr)
                    
                    ' Create array to store tour data
                    Dim tourData(5) As Variant
                    tourData(0) = "" ' Tour name (filled later)
                    tourData(1) = "" ' Tour date (filled later)
                    tourData(2) = "" ' Tour type (filled later)
                    tourData(3) = weight ' Total weight
                    tourData(4) = volume ' Total volume
                    tourData(5) = "" ' AB Numbers (filled later)
                    
                    ' Add to dictionary
                    tourSummary.Add tourNumber, tourData
                End If
            ' Process regular rows (with stop information)
            ElseIf IsNumeric(ws.Cells(i, 3).Value) Then
                ' Get tour name/date if not already in dictionary
                If Not tourSummary.Exists(tourNumber) Then
                    ' Create a new tour entry
                    Dim newTourData(5) As Variant
                    newTourData(0) = "" ' Tour name (filled below)
                    newTourData(1) = "" ' Tour date (filled below)
                    newTourData(2) = "" ' Tour type (filled below)
                    newTourData(3) = 0 ' Total weight (filled from formula row)
                    newTourData(4) = 0 ' Total volume (filled from formula row)
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
    
CleanExit:
    ' Reset status bar and screen updating
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "Tour summary created in the TourSummary worksheet.", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description & " (Line: " & Erl & ")", vbCritical
    Resume CleanExit
End Sub

Function FormatCombinedItems(artikelTypen As String, warenText As String) As String
    ' Format the items list combining Artikeltypen and Warenbeschreibung
    On Error Resume Next
    
    ' Default for empty input
    If Len(artikelTypen) = 0 And Len(warenText) = 0 Then
        FormatCombinedItems = "No items"
        Exit Function
    End If
    
    Dim formattedList As String
    formattedList = ""
    
    ' If we have Warenbeschreibung data
    If Len(warenText) > 0 Then
        ' Split the items by the "----------" separator
        Dim itemsArray() As String
        itemsArray = Split(warenText, "----------")
        
        ' Parse Artikeltypen into an array if not empty
        Dim artikelArray() As String
        Dim artikelValues As Variant
        
        If Len(artikelTypen) > 0 Then
            ' Split by commas if it contains them
            If InStr(artikelTypen, ",") > 0 Then
                artikelArray = Split(artikelTypen, ",")
            Else
                ' If it's one long string of numbers, try to split it intelligently
                Dim tempArtikel As String
                tempArtikel = artikelTypen
                
                ' Replace common separators if present
                tempArtikel = Replace(tempArtikel, ".", ",")
                tempArtikel = Replace(tempArtikel, ";", ",")
                
                artikelArray = Split(tempArtikel, ",")
            End If
        End If
        
        ' Format each item
        Dim i As Long
        For i = 0 To UBound(itemsArray)
            If Len(Trim(itemsArray(i))) > 0 Then
                ' Clean up the item text
                Dim itemLine As String
                itemLine = Trim(Replace(Replace(Replace(itemsArray(i), vbCrLf, ""), vbCr, ""), vbLf, ""))
                
                ' Add formatted item to the list
                If Len(formattedList) > 0 Then formattedList = formattedList & vbCrLf
                
                ' Split by pipe character to get the item parts
                Dim parts() As String
                If InStr(itemLine, "|") > 0 Then
                    parts = Split(itemLine, "|")
                    
                    If UBound(parts) >= 3 Then
                        ' Get position/artikeltyp for this item if available
                        Dim itemPosition As String
                        itemPosition = ""
                        
                        If i < UBound(artikelArray) + 1 Then
                            itemPosition = Trim(artikelArray(i))
                        End If
                        
                        ' Get all item numbers (NR1|NR2|NR3|NR4)
                        Dim allItemNumbers As String
                        allItemNumbers = parts(0) & "|" & parts(1) & "|" & parts(2) & "|" & parts(3)
                        
                        ' Extract item description from part 3
                        Dim itemDesc As String
                        itemDesc = Trim(parts(3))
                        
                        ' Format as: "• Position | Item Numbers | Item Description"
                        If Len(itemPosition) > 0 Then
                            formattedList = formattedList & "• " & itemPosition & " | " & allItemNumbers & " | " & itemDesc
                        Else
                            formattedList = formattedList & "• " & allItemNumbers & " | " & itemDesc
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
    Else
        ' If no warenText but we have artikelTypen
        formattedList = "• " & artikelTypen
    End If
    
    FormatCombinedItems = formattedList
End Function

Function ExtractNumericValue(formulaText As String) As Double
    ' Extract numeric value from formula text like "?=2813,84"
    On Error Resume Next
    
    ' Default value
    ExtractNumericValue = 0
    
    ' Check for empty input
    If IsEmpty(formulaText) Or formulaText = "" Then Exit Function
    
    Dim numText As String
    
    ' Check for the sum symbol (may appear differently in different encodings)
    If InStr(formulaText, "?=") > 0 Then
        numText = Mid(formulaText, InStr(formulaText, "?=") + 2)
        numText = Replace(numText, ",", ".")
        ExtractNumericValue = CDbl(numText)
    ' Try alternative symbols for sum
    ElseIf InStr(formulaText, "S=") > 0 Then
        numText = Mid(formulaText, InStr(formulaText, "S=") + 2)
        numText = Replace(numText, ",", ".")
        ExtractNumericValue = CDbl(numText)
    ElseIf InStr(formulaText, "?=") > 0 Then
        numText = Mid(formulaText, InStr(formulaText, "?=") + 2)
        numText = Replace(numText, ",", ".")
        ExtractNumericValue = CDbl(numText)
    ' Handle direct numeric values
    ElseIf IsNumeric(Replace(formulaText, ",", ".")) Then
        ExtractNumericValue = CDbl(Replace(formulaText, ",", "."))
    End If
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
