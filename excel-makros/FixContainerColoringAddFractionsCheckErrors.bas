Sub FixContainerColoringAddFractionsCheckErrors()
    Dim ws As Worksheet
    Dim topRange As Range, lastRows As Range
    Dim topCell As Range, lastCell As Range
    Dim containerMap As Object
    Dim colorMap As Object
    Dim containerCount As Object
    Dim containerCells As Object
    Dim containerDates As Object ' New dictionary to store container dates
    Dim i As Integer, j As Integer, k As Integer
    Dim topValue As Variant, lastValue As String
    Dim colorCode As Long
    Dim cellDate As Variant, cellTime As Variant
    Dim sortedKeys() As String, temp As Variant
    Dim Key As Variant
    Dim fractionCell As Range
    Dim row As Long, col As Long
    Dim textColor As Long ' Added for contrasting text color
    Dim containerDay As Integer, tourDay As Integer ' New variables for date comparison

    Set ws = ThisWorkbook.Sheets("Tourenplan Dispo")
    Set containerMap = CreateObject("Scripting.Dictionary")
    Set colorMap = CreateObject("Scripting.Dictionary")
    Set containerCount = CreateObject("Scripting.Dictionary")
    Set containerCells = CreateObject("Scripting.Dictionary")
    Set containerDates = CreateObject("Scripting.Dictionary") ' Initialize container dates dictionary

    ' Touren Felder (Top Section)
    Dim topColumns As Variant
    topColumns = Array(4, 7, 10, 13, 16, 19) ' D, G, J, M, P, S
    
    Dim dateColumns As Variant
    dateColumns = Array(2, 5, 8, 11, 14, 17) ' B, E, H, K, N, Q
    
    ' WAB Felder (Bottom Section)
    Dim lastColumns As Variant
    lastColumns = Array(2, 5, 8, 11, 14, 17) ' B, E, H, K, N, Q
    
    ' STEP 0: First collect all container dates from the bottom section
    For i = LBound(lastColumns) To UBound(lastColumns)
        Set lastRows = ws.Range(ws.Cells(35, lastColumns(i)), ws.Cells(38, lastColumns(i)))
        
        For Each lastCell In lastRows
            lastValue = CleanNumber(lastCell.value)
            If lastValue <> "" Then
                ' Get the date from the column header (row 2)
                Dim containerDate As Variant
                containerDate = ws.Cells(2, lastColumns(i)).value
                
                ' Store the container and its date
                If IsDate(containerDate) Then
                    ' If container already exists in dictionary, only update if this date is earlier
                    If containerDates.Exists(lastValue) Then
                        If CDate(containerDate) < CDate(containerDates(lastValue)) Then
                            containerDates(lastValue) = containerDate
                        End If
                    Else
                        containerDates.Add lastValue, containerDate
                    End If
                End If
            End If
        Next lastCell
    Next i
    
    ' STEP 1: Collect all container cells and assign colors
    For i = LBound(topColumns) To UBound(topColumns)
        Set topRange = ws.Range(ws.Cells(3, topColumns(i)), ws.Cells(23, topColumns(i)))
        
        For Each topCell In topRange
            topValue = CleanNumber(topCell.value)
            If topValue <> "" Then
                ' Extract Date from Row 2
                cellDate = ws.Cells(2, dateColumns(i)).value
                
                ' MERGE TOUR NAME CELLS FOR ALL TOURS
                Dim tourRow As Long
                tourRow = topCell.row - 1
                
                ' Only process if this is a row that would contain a tour name (first row of each 3-row tour)
                If tourRow >= 3 And (tourRow - 3) Mod 3 = 0 Then
                    ' First check if these cells are merged, if not, merge them
                    On Error Resume Next
                    If Not ws.Cells(tourRow, dateColumns(i)).MergeCells Then
                        ' Unmerge just to be safe (in case there's partial merging)
                        ws.Range(ws.Cells(tourRow, dateColumns(i)), ws.Cells(tourRow, dateColumns(i) + 2)).UnMerge
                        
                        ' Save the tour name
                        Dim tourNameForMerging As String
                        tourNameForMerging = ws.Cells(tourRow, dateColumns(i)).value
                        
                        ' Clear any values in these cells that might interfere with merging
                        ws.Cells(tourRow, dateColumns(i) + 1).ClearContents
                        ws.Cells(tourRow, dateColumns(i) + 2).ClearContents
                        
                        ' Merge the cells and restore the tour name
                        ws.Range(ws.Cells(tourRow, dateColumns(i)), ws.Cells(tourRow, dateColumns(i) + 2)).Merge
                        ws.Cells(tourRow, dateColumns(i)).value = tourNameForMerging
                    End If
                    On Error GoTo 0
                End If
                
                ' IMPORTANT FIX: Extract Time from row DIRECTLY below container number
                ' This is the critical fix needed
                cellTime = ws.Cells(topCell.row + 1, topCell.Column).value
                
                ' Create a sort key for this container occurrence
                Dim sortKey As String
                If IsDate(cellDate) Then
                    sortKey = Format(cellDate, "yyyy-mm-dd")
                Else
                    sortKey = "1900-01-01"
                End If
                
                ' FIX: Handle time values correctly for sorting
                ' Convert any time format to a standardized string for sorting
                If Not IsEmpty(cellTime) Then
                    ' For numeric time formats (e.g., 8.5 representing 8:30)
                    If IsNumeric(cellTime) Then
                        sortKey = sortKey & " " & Format(cellTime, "hh:mm:ss")
                    ' For time strings like "08:00" or "0800"
                    ElseIf VarType(cellTime) = vbString Then
                        If Len(cellTime) = 4 And IsNumeric(cellTime) Then
                            ' Handle "0800" format
                            sortKey = sortKey & " " & Left(cellTime, 2) & ":" & Right(cellTime, 2) & ":00"
                        ElseIf InStr(cellTime, ":") > 0 Then
                            ' Handle "08:00" format
                            sortKey = sortKey & " " & cellTime & ":00"
                        Else
                            sortKey = sortKey & " 99:99:99" ' Invalid time format
                        End If
                    Else
                        sortKey = sortKey & " 99:99:99" ' Unknown time format
                    End If
                Else
                    sortKey = sortKey & " 99:99:99" ' No time value
                End If
                
                ' Ensure uniqueness by adding cell address at the end
                sortKey = sortKey & "_" & topCell.Address
                
                ' Store container cell with sort key
                If Not containerCells.Exists(topValue) Then
                    Set containerCells(topValue) = CreateObject("Scripting.Dictionary")
                End If
                
                ' Use the sortKey as dictionary key to enable proper sorting
                containerCells(topValue).Add sortKey, topCell
                
                ' Count containers
                If containerCount.Exists(topValue) Then
                    containerCount(topValue) = containerCount(topValue) + 1
                Else
                    containerCount.Add topValue, 1
                End If
                
                ' Assign color if not already assigned
                If Not colorMap.Exists(topValue) Then
                    Randomize topValue
                    ' Ensure we never generate red (255,0,0) as a random color
                    Dim r As Integer, g As Integer, b As Integer
                    r = Int(Rnd * 230) ' Max 230 instead of 255 to avoid pure red
                    g = Int(Rnd * 255)
                    b = Int(Rnd * 255)
                    ' If it's too close to red, add some green to make it distinct
                    If r > 200 And g < 50 And b < 50 Then
                        g = g + 100 ' Add some green to avoid looking like the error color
                    End If
                    colorCode = RGB(r, g, b)
                    colorMap.Add topValue, colorCode
                End If
                colorCode = colorMap(topValue)
                
                ' Get contrasting text color (black or white)
                textColor = GetContrastingTextColor(colorCode)
                
                ' NEW: Check if container date is later than tour date
                Dim isInvalidDate As Boolean
                isInvalidDate = False
                
                If containerDates.Exists(topValue) And IsDate(cellDate) And IsDate(containerDates(topValue)) Then
                    ' Compare dates - if container is scheduled after tour, mark as invalid
                    If CDate(containerDates(topValue)) > CDate(cellDate) Then
                        isInvalidDate = True
                        ' If invalid, override with red color
                        colorCode = RGB(255, 0, 0)
                        textColor = RGB(255, 255, 255) ' White text on red background
                    End If
                End If
                
                ' Apply color to top section
                topCell.Interior.Color = colorCode
                topCell.Font.Color = textColor
                
                ' Also color time cell below
                topCell.Offset(0, 0).Interior.Color = colorCode
                topCell.Offset(0, 0).Font.Color = textColor
            End If
        Next topCell
    Next i
    
    ' STEP 2: Assign Colors for WAB Fields
    For i = LBound(topColumns) To UBound(topColumns)
        Set lastRows = ws.Range(ws.Cells(35, lastColumns(i)), ws.Cells(38, lastColumns(i)))
        
        For Each lastCell In lastRows
            lastValue = CleanNumber(lastCell.value)
            If lastValue <> "" Then
                ' Assign color if not already assigned
                If Not colorMap.Exists(lastValue) Then
                    If containerCount.Exists(lastValue) Then
                        Randomize lastValue
                        colorCode = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
                        colorMap.Add lastValue, colorCode
                    End If
                End If
                colorCode = colorMap(lastValue)
                
                ' Get contrasting text color (black or white)
                textColor = GetContrastingTextColor(colorCode)
                
                ' Apply color and text color to cells
                lastCell.Interior.Color = colorCode
                lastCell.Font.Color = textColor
                
                lastCell.Offset(0, 1).Interior.Color = colorCode
                lastCell.Offset(0, 1).Font.Color = textColor
                
                lastCell.Offset(1, 0).Interior.Color = colorCode
                lastCell.Offset(1, 0).Font.Color = textColor
                
                lastCell.Offset(1, 1).Interior.Color = colorCode
                lastCell.Offset(1, 1).Font.Color = textColor
                
                lastCell.Offset(0, 2).Interior.Color = colorCode
                lastCell.Offset(0, 2).Font.Color = textColor
                
                lastCell.Offset(1, 2).Interior.Color = colorCode
                lastCell.Offset(1, 2).Font.Color = textColor
            End If
        Next lastCell
    Next i
    
    ' STEP 3: Assign fractions globally for each container
    For Each topValue In containerCells.Keys
        If containerCells(topValue).count > 0 Then
            ' Get all sort keys for this container
            ReDim sortedKeys(containerCells(topValue).count - 1)
            j = 0
            For Each Key In containerCells(topValue).Keys
                sortedKeys(j) = Key
                j = j + 1
            Next Key
            
            ' Sort the keys to ensure correct order by date + time
            For j = LBound(sortedKeys) To UBound(sortedKeys) - 1
                For k = j + 1 To UBound(sortedKeys)
                    If sortedKeys(j) > sortedKeys(k) Then
                        temp = sortedKeys(j)
                        sortedKeys(j) = sortedKeys(k)
                        sortedKeys(k) = temp
                    End If
                Next k
            Next j
            
            ' Assign fractions in order
            For j = LBound(sortedKeys) To UBound(sortedKeys)
                ' Get the container cell
                Set topCell = containerCells(topValue)(sortedKeys(j))
                
                ' CRITICAL: Get EXACT column and row for the fraction cell
                col = topCell.Column - 1
                row = topCell.row
                
                ' Now use direct cell reference - avoid using Offset which might introduce errors
                ws.Cells(row, col).NumberFormat = "@"
                
                ' Get the tour cell and check if it has an invalid date
                Set topCell = containerCells(topValue)(sortedKeys(j))
                Dim tourCellDate As Variant
                Dim tourColIndex As Integer
                tourColIndex = GetArrayIndex(topColumns, topCell.Column)
                If tourColIndex >= 0 Then
                    tourCellDate = ws.Cells(2, dateColumns(tourColIndex)).value
                End If
                
                Dim isDateError As Boolean
                isDateError = False
                
                If containerDates.Exists(topValue) And IsDate(tourCellDate) And IsDate(containerDates(topValue)) Then
                    If CDate(containerDates(topValue)) > CDate(tourCellDate) Then
                        isDateError = True
                        ' Format the Date Error text with red background and black bold text
                        ws.Cells(row, col).value = "Date Error"
                        ws.Cells(row, col).Interior.Color = RGB(255, 0, 0)
                        ws.Cells(row, col).Font.Color = RGB(0, 0, 0)
                        ws.Cells(row, col).Font.Bold = True
                    End If
                End If
                
                ' Only use the fraction if it's not a date error
                If Not isDateError Then
                    ws.Cells(row, col).value = (j + 1) & "/" & containerCells(topValue).count
                End If
            Next j
        End If
    Next topValue
    
    ' VERIFICATION STEP: Double-check no values in cells below fractions
    ' This is a safeguard to ensure no values appear where they shouldn't
    For i = LBound(topColumns) To UBound(topColumns)
        Set topRange = ws.Range(ws.Cells(3, topColumns(i)), ws.Cells(23, topColumns(i)))
        
        For Each topCell In topRange
            If topCell.value <> "" Then
                ' Check the cell below the fraction
                On Error Resume Next
                ws.Cells(topCell.row + 1, topCell.Column - 1).value = ""
                On Error GoTo 0
            End If
        Next topCell
    Next i
    
    ' STEP 3.5: Handle DIRECT_CONTAINER Tours (rows 25-33)
    ' Dictionary to track unique direct tours for unique coloring
    Dim directTourColors As Object
    Set directTourColors = CreateObject("Scripting.Dictionary")
    Dim nextLightColorIndex As Integer
    nextLightColorIndex = 0
    
    ' Arrays to track direct tours for each day
    Dim directTourCount(5) As Integer ' One for each day (0-5)
    
    ' Initialize direct tour counters
    For i = 0 To 5
        directTourCount(i) = 0
    Next i
    
    ' Scan for direct tours
    For i = LBound(dateColumns) To UBound(dateColumns)
        ' Loop through each tour (3 rows per tour)
        For row = 3 To 23 Step 3
            ' Get the tour name
            Dim tourName As String
            tourName = Trim(ws.Cells(row, dateColumns(i)).value)
            
            ' Check if this is a direct tour (not starting with excluded prefixes)
            If tourName <> "" And _
               Not (Left(tourName, 2) = "SC" Or _
                    Left(tourName, 4) = "ULGB" Or _
                    Left(tourName, 14) = "PLANT SC LEITER") Then
                
                ' This is a direct tour - get a light background color
                Dim lightColor As Long
                Dim directTourKey As String
                
                ' Create a unique key for this tour (tour name + container number)
                Dim containerCol As Integer
                containerCol = dateColumns(i) + 2
                Dim containerNumber As String
                containerNumber = CleanNumber(ws.Cells(row + 1, containerCol).value)
                directTourKey = tourName & "_" & containerNumber
                
                ' If we haven't seen this tour before, assign a new light color
                If Not directTourColors.Exists(directTourKey) Then
                    ' Create a series of light pastel colors
                    Select Case nextLightColorIndex
                        Case 0
                            lightColor = RGB(235, 235, 255) ' Light blue
                        Case 1
                            lightColor = RGB(255, 235, 235) ' Light red
                        Case 2
                            lightColor = RGB(235, 255, 235) ' Light green
                        Case 3
                            lightColor = RGB(255, 255, 235) ' Light yellow
                        Case 4
                            lightColor = RGB(255, 235, 255) ' Light magenta
                        Case 5
                            lightColor = RGB(235, 255, 255) ' Light cyan
                        Case Else
                            ' If we run out of predefined colors, create a random light color
                            Randomize
                            lightColor = RGB(200 + Int(Rnd * 55), 200 + Int(Rnd * 55), 200 + Int(Rnd * 55))
                    End Select
                    
                    ' Store this color for this tour
                    directTourColors.Add directTourKey, lightColor
                    
                    ' Increment for next new tour
                    nextLightColorIndex = nextLightColorIndex + 1
                Else
                    ' Use the existing color for this tour
                    lightColor = directTourColors(directTourKey)
                End If
                
                ' Process this direct tour
                Dim startCol As Integer, endCol As Integer
                startCol = dateColumns(i)
                endCol = dateColumns(i) + 2
                
                ' IMPORTANT: Fix tour name spanning in the Upper Section tours
                ' First check if these cells are merged, if not, merge them
                On Error Resume Next
                If Not ws.Cells(row, startCol).MergeCells Then
                    ' Unmerge just to be safe (in case there's partial merging)
                    ws.Range(ws.Cells(row, startCol), ws.Cells(row, endCol)).UnMerge
                    
                    ' Save the tour name
                    Dim originalTourName As String
                    originalTourName = ws.Cells(row, startCol).value
                    
                    ' Clear any values in these cells that might interfere with merging
                    ws.Cells(row, startCol + 1).ClearContents
                    ws.Cells(row, startCol + 2).ClearContents
                    
                    ' Merge the cells and restore the tour name
                    ws.Range(ws.Cells(row, startCol), ws.Cells(row, endCol)).Merge
                    ws.Cells(row, startCol).value = originalTourName
                End If
                On Error GoTo 0
                
                ' Color the entire tour area with light background
                ws.Range(ws.Cells(row, startCol), ws.Cells(row + 2, endCol)).Interior.Color = lightColor
                
                ' Set all text to black in the tour
                ws.Range(ws.Cells(row, startCol), ws.Cells(row + 2, endCol)).Font.Color = RGB(0, 0, 0)
                
                ' Store this for copying to DIRECT_CONTAINER section
                If directTourCount(i) < 3 Then
                    ' Copy to DIRECT_CONTAINER section (rows 25-33)
                    Dim destRow As Integer
                    destRow = 25 + (directTourCount(i) * 3)
                    
                    ' Clear the specific cells we're going to use
                    ws.Range(ws.Cells(destRow, startCol), ws.Cells(destRow + 2, endCol)).ClearContents
                    ws.Range(ws.Cells(destRow, startCol), ws.Cells(destRow + 2, endCol)).Interior.ColorIndex = xlNone
                    
                    ' IMPORTANT: Explicitly handle cell merging for proper display
                    
                    ' 1. First unmerge any existing merged cells in this area (to be safe)
                    On Error Resume Next
                    ws.Range(ws.Cells(destRow, startCol), ws.Cells(destRow, endCol)).UnMerge
                    On Error GoTo 0
                    
                    ' 2. Set the tour name and merge cells for proper spanning
                    ws.Cells(destRow, startCol).value = tourName
                    ws.Range(ws.Cells(destRow, startCol), ws.Cells(destRow, endCol)).Merge
                    
                    ' 3. Copy tour number (like T-Nr. 474442) if present
                    If Not IsEmpty(ws.Cells(row + 1, startCol).value) Then
                        ws.Cells(destRow + 1, startCol).value = ws.Cells(row + 1, startCol).value
                    End If
                    
                    ' 4. Copy container number
                    ws.Cells(destRow + 1, endCol).value = containerNumber
                    
                    ' 5. Copy time
                    ws.Cells(destRow + 2, endCol).value = ws.Cells(row + 2, endCol).value
                    
                    ' 5. Set background color for the cells
                    ws.Range(ws.Cells(destRow, startCol), ws.Cells(destRow + 2, endCol)).Interior.Color = lightColor
                    
                    ' 6. Set all text to black
                    ws.Range(ws.Cells(destRow, startCol), ws.Cells(destRow + 2, endCol)).Font.Color = RGB(0, 0, 0)
                    
                    ' Increment counter
                    directTourCount(i) = directTourCount(i) + 1
                End If
            End If
        Next row
    Next i
    
    ' STEP 4: Count and calculate total workers per day
    ' Define the worker calculation fields
    Dim workerTotalColumns As Variant
    workerTotalColumns = Array(4, 7, 10, 13, 16, 19) ' D, G, J, M, P, S
    
    ' Initialize arrays to store count for each day
    Dim bmlWorkers(5) As Integer ' One for each day (0-5)
    Dim lpWorkers(5) As Integer ' One for each day (0-5)
    
    ' Loop through days
    For i = LBound(topColumns) To UBound(topColumns)
        ' Reset counters for this day
        bmlWorkers(i) = 0
        lpWorkers(i) = 0
        
        ' Loop through each tour row for this day
        For row = 3 To 23 Step 3
            ' Get the worker text
            Dim workerText As String
            workerText = Trim(ws.Cells(row + 2, dateColumns(i)).value)
            
            ' Check if there's actual text (not empty)
            If workerText <> "" Then
                ' Parse the worker text to count BML and LP workers
                Dim bmlCount As Integer, lpCount As Integer
                ParseWorkerText workerText, bmlCount, lpCount
                
                ' Add to day totals
                bmlWorkers(i) = bmlWorkers(i) + bmlCount
                lpWorkers(i) = lpWorkers(i) + lpCount
            End If
        Next row
        
        ' Write the totals to the summary fields (row 34)
        ' If we have both BML and LP workers, format as "BML+LP"
        If lpWorkers(i) > 0 Then
            ws.Cells(34, workerTotalColumns(i)).value = bmlWorkers(i) & "+" & lpWorkers(i)
        Else
            ws.Cells(34, workerTotalColumns(i)).value = bmlWorkers(i)
        End If
    Next i
    
    ' Cleanup
    Set containerMap = Nothing
    Set colorMap = Nothing
    Set containerCount = Nothing
    Set containerCells = Nothing
    Set containerDates = Nothing
    Set directTourColors = Nothing
End Sub

' Function to clean numbers (removes text)
Function CleanNumber(value As String) As String
    Dim cleanedValue As String
    Dim i As Integer
    cleanedValue = ""
    
    If Not IsEmpty(value) Then
        For i = 1 To Len(value)
            If Mid(value, i, 1) Like "#" Then
                cleanedValue = cleanedValue & Mid(value, i, 1)
            End If
        Next i
    End If
    
    CleanNumber = Trim(cleanedValue)
End Function

' Function to find index of value in array
Function GetArrayIndex(arr As Variant, valueToFind As Variant) As Integer
    Dim i As Integer
    GetArrayIndex = -1
    
    For i = LBound(arr) To UBound(arr)
        If arr(i) = valueToFind Then
            GetArrayIndex = i
            Exit Function
        End If
    Next i
End Function

' Function to determine if a color is light or dark and return a contrasting color
Function GetContrastingTextColor(ByVal backgroundColor As Long) As Long
    ' Extract RGB components
    Dim r As Integer, g As Integer, b As Integer
    r = backgroundColor Mod 256
    g = (backgroundColor \ 256) Mod 256
    b = (backgroundColor \ 65536) Mod 256
    
    ' Calculate luminance (perceived brightness)
    ' Using the formula: 0.299*R + 0.587*G + 0.114*B
    Dim luminance As Double
    luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
    
    ' If color is dark (luminance < 0.5), return white, else return black
    If luminance < 0.5 Then
        GetContrastingTextColor = RGB(255, 255, 255) ' White
    Else
        GetContrastingTextColor = RGB(0, 0, 0) ' Black
    End If
End Function

' Function to parse worker text and extract worker counts
Sub ParseWorkerText(ByVal workerText As String, ByRef bmlCount As Integer, ByRef lpCount As Integer)
    ' Initialize counts
    bmlCount = 0
    lpCount = 0
    
    ' Remove any extra spaces
    workerText = Trim(workerText)
    
    ' Process BML count
    Dim bmlPos As Integer
    bmlPos = InStr(1, workerText, "BML", vbTextCompare)
    
    If bmlPos > 1 Then
        ' Extract number before "BML"
        Dim bmlNumStr As String, i As Integer
        bmlNumStr = ""
        
        ' Go backwards from BML position to find the number
        For i = bmlPos - 1 To 1 Step -1
            If IsNumeric(Mid(workerText, i, 1)) Then
                bmlNumStr = Mid(workerText, i, 1) & bmlNumStr
            Else
                If i < bmlPos - 1 Then
                    Exit For ' Found a non-numeric character after finding at least one digit
                End If
            End If
        Next i
        
        ' Convert to number if we found something
        If Len(bmlNumStr) > 0 Then
            bmlCount = CInt(bmlNumStr)
        End If
    End If
    
    ' Process LP count
    Dim lpPos As Integer
    lpPos = InStr(1, workerText, "LP", vbTextCompare)
    
    If lpPos > 1 Then
        ' Extract number before "LP"
        Dim lpNumStr As String
        lpNumStr = ""
        
        ' Go backwards from LP position to find the number
        For i = lpPos - 1 To 1 Step -1
            If IsNumeric(Mid(workerText, i, 1)) Then
                lpNumStr = Mid(workerText, i, 1) & lpNumStr
            Else
                If i < lpPos - 1 Then
                    Exit For ' Found a non-numeric character after finding at least one digit
                End If
            End If
        Next i
        
        ' Convert to number if we found something
        If Len(lpNumStr) > 0 Then
            lpCount = CInt(lpNumStr)
        End If
    End If
End Sub
