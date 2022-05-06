Attribute VB_Name = "SvnUDF"
'@Folder("Sovereign")
'@IgnoreModule AssignmentNotUsed
Option Explicit

Public Function inchesMM(ByVal inputVal As Variant) As Variant
    On Error GoTo UDFErr
    inchesMM = WorksheetFunction.Convert(inputVal, "in", "mm")
    Exit Function
UDFErr:
    inchesMM = CVErr(xlErrValue)
End Function

Public Function finMM(ByVal footIn As Variant, ByVal inchIn As Variant) As Variant
    On Error GoTo UDFErr
    finMM = WorksheetFunction.Convert(footIn, "ft", "mm") + WorksheetFunction.Convert(inchIn, "in", "mm")
    Exit Function
UDFErr:
    finMM = CVErr(xlErrValue)
End Function

Public Function TwoDigitCount(ByVal Start As Long, ByVal count As Long) As String
    Dim i As Long
    For i = Start To Start + count - 1
        TwoDigitCount = TwoDigitCount + IIf(TwoDigitCount = vbNullString, vbNullString, ",") & Format(i, "00")
    Next
End Function

Public Function convertUnits(ByVal inValue As Variant, ByVal inUnit As String, ByVal outUnit As String) As Variant
    ' Extends CONVERT() to allow for extra unit types.
    On Error Resume Next
    convertUnits = WorksheetFunction.Convert(inValue, inUnit, outUnit)
    On Error GoTo 0
    If IsEmpty(convertUnits) Then convertUnits = splitSlashConvert(inValue, inUnit, outUnit)
    If IsEmpty(convertUnits) Then convertUnits = splitBaseConvert(inValue, inUnit, outUnit)
    If IsEmpty(convertUnits) Then convertUnits = CVErr(xlErrValue)
End Function

Public Function ReplaceBracketed(ByVal inString As String, ByVal inContent As String, ByVal replaceContent As String) As String
    Dim char As Long
    Dim currChar As String
    Dim outString As String
    Dim EarliestLeftBracket As Long
    Dim BracketCount As Long
    For char = 1 To Len(inString)
        currChar = Mid(inString, char, 1)
        
        If currChar = "(" Then
            If BracketCount = 0 Then EarliestLeftBracket = char
            BracketCount = BracketCount + 1
        ElseIf currChar = ")" Then
            BracketCount = BracketCount - 1
            If BracketCount = 0 Then
                outString = outString + replace(Mid(inString, EarliestLeftBracket, char - EarliestLeftBracket), inContent, replaceContent)
            End If
        End If
        
        If BracketCount = 0 Then
            outString = outString + currChar
        End If
    Next
    
    ReplaceBracketed = outString
End Function

Private Function splitBaseConvert(ByVal inValue As Variant, ByVal inUnit As String, ByVal outUnit As String) As Variant
    ' Extends convert to accept any exponent
    Dim splitInUnit() As String
    Dim splitOutUnit() As String
    Dim baseConvertRate As Double
    
    splitInUnit = Split(inUnit, "^")
    splitOutUnit = Split(outUnit, "^")
    
    If (UBound(splitInUnit) = 1 And UBound(splitOutUnit) = 1) Then ' Only one exponent
        If (splitInUnit(1) = splitOutUnit(1)) Then ' Same Exponent
        
            ' (E.g. mm^4 to m^4)
            On Error Resume Next
            baseConvertRate = WorksheetFunction.Convert(1, splitInUnit(0), splitOutUnit(0))
            splitBaseConvert = inValue * (baseConvertRate ^ CInt(splitInUnit(1)))
            On Error GoTo 0
            Exit Function
            
        End If
        
    End If
End Function

Private Function splitSlashConvert(ByVal inValue As Variant, ByVal inUnit As String, ByVal outUnit As String) As Variant
    ' Extends convert to accept converting divisor units with any exponent
    Dim splitInUnit() As String
    Dim splitOutUnit() As String
    Dim divisorConvertRate As Double
    
    splitInUnit = Split(inUnit, "/")
    splitOutUnit = Split(outUnit, "/")
    
    If (UBound(splitInUnit) = 1 And UBound(splitOutUnit) = 1) Then ' Only one divisor
        If (splitInUnit(0) = splitOutUnit(0)) Then ' Same Base
        
            ' (E.g. KN/m to KN/mm)
            On Error GoTo SplitBase
            divisorConvertRate = WorksheetFunction.Convert(1, splitInUnit(1), splitOutUnit(1))
            splitSlashConvert = inValue / (divisorConvertRate)
            On Error GoTo 0
            Exit Function
            
SplitBase:                                       ' (E.g. KN/m^2 to KN/mm^2)
            On Error Resume Next
            divisorConvertRate = splitBaseConvert(1, splitInUnit(1), splitOutUnit(1))
            splitSlashConvert = inValue / (divisorConvertRate)
            On Error GoTo 0
            Exit Function
            
        End If
    End If
End Function

Public Function Lerp(ByVal xGoal As Variant, ByVal xStart As Variant, ByVal yStart As Variant, ByVal xEnd As Variant, ByVal yEnd As Variant) As Variant
    Lerp = yStart + (xGoal - xStart) * (yEnd - yStart) / (xEnd - xStart)
End Function

Public Function getWord(ByVal inString As Variant, ByRef Del As String, ByRef Index As Variant) As Variant
    On Error GoTo UDFErr
    getWord = Split(inString, Del)(Index)
Finish:
    Exit Function
UDFErr:
    getWord = CVErr(xlErrNA)
    Resume Finish
End Function

Public Function getWordSafe(ByVal inString As Variant, ByVal Del As String, ByVal Index As Long) As Variant
    Dim newIndex As Long
    newIndex = Index
    On Error GoTo UDFErr
    Dim arr() As String
    arr = Split(inString, Del)
    If Index = -1 Then newIndex = UBound(arr)
    getWordSafe = arr(newIndex)
Finish:
    Exit Function
UDFErr:
    getWordSafe = inString
    Resume Finish
End Function

Public Function getWords(ByVal inString As Variant, ByVal Del As String) As Collection
    Set getWords = New Collection
    On Error GoTo UDFErr
    Dim arr() As String
    arr = Split(inString, Del)
    On Error GoTo 0
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        getWords.Add arr(i)
    Next
Finish:
    Exit Function
UDFErr:
    getWords.Add inString
    Resume Finish
End Function

Public Function selectValue(ByVal value As Variant, ByVal inputs As Range, ByVal outputs As Range) As Variant
    If (outputs.Areas.count <> inputs.Areas.count) Then
        selectValue = "Invalid Ranges"
        Exit Function
    End If
    
    Dim refIndex As Long
    If (outputs.Areas.count = 1) Then
        For refIndex = 1 To outputs.Cells.count
            If (value = inputs.Cells(refIndex).Value2) Then
                Set selectValue = outputs(refIndex)
                Exit Function
            End If
        Next
    Else
        For refIndex = 1 To outputs.Areas.count
            If (value = inputs.Areas(refIndex).Value2) Then
                Set selectValue = outputs.Areas(refIndex)
                Exit Function
            End If
        Next
    End If
End Function

Public Function usedRows(ByVal inRange As Range, Optional ByVal Offset As Long = 0) As Range
    Set usedRows = inRange.Resize(inRange.Worksheet.usedRange.Rows.count - Offset).Offset(Offset)
End Function

Public Function usedColumns(ByVal inRange As Range, Optional ByVal Offset As Long = 0) As Range
    Set usedColumns = inRange.Resize(inRange.Rows.count, inRange.Worksheet.usedRange.Rows.count - Offset).Offset(0, Offset)
End Function

Public Function selectFromSet(ByRef Values As Variant, ByVal inputs As Range, ByVal outputs As Range) As Variant
    On Error GoTo UDFErr
    If (IsArray(Values)) Then
        Dim pri As Variant
        For Each pri In Values
            Set selectFromSet = selectValue(pri, inputs, outputs)
            If (selectFromSet <> Empty) Then
                Exit Function
            End If
        Next
    Else
        Set selectFromSet = selectValue(Values, inputs, outputs)
        If (selectFromSet <> Empty) Then
            Exit Function
        End If
    End If
    selectFromSet = CVErr(xlErrNA)
    Exit Function
UDFErr:
    selectFromSet = CVErr(xlErrValue)
End Function

Public Function selectLeast(ByVal inputs As Range, ByVal outputs As Range) As Variant
    On Error GoTo UDFErr
    Set selectLeast = selectFromSet(WorksheetFunction.Min(inputs), inputs, outputs)
    Exit Function
UDFErr:
    selectLeast = CVErr(xlErrValue)
End Function

Public Function selectMost(ByVal inputs As Range, ByVal outputs As Range) As Variant
    On Error GoTo UDFErr
    Set selectMost = selectFromSet(WorksheetFunction.Max(inputs), inputs, outputs)
    Exit Function
UDFErr:
    selectMost = CVErr(xlErrValue)
End Function

Public Function arrayMaxIFS(ParamArray valueOutput() As Variant) As Variant
    arrayMaxIFS = arrayMod(valueOutput, 2, 1)(arrayGetMaxIndex(arrayMod(valueOutput, 2, 0)))
End Function

Public Function getWeekdayDistribution(ByVal inDateRange As Range, ByVal inRowRange As Range, ByVal inFilterRange As Range, ByVal outDateRange As Range, ByVal outRowRange As Range, ByVal outFilterGreater As Variant) As Variant
    ' Set default Values
    Dim outArray() As Variant
    ReDim outArray(1 To 1, 1 To 1)
    outArray(1, 1) = CVErr(xlErrNA)
    
    ' Dynamic resize
    Dim used As Long
    used = inDateRange.Worksheet.usedRange.Rows.count
    ReDim outArray(1 To outRowRange.Rows.count, 1 To outDateRange.Columns.count - 1)
    
    ' Retrieve Variant Arrays
    Dim inDates As Variant
    Dim inRows As Variant
    Dim inFilter As Variant
    Dim outDates As Variant
    Dim outRows As Variant
    inDates = inDateRange.Resize(used - 1).Offset(1).Value2
    inRows = inRowRange.Resize(used - 1).Offset(1).Value2
    inFilter = inFilterRange.Resize(used - 1).Offset(1).Value2
    outDates = outDateRange.Value2
    outRows = outRowRange.Value2
    
    ' Optional - Key-Cache Dates and Rows for performance
    
    ' Process every input row
    Dim i As Long
    Dim outRowIndex As Long
    Dim outColIndex As Long
    For i = LBound(inDates, 1) To UBound(inDates, 1)
        ' Find output row
        For outRowIndex = LBound(outRows, 1) To UBound(outRows, 1)
            If inRows(i, 1) = outRows(outRowIndex, 1) Then Exit For
            If outRowIndex = UBound(outRows, 1) Then GoTo NextRow
        Next
        For outColIndex = LBound(outDates, 2) To UBound(outDates, 2) - 1
            If inDates(i, 1) >= outDates(1, outColIndex) And inDates(i, 1) < outDates(1, outColIndex + 1) Then Exit For
            If outColIndex = UBound(outDates, 2) - 1 Then GoTo NextRow
        Next
        
        outArray(outRowIndex, outColIndex) = CLng(outArray(outRowIndex, outColIndex)) + (IIf(inFilter(i, 1) >= outFilterGreater, 3, 1) * 10 ^ (6 - (Weekday(inDates(i, 1), vbMonday) - 1)))
NextRow:
    Next
    
    
    getWeekdayDistribution = Application.Transpose(Application.Transpose(outArray))
End Function

Public Function fatigueSeek(ByVal goal As Double, ByVal row As String, ByVal minCol As Long, ByVal maxCol As Long, ByVal startCol As Long, ByVal LookupAreas As Range) As Long
    Dim attempt As Double
    Dim endCol As Long
    Dim maxIter As Long
    maxIter = 200
    Dim rangeSize As Long
    For rangeSize = 0 To maxIter
        endCol = startCol + rangeSize
        ' rangeStartEnd Sum with a static approximate on either side
        attempt = WorksheetFunction.Max(0, minCol - startCol) * rangeLookup(minCol, row, LookupAreas) _
      + WorksheetFunction.Sum(rangeStartEnd(WorksheetFunction.Median(minCol, startCol, maxCol), WorksheetFunction.Median(minCol, endCol, maxCol), row, row, LookupAreas)) _
      + WorksheetFunction.Max(0, endCol - maxCol) * rangeLookup(maxCol, row, LookupAreas)
        If attempt > goal Then Exit For
    Next
    fatigueSeek = endCol - 1                     ' Remaining years can't include partial years
End Function

Public Function rangeLookup(ByVal colVal As Variant, ByVal rowVal As Variant, ByVal LookupAreas As Range, Optional ByVal rowMatchType As Long = 0) As Variant
    ' Finds a cell on a sheet based on column and row values using a formatted range
    ' lookupAreas Area 1 contains the column of rowVals (keys/IDs) to find the row
    ' lookupAreas Area 2 contains the row of colVals (headers) to find the column
    
    Dim foundCol As Range
    Dim foundRow As Range
    Dim outRange As Range
    
    ' Find Row and Column using Index-Match (Fastest for single-instance searches)
    On Error Resume Next                         ' Disable Error Checking for WorksheetFunctions
    Set foundRow = WorksheetFunction.Index(LookupAreas.Areas(1), WorksheetFunction.Match(rowVal, LookupAreas.Areas(1), rowMatchType))
    Set foundCol = WorksheetFunction.Index(LookupAreas.Areas(2), WorksheetFunction.Match(colVal, LookupAreas.Areas(2), 0))
    On Error GoTo 0                              'Enable Error Checking

    If (foundRow Is Nothing Or foundCol Is Nothing) Then
        rangeLookup = CVErr(xlErrNA)
    Else
        Set outRange = LookupAreas.Worksheet.Cells(foundRow.row, foundCol.column)
        If Not (IsError(outRange)) Then          ' Errors fail empty-check
            If (IsEmpty(outRange)) Then          ' Flag blank cells (Don't treat them as 0)
                rangeLookup = CVErr(xlErrNA)     ' Short circuit AND avoids failing on error
                Exit Function
            End If
        End If
        Set rangeLookup = outRange
    End If
    
End Function

Public Function rangeMultiLookup(ByVal colVal As Variant, ByVal rowValues As Variant, ByVal LookupAreas As Range) As Variant
    ' Finds a cell on a sheet based on column and row values using a formatted range
    ' lookupAreas Area 1 contains the column of rowVals (keys/IDs) to find the row
    ' lookupAreas Area 2 contains the row of colVals (headers) to find the column
    
    Dim foundCol As Range
    Dim foundRow As Range
    Dim outRange As Range
    
    ' Find Row and Column using Index-Match (Fastest for single-instance searches)
    On Error Resume Next                         ' Disable Error Checking for WorksheetFunctions
    Set foundCol = WorksheetFunction.Index(LookupAreas.Areas(2), WorksheetFunction.Match(colVal, LookupAreas.Areas(2), 0))
    On Error GoTo 0                              'Enable Error Checking
    
    Dim RowArray As Variant
    Dim j As Long
    Dim n As Long
    RowArray = LookupAreas.Areas(1).Value2
    For j = LBound(RowArray, 1) To UBound(RowArray, 1) ' Rowwise
        For n = LBound(RowArray, 2) To UBound(RowArray, 2) ' Columnwise
            If RowArray(j, n) <> rowValues(n) Then Exit For
            If n = UBound(RowArray, 2) Then Set foundRow = LookupAreas.Areas(1).Rows(j)
        Next
        If Not foundRow Is Nothing Then Exit For
    Next j

    If (foundRow Is Nothing Or foundCol Is Nothing) Then
        rangeMultiLookup = CVErr(xlErrNA)
    Else
        Set outRange = LookupAreas.Worksheet.Cells(foundRow.row, foundCol.column)
        If Not (IsError(outRange)) Then          ' Errors fail empty-check
            If (IsEmpty(outRange)) Then          ' Flag blank cells (Don't treat them as 0)
                rangeMultiLookup = CVErr(xlErrNA) ' Short circuit AND avoids failing on error
                Exit Function
            End If
        End If
        Set rangeMultiLookup = outRange
    End If
    
End Function

Public Function indexMatch(ByVal findVal As Variant, ByVal inputRange As Range, ByVal outputRange As Range) As Variant
    ' UDF Variant of index match, with simplified parameters.
    Dim foundCell As Range

    On Error Resume Next                         ' Disable Error Checking for WorksheetFunctions
    Set foundCell = WorksheetFunction.Index(outputRange, WorksheetFunction.Match(findVal, inputRange, 0))
    On Error GoTo 0                              'Enable Error Checking
    
    If (foundCell Is Nothing) Then
        indexMatch = CVErr(xlErrNA)
    Else
        Set indexMatch = foundCell
    End If
End Function

Public Function closestMatch(ByVal findVal As Variant, ByVal inputRange As Range, ByVal outputRange As Range, Optional ByVal variance As Double = -1) As Variant
    Dim bestIndex As Long
    bestIndex = -1
    Dim bestVariance As Double
    bestVariance = variance
    Dim Index As Long
    For Index = 0 To inputRange.count
        Dim difference As Double
        difference = Abs(inputRange.Rows(Index).Value2 - findVal)
        If bestVariance = -1 Or difference < bestVariance Then
            bestVariance = difference
            bestIndex = Index
        End If
    Next
    On Error Resume Next
    If bestIndex <> -1 And bestVariance <> -1 Then Set closestMatch = WorksheetFunction.Index(outputRange, bestIndex)
    If IsEmpty(closestMatch) Then Set closestMatch = CVErr(xlErrNA)
    On Error GoTo 0
End Function

Public Function rangeStartEnd(ByVal colValStart As Variant, ByVal colVaEnd As Variant, ByVal rowValStart As Variant, ByVal rowValEnd As Variant, ByVal LookupAreas As Range) As Variant
    Dim startCell As Range
    Dim endCell As Range
    
    On Error GoTo startError
    Set startCell = rangeLookup(colValStart, rowValStart, LookupAreas)
    On Error GoTo endError
    Set endCell = rangeLookup(colVaEnd, rowValEnd, LookupAreas)
    On Error GoTo 0
    
    rangeStartEnd = startCell.Worksheet.Range(startCell, endCell).Value2
        
    ' Error Handling
    If (False) Then
startError:
        rangeStartEnd = rangeLookup(colValStart, rowValStart, LookupAreas)
    End If
    If (False) Then
endError:
        rangeStartEnd = rangeLookup(colVaEnd, rowValEnd, LookupAreas)
    End If
End Function

Public Function getWordsMulti(ByVal inString As Variant, ParamArray Dels()) As Variant
    Dim currentDel As String
    Dim i As Long
    Dim currentSplit As Variant
    Dim outCollection As Collection
    Dim newOutCollection As Collection
    Set newOutCollection = New Collection
    newOutCollection.Add inString
    For i = LBound(Dels) To UBound(Dels)
        currentDel = Dels(i)
        Set outCollection = newOutCollection
        Set newOutCollection = New Collection
        For Each currentSplit In outCollection
            Set newOutCollection = collectionAppend(newOutCollection, getWords(currentSplit, currentDel))
        Next
    Next
    getWordsMulti = collectionToArray(newOutCollection)
End Function


Public Function MatchInstance(ByVal outColumn As Range, ByVal inColumn As Range, ByVal value As Variant, ByVal inInstance As Long) As Variant
    Dim used As Long
    used = inColumn.Worksheet.usedRange.Rows.count
    Dim Instance As Long
    Instance = inInstance
    
    Dim inValues As Variant
    inValues = inColumn.Resize(used - 1).Offset(1).Value2
    Dim outValues As Variant
    outValues = outColumn.Resize(used - 1).Offset(1).Value2
    
    Dim i As Long
    For i = LBound(inValues) To UBound(inValues)
        If inValues(i, 1) = value Then
            Instance = Instance - 1
            If Instance = 0 Then
                MatchInstance = outValues(i, 1)
                Exit Function
            End If
        End If
    Next
    MatchInstance = CVErr(xlErrNA)
End Function

Public Function GetInstances(ByVal outColumn As Range, ByVal inColumn As Range, ByVal value As Variant, Optional ByVal inColumn2 As Range, Optional ByVal Value2 As Variant, Optional ByVal notValue As Boolean = False) As Variant
    Dim outArray() As Variant
    ReDim outArray(1 To 1, 1 To 1)
    outArray(1, 1) = CVErr(xlErrNA)
    
    Dim used As Long
    used = inColumn.Worksheet.usedRange.Rows.count

    Dim inValues As Variant
    inValues = inColumn.Resize(used - 1).Offset(1).Value2
    Dim inValues2 As Variant
    If Not inColumn2 Is Nothing Then inValues2 = inColumn2.Resize(used - 1).Offset(1).Value2
    Dim outValues As Variant
    outValues = outColumn.Resize(used - 1).Offset(1).Value2
    
    Dim Instance As Long
    Instance = 1
    
    Dim i As Long
    For i = LBound(inValues) To UBound(inValues)
        If inValues(i, 1) = value Then
            If Not inColumn2 Is Nothing Then
                If IsMissing(Value2) Then If inValues2(i, 1) = vbNullString Then GoTo NextRow
                If Not IsMissing(Value2) Then If IIf(inValues2(i, 1) <> Value2, Not notValue, notValue) Then GoTo NextRow
            End If
            ReDim Preserve outArray(LBound(outArray, 1) To UBound(outArray, 1), LBound(outArray, 2) To Instance)
            outArray(1, Instance) = outValues(i, 1)
            Instance = Instance + 1
        End If
NextRow:
    Next
    GetInstances = Application.Transpose(outArray)
End Function

Public Function ArrayFormulaCount() As Variant
    Dim outArray() As Variant
    ReDim outArray(1 To Application.Caller.Rows.count) As Variant
    
    Dim i As Long
    For i = LBound(outArray) To UBound(outArray)
        outArray(i) = i
    Next
    
    ArrayFormulaCount = outArray
End Function

Public Function rSelect(ParamArray var() As Variant) As Range
    Set rSelect = var(var(0))
End Function

Public Function NamedRange(ByVal rangeName As String) As Variant
    ' non-volatile shortcut for looking up a named range - useful for data sheets that don't update.
    On Error GoTo UDFErr
    Set NamedRange = ThisWorkbook.Names(rangeName).RefersToRange
    Exit Function
UDFErr:
    NamedRange = CVErr(xlErrValue)
End Function

Public Function stringCount(ByVal inputString As String, ByVal stringToFind As String) As Variant
    stringCount = (Len(inputString) - Len(replace(inputString, stringToFind, vbNullString))) / Len(stringToFind)
End Function

Public Function StringMatches(ByVal substring As String, ByVal Lookup As Range) As Variant
    StringMatches = arrayGetEquals(substring, Lookup.Value2, True)
End Function

Public Function FilterRange(ByVal inRange As Variant, ByVal Filter As Variant) As Variant
    FilterRange = arrayGetFiltered(arrayForce1D(inRange.Value2), Filter)
End Function


