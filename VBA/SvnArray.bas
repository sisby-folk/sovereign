Attribute VB_Name = "SvnArray"
'@Folder("Sovereign")
Option Explicit
Option Compare Text

' ALL STRING COMPARISONS ARE CASE INSENSITIVE

Public Function arrayMatch(ByVal inRange1 As Variant, ByVal inRange2 As Variant, Optional ByVal row2 As Long = -1) As Variant
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    
    If IsEmpty(inRange1) Or IsEmpty(inRange2) Then Exit Function
    
    If row2 <> -1 Then
        For i = 1 To UBound(inRange1, 1)
            For j = 1 To UBound(inRange1, 2)
                For l = 1 To UBound(inRange2, 2)
                    If inRange1(i, j) = inRange2(row2, l) Then
                        arrayMatch = inRange1(i, j)
                        Exit Function
                    End If
                Next
            Next
        Next
    Else
        For i = 1 To UBound(inRange1, 1)
            For j = 1 To UBound(inRange1, 2)
                For k = 1 To UBound(inRange2, 1)
                    For l = 1 To UBound(inRange2, 2)
                        If inRange1(i, j) = inRange2(k, l) Then
                            arrayMatch = inRange1(i, j)
                            Exit Function
                        End If
                    Next
                Next
            Next
        Next
    End If
End Function

Public Function arrayMod(ByVal inArray As Variant, ByVal inMod As Long, ByVal offset As Long) As Variant
    Dim outArray() As Variant
    ReDim outArray(LBound(inArray) To LBound(inArray) + ((UBound(inArray) - LBound(inArray) - offset) \ inMod))
    Dim i As Long
    For i = LBound(outArray) To UBound(outArray)
        outArray(i) = inArray(LBound(outArray) + (i - LBound(outArray)) * inMod + offset)
    Next
    arrayMod = outArray
End Function

Public Function arrayReverse(ByVal inArray As Variant) As Variant
    Dim outArray() As Variant
    ReDim outArray(LBound(inArray) To UBound(inArray))
    Dim i As Long
    For i = 0 To (UBound(inArray) - LBound(inArray))
        outArray(UBound(inArray) - i) = inArray(LBound(inArray) + i)
    Next
    arrayReverse = outArray
End Function

Public Function arrayGetMatch(ByVal inValue As Variant, ByVal inRange As Variant, Optional ByVal column As Long = -1, Optional ByVal row As Long = -1) As Variant
    Dim i As Long
    
    If column <> -1 Then
        For i = LBound(inRange, 1) To UBound(inRange, 1)
            If inValue = inRange(i, column) Then
                arrayGetMatch = i
                Exit For
            End If
        Next
    ElseIf row <> -1 Then
        For i = LBound(inRange, 2) To UBound(inRange, 2)
            If inValue = inRange(row, i) Then
                arrayGetMatch = i
                Exit For
            End If
        Next
    Else
        For i = LBound(inRange) To UBound(inRange)
            If inValue = inRange(i) Then
                arrayGetMatch = i
                Exit For
            End If
        Next
    End If
End Function

Public Function arrayFromParams(ByVal inArray As Variant) As Long
    
End Function

Public Function arrayGetMaxIndex(ByVal inArray As Variant) As Long
    Dim maxValue As Variant
    maxValue = WorksheetFunction.Max(inArray)
    For arrayGetMaxIndex = LBound(inArray) To UBound(inArray)
        If inArray(arrayGetMaxIndex) = maxValue Then Exit Function
    Next
    arrayGetMaxIndex = -1
End Function

Public Function arrayGetIndex(ByVal inArray As Variant, ByVal i As Variant, Optional ByVal j As Variant) As Variant
    If IsEmpty(i) Then Exit Function
    Select Case arrayGetDimension(inArray)
        Case 1
            arrayGetIndex = inArray(i)
        Case 2
            If IsEmpty(j) Then Exit Function
            arrayGetIndex = inArray(i, j)
    End Select
End Function

Public Function arrayConcat(ByVal array1 As Variant, ByVal array2 As Variant, ByVal Del As String) As Variant
    arrayConcat = arrayCheckEmpty(array1, array2)
    If Not IsEmpty(arrayConcat) Then Exit Function
    Dim outArray() As Variant
    
    Dim i As Long
    Dim j As Long
    If arrayGetDimension(array1) = 1 Then
        ReDim outArray(LBound(array1) To UBound(array1))

        For i = LBound(array1) To UBound(array1)
            outArray(i) = array1(i) & Del & array2(i)
        Next
    Else
        ReDim outArray(LBound(array1, 1) To UBound(array1, 1), LBound(array1, 2) To UBound(array1, 2))
        For i = LBound(array1, 1) To UBound(array1, 1)
            For j = LBound(array1, 2) To UBound(array1, 2)
                outArray(i, j) = array1(i, j) & Del & array2(i, j)
            Next
        Next
    End If
    arrayConcat = outArray
End Function

''' Array Tools '''

'' Unary Manipulation ''

Public Function arrayGetDimension(ByVal inArray As Variant) As Long
    Dim Test As Long
    On Error Resume Next
    Do
        arrayGetDimension = arrayGetDimension + 1
        Test = UBound(inArray, arrayGetDimension)
        Pass var:=Test
    Loop Until Err.Number <> 0
    Err.Clear
    On Error GoTo 0
    arrayGetDimension = arrayGetDimension - 1
End Function

Public Function arrayFromMirror(ByVal value As Variant, ByVal matchArray As Variant) As Variant
    If IsArray(value) Then
        If UBound(value) - LBound(value) <> 0 Then
            arrayFromMirror = arrayForce1D(value)
            Exit Function
        End If
    End If
    Dim outArray() As Variant
    ReDim outArray(LBound(matchArray) To UBound(matchArray))
    Dim i As Long
    For i = LBound(outArray) To UBound(outArray)
        If IsArray(value) Then outArray(i) = value(LBound(value))
        If Not IsArray(value) Then outArray(i) = value
    Next
    arrayFromMirror = outArray
End Function

Public Function arrayForce1D(ByVal inArray As Variant) As Variant
    Dim outArray() As Variant
    If Not IsArray(inArray) Then
        ReDim outArray(1 To 1) As Variant
        outArray(1) = inArray
        arrayForce1D = outArray
    ElseIf arrayGetDimension(inArray) = 2 Then
        ReDim outArray(LBound(inArray, 1) To UBound(inArray, 1))
        Dim i As Long
        For i = LBound(inArray, 1) To UBound(inArray, 1)
            outArray(i) = inArray(i, 1)
        Next
        arrayForce1D = outArray
    Else
        arrayForce1D = inArray
    End If
End Function

'@EntryPoint
Public Function arrayOffset(ByVal inArray As Variant, ByVal offsetR As Long, Optional ByVal offsetC As Long = 0, Optional ByVal replace As Variant, Optional ByVal spaceR As Long = 1, Optional ByVal spaceC As Long = 1) As Variant
    Dim outArray() As Variant
    Dim i As Long
    Dim j As Long
    inArray = arrayForce2D(inArray)
    ReDim outArray(LBound(inArray, 1) To LBound(inArray, 1) + (UBound(inArray, 1) - LBound(inArray, 1)) * spaceR + offsetR, LBound(inArray, 2) To LBound(inArray, 2) + (UBound(inArray, 2) - LBound(inArray, 2)) * spaceC + offsetC)
    If Not IsMissing(replace) Then
        For i = LBound(outArray, 1) To UBound(outArray, 1)
            For j = LBound(outArray, 2) To UBound(outArray, 2)
                outArray(i, j) = replace
            Next
        Next
    End If
    For i = LBound(inArray, 1) To UBound(inArray, 1)
        For j = LBound(inArray, 2) To UBound(inArray, 2)
            outArray(LBound(inArray, 1) + (i - LBound(inArray, 1)) * spaceR + offsetR, LBound(inArray, 2) + (j - LBound(inArray, 2)) * spaceC + offsetC) = inArray(i, j)
        Next
    Next
    arrayOffset = outArray
End Function

Public Function arrayForce2D(ByVal inArray As Variant) As Variant
    Dim outArray() As Variant
    If Not IsArray(inArray) Then
        ReDim outArray(1 To 1, 1 To 1) As Variant
        outArray(1, 1) = inArray
        arrayForce2D = outArray
    ElseIf arrayGetDimension(inArray) = 1 Then
        ReDim outArray(LBound(inArray) To UBound(inArray), 1 To 1)
        Dim i As Long
        For i = LBound(inArray) To UBound(inArray)
            outArray(i, 1) = inArray(i)
        Next
        arrayForce2D = outArray
    Else
        arrayForce2D = inArray
    End If
End Function

Public Function arrayGetReplace(ByVal inArray As Variant, ByVal toReplace As Variant, ByVal replaceWith As Variant, Optional ByVal AllLongs As Boolean = False) As Variant
    If IsEmpty(inArray) Then
        arrayGetReplace = Empty
        Exit Function
    End If
    Dim valArray As Variant
    valArray = arrayForce2D(inArray)
    Dim currentArrayElement As Variant
    For Each currentArrayElement In valArray
        If AllLongs And IsNumeric(currentArrayElement) Then currentArrayElement = CLng(currentArrayElement)
        If AllLongs And Not IsNumeric(currentArrayElement) Then currentArrayElement = replaceWith
        If currentArrayElement = toReplace Then currentArrayElement = replaceWith
    Next
    arrayGetReplace = valArray
End Function

Public Function arrayGetCount(ByVal inArray As Variant) As Long
    If IsEmpty(inArray) Then
        arrayGetCount = 0
        Exit Function
    End If
    If arrayGetDimension(inArray) = 1 Then arrayGetCount = UBound(inArray) - LBound(inArray) + 1
    If arrayGetDimension(inArray) = 2 Then arrayGetCount = (UBound(inArray, 1) - LBound(inArray, 1) + 1) * (UBound(inArray, 2) - LBound(inArray, 2) + 1) ' XXX: Counts BOTH ways. may be unwanted
End Function

Public Function arrayGetJoin(ByVal inArray As Variant, ByVal Del As String) As Variant
    If IsEmpty(inArray) Then
        arrayGetJoin = vbNullString
        Exit Function
    End If
    Dim array1 As Variant
    array1 = arrayForce2D(inArray)
    Dim i As Long
    Dim j As Long
    For i = LBound(array1, 1) To UBound(array1, 1)
        For j = LBound(array1, 2) To UBound(array1, 2)
            If Not IsEmpty(array1(i, j)) Then arrayGetJoin = arrayGetJoin & IIf(IsEmpty(arrayGetJoin), vbNullString, Del) & CStr(array1(i, j))
        Next
    Next
End Function

Public Function arrayGetSum(ByVal inArray As Variant) As Long
    If IsEmpty(inArray) Then
        arrayGetSum = 0
        Exit Function
    End If
    arrayGetSum = WorksheetFunction.Sum(inArray)
End Function

'' Binary Operations ''

Public Function arrayCheckEmpty(ByVal array1 As Variant, ByVal array2 As Variant) As Variant
    If Not IsArray(array1) Then
        arrayCheckEmpty = array2
    ElseIf UBound(array1) < LBound(array1) Then
        arrayCheckEmpty = array2
    ElseIf Not IsArray(array2) Then
        arrayCheckEmpty = array1
    ElseIf UBound(array2) < LBound(array2) Then
        arrayCheckEmpty = array1
    End If
End Function

Public Function arrayAppendRow(ByVal inArray1 As Variant, ByVal array2 As Variant) As Variant
    arrayAppendRow = arrayCheckEmpty(inArray1, array2)
    If Not IsEmpty(arrayAppendRow) Then Exit Function
    
    Dim array1 As Variant
    array1 = arrayForce2D(inArray1)
    
    Dim outArray() As Variant
    ReDim outArray(LBound(array1, 1) To UBound(array1, 1), LBound(array1, 2) To UBound(array1, 2) + 1)
    
    Dim i As Long
    Dim j As Long
    For i = LBound(array1, 1) To UBound(array1, 1)
        For j = LBound(array1, 2) To UBound(array1, 2)
            outArray(i, j) = array1(i, j)
        Next
    Next
    
    For i = LBound(array2) To UBound(array2)
        outArray(i, UBound(outArray, 2)) = array2(i)
    Next
    arrayAppendRow = outArray
End Function

Public Function arrayGetTransposed(ByVal inArray As Variant, Optional ByVal Dimension = 2) As Variant
    Dim outArray() As Variant
    Dim i As Long
    Dim j As Long
    Select Case arrayGetDimension(inArray)
        Case 1
            Select Case Dimension
                Case 1
                    ReDim outArray(LBound(inArray) To UBound(inArray), 1 To 1)
                    For i = LBound(inArray) To UBound(inArray)
                        outArray(i, 1) = inArray(i)
                    Next
                Case 2
                    ReDim outArray(1 To 1, LBound(inArray) To UBound(inArray))
                    For i = LBound(inArray) To UBound(inArray)
                        outArray(1, i) = inArray(i)
                    Next
            End Select
        Case 2
            Select Case Dimension
                Case 1
                    ReDim outArray(LBound(inArray, 2) To UBound(inArray, 2), LBound(inArray, 1) To UBound(inArray, 1))
                    For i = LBound(inArray, 1) To UBound(inArray, 1)
                        For j = LBound(inArray, 2) To UBound(inArray, 2)
                            outArray(j, i) = inArray(i, j)
                        Next
                    Next
                Case 2
                    ReDim outArray(LBound(inArray, 2) To UBound(inArray, 2), LBound(inArray, 1) To UBound(inArray, 1))
                    For i = LBound(inArray, 1) To UBound(inArray, 1)
                        For j = LBound(inArray, 2) To UBound(inArray, 2)
                            outArray(j, i) = inArray(i, j)
                        Next
                    Next
            End Select
    End Select
    arrayGetTransposed = outArray
End Function

Public Function arrayGetEquals(ByVal value As Variant, ByVal valueArray As Variant, Optional ByVal LRStringComp As Boolean = False, Optional ByVal onError As Variant = 0) As Variant
    Dim valArray As Variant
    valArray = arrayForce1D(valueArray)
    Dim compArray As Variant
    compArray = arrayFromMirror(value, valArray)
    arrayGetEquals = arrayEquals(compArray, valArray, LRStringComp, onError)
End Function

Public Function arrayEquals(ByVal array1 As Variant, ByVal array2 As Variant, Optional ByVal LRStringComp As Boolean = False, Optional ByVal onError As Variant = 0) As Variant

    Dim outArray() As Variant
    ReDim outArray(LBound(array1) To UBound(array1))
    
    Dim i As Long
    For i = LBound(array1) To UBound(array1)
        If IsError(array1(i)) Or IsError(array2(i)) Then
            outArray(i) = onError
        Else
            If LRStringComp Then outArray(i) = IIf(InStr(1, CStr(array2(i)), CStr(array1(i)), vbTextCompare) <> 0, 1, 0)
            If Not LRStringComp Then outArray(i) = IIf(array1(i) = array2(i), 1, 0)
        End If
    Next
    arrayEquals = outArray
End Function

Public Function arrayMult(ByVal array1 As Variant, ByVal array2 As Variant) As Variant
    arrayMult = arrayCheckEmpty(array1, array2)
    If Not IsEmpty(arrayMult) Then Exit Function
    
    Dim outArray() As Variant
    ReDim outArray(LBound(array1) To UBound(array1))
    Dim i As Long
    For i = LBound(array1) To UBound(array1)
        outArray(i) = array1(i) * array2(i)
    Next
    arrayMult = outArray
End Function
Public Function arrayBoolMult(ByVal array1 As Variant, ByVal array2 As Variant, ByVal Falsy As Variant) As Variant
    arrayBoolMult = arrayCheckEmpty(array1, array2)
    If Not IsEmpty(arrayBoolMult) Then Exit Function
    
    Dim outArray() As Variant
    ReDim outArray(LBound(array1) To UBound(array1))
    Dim i As Long
    For i = LBound(array1) To UBound(array1)
        outArray(i) = IIf(array2(i) = True, array1(i), Falsy)
    Next
    arrayBoolMult = outArray
End Function

Public Function arrayOr(ByVal array1 As Variant, ByVal array2 As Variant) As Variant
    arrayOr = arrayCheckEmpty(array1, array2)
    If Not IsEmpty(arrayOr) Then Exit Function
    
    Dim outArray() As Variant
    ReDim outArray(LBound(array1) To UBound(array1))
    Dim i As Long
    For i = LBound(array1) To UBound(array1)
        outArray(i) = IIf(array1(i) = 1 Or array2(i) = 1, 1, 0)
    Next
    arrayOr = outArray
End Function

Public Function arrayNon(ByVal inArray As Variant, ByVal nonValue As Variant, Optional ByVal blank As Variant = CVErr(xlErrNA)) As Variant
    arrayNon = blank
    Dim outArray() As Variant
    Dim count As Long
    Dim i As Long
    For i = LBound(inArray) To UBound(inArray)
        If inArray(i) <> nonValue Then count = count + 1
    Next
    If count = 0 Then Exit Function
    ReDim outArray(1 To count)
    count = 1
    For i = LBound(inArray) To UBound(inArray)
        If inArray(i) <> nonValue Then
            outArray(count) = inArray(i)
            count = count + 1
        End If
    Next
    arrayNon = outArray
End Function

Public Function arrayMode(ByVal inArray As Variant) As Variant
    On Error Resume Next
    arrayMode = WorksheetFunction.Mode_Singl(inArray)
    On Error GoTo 0
    If arrayGetCount(inArray) = 1 Or IsEmpty(arrayMode) Then
        arrayMode = inArray(LBound(inArray))
    End If
End Function

Public Function arrayAdd(ByVal array1 As Variant, ByVal array2 As Variant) As Variant
    arrayAdd = arrayCheckEmpty(array1, array2)
    If Not IsEmpty(arrayAdd) Then Exit Function
    
    Dim outArray() As Variant
    ReDim outArray(LBound(array1) To UBound(array1))
    Dim i As Long
    For i = LBound(array1) To UBound(array1)
        outArray(i) = IIf(array1(i) = vbNullString, 0, array1(i)) + IIf(array2(i) = vbNullString, 0, array2(i))
    Next
    arrayAdd = outArray
End Function

Public Function arrayGetFiltered(ByVal inArray As Variant, ByVal inFilter As Variant) As Variant
    If WorksheetFunction.Sum(inFilter) > 0 Then
        Dim valArray As Variant
        valArray = arrayForce2D(inArray)
        
        Dim outArray() As Variant
        ReDim outArray(1 To WorksheetFunction.Sum(inFilter), LBound(valArray, 2) To UBound(valArray, 2))
        
        Dim i As Long
        Dim j As Long
        Dim count As Long
        For i = LBound(valArray, 1) To UBound(valArray, 1)
            If inFilter(i) = 1 Then
                For j = LBound(valArray, 2) To UBound(valArray, 2)
                    outArray(LBound(outArray) + count, j) = valArray(i, j)
                Next
                count = count + 1
            End If
        Next
    
        arrayGetFiltered = outArray
    End If
End Function

Public Function arrayGetCopy(ByVal inArray As Variant) As Variant
    Dim outArray() As Variant
    ReDim outArray(LBound(inArray) To UBound(inArray))
    
    Dim i As Long
    For i = LBound(outArray) To UBound(outArray)
        outArray(i) = inArray(i)
    Next
    arrayGetCopy = outArray
End Function

Public Function arrayAppend(ByVal inArray1 As Variant, ByVal inArray2 As Variant) As Variant
    arrayAppend = arrayCheckEmpty(inArray1, inArray2)
    If Not IsEmpty(arrayAppend) Then Exit Function
    
    Dim array1 As Variant
    Dim array2 As Variant
    array1 = arrayForce2D(inArray1)
    array2 = arrayForce2D(inArray2)
    
    Dim outArray() As Variant
    ' 1 to 5, 1 to 5 = 1 To 1+5+5-1
    ' 1 to 1, 1 to 1 = 1 To 1+1+1-1
    ReDim outArray(LBound(array1) To (1 + UBound(array1, 1) + UBound(array2, 1) - LBound(array2, 1)), LBound(array1, 2) To UBound(array1, 2))
    
    Dim i As Long
    Dim j As Long
    For i = LBound(outArray, 1) To UBound(outArray, 1)
        For j = LBound(outArray, 2) To UBound(outArray, 2)
            If i <= UBound(array1) Then outArray(i, j) = array1(i, j)
            If i > UBound(array1) Then outArray(i, j) = array2(i - (UBound(array1) - LBound(array1) + 1), j)
        Next
    Next
    arrayAppend = outArray
End Function

Public Function arrayFromArgs(ParamArray Elements() As Variant) As Variant
    Dim outArray() As Variant
    ReDim outArray(1 To 1, 1 To (UBound(Elements) - LBound(Elements) + 1))
    Dim j As Long
    For j = 1 To UBound(outArray, 2)
        outArray(1, j) = Elements(LBound(Elements) + j - 1)
    Next
    arrayFromArgs = outArray
End Function
Public Function arrayFrom1DArgs(ParamArray Elements() As Variant) As Variant
    Dim outArray() As Variant
    ReDim outArray(1 To (UBound(Elements) - LBound(Elements) + 1))
    Dim j As Long
    For j = 1 To UBound(outArray)
        outArray(j) = Elements(LBound(Elements) + j - 1)
    Next
    arrayFrom1DArgs = outArray
End Function

Public Function arrayContainsRow(ByVal inArray As Variant, ByVal row As Variant) As Boolean
    If IsEmpty(inArray) Then
        arrayContainsRow = False
        Exit Function
    End If
    
    Dim i As Long
    Dim j As Long
    For i = LBound(inArray, 1) To UBound(inArray, 1)
        arrayContainsRow = True
        For j = LBound(inArray, 2) To UBound(inArray, 2)
            arrayContainsRow = arrayContainsRow And inArray(i, j) = row(1, j)
        Next
        If arrayContainsRow Then Exit Function
    Next
    
    arrayContainsRow = False
End Function

