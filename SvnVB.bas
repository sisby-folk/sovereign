Attribute VB_Name = "SvnVB"
'@Folder("Sovereign")
Option Explicit

''' Collection Manipulation '''

Public Function collectionToArray(ByVal inCol As Collection) As Variant()
    Dim outVar() As Variant
    If (inCol.count < 1) Then
        collectionToArray = outVar
        Exit Function
    End If
    ReDim outVar(0 To inCol.count - 1)
    Dim itemIndex As Long
    For itemIndex = 1 To inCol.count
        outVar(itemIndex - 1) = inCol.Item(itemIndex)
    Next
    collectionToArray = outVar
End Function

Public Function HandleError(ByVal inErr As ErrObject) As Long
    HandleError = MsgBox(Err.Description & vbNewLine & vbNewLine & "Debug?", vbYesNo, "Error " & CStr(Err.Number) & " occurred during generation")
End Function

Public Function collectionToVariant(ByVal inCol As Collection) As Variant
    collectionToVariant = WorksheetFunction.Index(collectionToArray(inCol), 1, 0)
End Function

Public Function collectionSetByKey(ByVal inCollection As Collection, ByVal value As Variant, ByVal key As String, Optional ByVal before As Long = -1) As Collection
    inCollection.Remove key
    If before = -1 Then inCollection.Add value, key
    If before <> -1 Then inCollection.Add value, key, before
    Set collectionSetByKey = inCollection
End Function

Public Function collectionFromConstant(ByVal count As Long, ByVal value As Variant) As Collection
    Dim i As Long
    Set collectionFromConstant = New Collection
    For i = 1 To count
        collectionFromConstant.Add value
    Next
End Function

Public Function collectionFromKeys(ByVal keys As Collection, ByVal value As Variant) As Collection
    Dim i As Long
    Set collectionFromKeys = New Collection
    For i = 1 To keys.count
        collectionFromKeys.Add value, keys(i)
    Next
End Function

Public Function stringPluralize(ByVal count As Long, ByVal singular As String, Optional ByVal Del As String = " | ") As String
    stringPluralize = IIf(count > 0, CStr(count) & " " & singular & IIf(count = 1, vbNullString, "s") & Del, "")
End Function

Public Function variantGetType(ByVal inVar As Variant)
    Select Case VarType(inVar)
        Case vbEmpty
            variantGetType = "Empty"
        Case vbNull
            variantGetType = "Null"
        Case vbInteger
            variantGetType = "Integer"
        Case vbLong
            variantGetType = "Long"
        Case vbDouble
            variantGetType = "Double"
        Case vbCurrency
            variantGetType = "Currency"
        Case vbDate
            variantGetType = "Date"
        Case vbString
            variantGetType = "String"
        Case vbObject
            variantGetType = "Object"
        Case vbError
            variantGetType = "Error"
        Case vbBoolean
            variantGetType = "Boolean"
        Case Else
            variantGetType = CStr(VarType(inVar))
    End Select
End Function

Public Function variantGetPretty(ByVal inVar As Variant) As String
    If VarType(inVar) = vbError Then
        Select Case CLng(inVar)
            Case 2007
                variantGetPretty = "#DIV/0!"
            Case 2042
                variantGetPretty = "#N/A"
            Case 2029
                variantGetPretty = "#NAME?"
            Case 2000
                variantGetPretty = "#NULL!"
            Case 2036
                variantGetPretty = "#NUM!"
            Case 2023
                variantGetPretty = "#REF!"
            Case 2015
                variantGetPretty = "#VALUE!"
        End Select
    ElseIf IsArray(inVar) Then
        variantGetPretty = "Array dim " & arrayGetDimension(inVar)
    Else
        variantGetPretty = CStr(inVar)
    End If
End Function

Public Function collectionMoveDirection(ByVal inCol As Collection, ByVal atIndex As Long, ByVal directionUp As Boolean) As Boolean ' BREAKS KEYS (Duh)
    collectionMoveDirection = False
    If atIndex < 1 Or atIndex > inCol.count Then Exit Function
    If directionUp Then
        If atIndex > 1 Then                      ' Sample - Item 2
            '(1 2 3) -> (1 2 1 3)
            inCol.Add inCol(atIndex - 1), After:=atIndex
            '(1 2 1 3) -> (2 1 3)
            inCol.Remove atIndex - 1
            collectionMoveDirection = True
        End If
    Else
        If atIndex < inCol.count Then            ' Sample - Item 2
            '(1 2 3 4) -> (1 3 2 3 4)
            inCol.Add inCol(atIndex + 1), before:=atIndex
            '(1 3 2 3 4) -> (1 3 2 4)
            inCol.Remove atIndex + 2
            collectionMoveDirection = True
        End If
    End If
End Function

Public Sub Pass(Optional ByVal var As Variant, Optional ByVal obj As Object)
    GoTo Skip
    Pass var, obj
Skip:
End Sub

Public Function collectionGetOrNothing(ByVal fromCollection As Collection, ByVal key As String) As Object
    Dim outObj As Object
    On Error Resume Next
    Set outObj = fromCollection(key)
    Set collectionGetOrNothing = outObj
    On Error GoTo 0
End Function

Public Function collectionGetPropertyValueOrEmpty(ByVal fromCollection As Collection, ByVal key As String) As Variant
    Dim outVar As Variant
    On Error Resume Next
    outVar = fromCollection(key).value
    collectionGetPropertyValueOrEmpty = outVar
    On Error GoTo 0
End Function

Public Function collectionGetOrEmpty(ByVal fromCollection As Collection, ByVal key As String) As Variant
    Dim outVar As Variant
    On Error Resume Next
    outVar = fromCollection(key)
    collectionGetOrEmpty = outVar
    On Error GoTo 0
End Function

Public Function varOrEmpty(ByVal inFormula As String) As Variant
    Dim outVar As Variant
    On Error Resume Next
    outVar = Application.Evaluate(inFormula)
    If Not IsEmpty(outVar) Then varOrEmpty = outVar
    On Error GoTo 0
End Function

Public Function boolOrEmpty(ByVal inFormula As Variant) As Variant
    Dim outVar As Variant
    On Error Resume Next
    outVar = Application.Evaluate(inFormula)
    If Not IsEmpty(outVar) Then outVar = IIf(VarType(outVar) = vbBoolean, outVar, Empty)
    boolOrEmpty = outVar
    On Error GoTo 0
End Function

Public Sub collectionTryRemove(ByVal fromCollection As Collection, ByVal key As Variant)
    On Error Resume Next
    fromCollection.Remove key
    On Error GoTo 0
End Sub

Public Sub collectionTryAdd(ByVal fromCollection As Collection, ByRef inObj As Object, ByVal key As Variant)
    On Error Resume Next
    fromCollection.Add inObj, key
    On Error GoTo 0
End Sub

Public Sub collectionTryAddVar(ByVal fromCollection As Collection, ByRef inVar As Variant, ByVal key As Variant)
    On Error Resume Next
    fromCollection.Add inVar, key
    On Error GoTo 0
End Sub

Public Function collectionFromArgs(ParamArray items() As Variant) As Collection
    Dim i As Long
    Set collectionFromArgs = New Collection
    For i = LBound(items) To UBound(items)
        collectionFromArgs.Add items(i)
    Next
End Function

Public Function stringGetSplitCollection(ByVal inString As String, ByVal Delimiter As String) As Collection
    Dim outCol As Collection
    Set outCol = New Collection
    Dim eachPart As Variant
    For Each eachPart In Split(inString, Delimiter)
        outCol.Add eachPart, eachPart
    Next
    Set stringGetSplitCollection = outCol
End Function

Public Function stringGetWordCollection(ByVal inString As Variant, ByVal Del As String) As Collection
    Set stringGetWordCollection = New Collection
    On Error GoTo UDFErr
    Dim arr() As String
    arr = Split(inString, Del)
    On Error GoTo 0
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        stringGetWordCollection.Add arr(i)
    Next
Finish:
    Exit Function
UDFErr:
    stringGetWordCollection.Add inString
    Resume Finish
End Function

Public Sub showDebugMessage(ByVal Lines As Collection)
    Dim outString As String
    Dim currentLine As Variant
    For Each currentLine In Lines
        If Len(outString) + Len(currentLine) > 1024 Then
            MsgBox outString, vbOKOnly, "Debug Message"
            outString = vbNullString
        End If
        outString = outString + currentLine & vbNewLine
    Next
    If Len(outString) > 1 Then MsgBox outString
End Sub

Public Function assertionFails(ByVal Assertion As Boolean, ByVal Message As String, ByVal title As String) As Boolean
    assertionFails = Not Assertion
    If assertionFails Then MsgBox Message, vbOKOnly, title
End Function

Public Function stringTrim(ByVal inString As String, ByVal trimString As String) As String
    stringTrim = inString
    If (Left(inString, Len(trimString)) = trimString) Then stringTrim = Right(stringTrim, Len(stringTrim) - Len(trimString))
    If (Right(inString, Len(trimString)) = trimString) Then stringTrim = Left(stringTrim, Len(stringTrim) - Len(trimString))
End Function

Public Function collectionAppend(ByVal col1 As Collection, ByVal col2 As Collection) As Collection
    Set collectionAppend = New Collection
    Dim currentElement As Variant
    For Each currentElement In col1
        collectionAppend.Add currentElement
    Next
    For Each currentElement In col2
        collectionAppend.Add currentElement
    Next
End Function

Public Function getNonEmpty(ByVal inA As Variant, ByVal inB As Variant) As Variant
    If IsEmpty(inA) Then getNonEmpty = inB
    If IsEmpty(inB) Then getNonEmpty = inA
End Function

Public Function getNonNothing(ByVal inA As Object, ByVal inB As Object) As Object
    If inA Is Nothing Then Set getNonNothing = inB
    If inB Is Nothing Then Set getNonNothing = inA
End Function

Private Sub quickSortAuxilliary(ByRef Field() As String, ByVal LB As Long, ByVal UB As Long)
    Dim P1 As Long
    Dim P2 As Long
    Dim Ref As String
    Dim TEMP As String

    P1 = LB
    P2 = UB
    Ref = Field((P1 + P2) / 2)

    Do
        Do While (Field(P1) < Ref)
            P1 = P1 + 1
        Loop

        Do While (Field(P2) > Ref)
            P2 = P2 - 1
        Loop

        If P1 <= P2 Then
            TEMP = Field(P1)
            Field(P1) = Field(P2)
            Field(P2) = TEMP

            P1 = P1 + 1
            P2 = P2 - 1
        End If
    Loop Until (P1 > P2)

    If LB < P2 Then quickSortAuxilliary Field, LB, P2
    If P1 < UB Then quickSortAuxilliary Field, P1, UB
End Sub

Public Function stringSort(ByVal inString As String, Optional ByVal Del As String = ", ") As String
    Dim splitString() As String
    On Error GoTo UDFErr
    splitString = Split(inString, Del)
    On Error GoTo 0
    quickSortAuxilliary splitString, LBound(splitString), UBound(splitString)
    stringSort = Join(splitString, Del)
    If Left(stringSort, Len(Del)) = Del Then stringSort = Right(stringSort, Len(stringSort) - Len(Del))
    Exit Function
UDFErr:
    stringSort = inString
End Function

