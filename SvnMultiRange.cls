VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SvnMultiRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Sovereign")
Option Explicit

Private pSheetRangeCollections As Collection
Private pWorkbook As Workbook
Private pUnionize As Boolean
Private pEntireRows As Boolean
Private pSingleCell As Boolean

Public Property Get Unionize() As Boolean
    Unionize = pUnionize
End Property

Public Property Get EntireRows() As Boolean
    EntireRows = pEntireRows
End Property

Public Property Get SingleCell() As Boolean
    SingleCell = pSingleCell
End Property

Public Function GetUnionRange() As Range
    Dim eachArea As Range
    For Each eachArea In Me.Areas
        If GetUnionRange Is Nothing Then
            Set GetUnionRange = eachArea
        Else
            Set GetUnionRange = Application.Union(GetUnionRange, eachArea)
        End If
    Next
End Function

Public Function Worksheets() As Collection
    Set Worksheets = New Collection
    Dim worksheetAreas As Collection
    For Each worksheetAreas In pSheetRangeCollections
        Worksheets.Add worksheetAreas(1).Worksheet, worksheetAreas(1).Worksheet.Name
    Next
End Function

Public Function Areas() As Collection
    Set Areas = New Collection
    Dim currentWorksheetRanges As Collection
    Dim currentRange As Range
    Dim currentArea As Range
    For Each currentWorksheetRanges In pSheetRangeCollections
        For Each currentRange In currentWorksheetRanges
            For Each currentArea In currentRange.Areas
                Areas.Add currentArea
            Next
        Next
    Next
End Function

Public Function AreasBySheet(ByVal inWorksheet As Worksheet) As Collection
    Set AreasBySheet = New Collection
    Dim currentRange As Range
    Dim currentArea As Range
    If Not collectionGetOrNothing(pSheetRangeCollections, inWorksheet.Name) Is Nothing Then
      For Each currentRange In pSheetRangeCollections(inWorksheet.Name)
        For Each currentArea In currentRange.Areas
          AreasBySheet.Add currentArea
        Next
      Next
    End If
End Function

Public Function ToString() As String
    Dim sortedRanges As Collection
    Dim outString As String
    If Not pWorkbook Is Nothing Then
        Set sortedRanges = collectionGetOrderedSheets(pSheetRangeCollections, pWorkbook)

        Dim currentRanges As Collection
        Dim currentRange As Range
        For Each currentRanges In sortedRanges
            For Each currentRange In currentRanges
                outString = outString & rangeGetAddress(currentRange, Equals:=False, Comma:=True)
            Next
        Next
        If Len(outString) > 1 Then
            outString = "=" & Left$(outString, Len(outString) - 2)
        End If
    End If
    ToString = outString
End Function

Public Function Pretty() As String
    Dim sheetCount As Long
    Dim areaCount As Long
    sheetCount = Me.Worksheets.count
    areaCount = Me.Areas.count
    Pretty = stringPluralize(areaCount, "Area") & IIf(sheetCount > 1, " over " & stringPluralize(sheetCount, "Worksheet"), vbNullString) & ": " & Left(Me.ToString, 200) & "..."
End Function

Public Function DimensionPretty()
    Dim currentArea As Range
    For Each currentArea In Me.Areas
        DimensionPretty = DimensionPretty & " (" & currentArea.Rows.count & ", " & currentArea.Columns.count & ")"
    Next
End Function

Public Function AddRanges(ByVal Formula As String) As Boolean
    Dim areaString As Variant
    Dim newRange As Range
    AddRanges = True
    For Each areaString In Split(Formula, ",")
        Set newRange = rangeOrNothing(areaString, pWorkbook)
        If Not newRange Is Nothing Then
            AddRange IIf(pEntireRows, newRange.EntireRow, newRange)
        Else
            AddRanges = False
        End If
    Next
End Function

Private Sub AddRange(ByVal inRange As Range)
    Dim key As String
    key = inRange.Worksheet.Name
    If collectionGetOrNothing(pSheetRangeCollections, key) Is Nothing Then
        Dim worksheetAreaCollection As Collection
        Set worksheetAreaCollection = New Collection
        worksheetAreaCollection.Add inRange
        pSheetRangeCollections.Add worksheetAreaCollection, key
    Else
        Dim newRange As Range
        If pUnionize Then
            Set newRange = Application.Union(pSheetRangeCollections(key)(1), inRange)
            pSheetRangeCollections(key).Remove 1
            pSheetRangeCollections(key).Add newRange
        Else
            Set newRange = inRange
            ' XXX: Desperately needs some kind of duplication check
            pSheetRangeCollections(key).Add newRange
        End If
    End If
End Sub

Public Sub Union(ByVal inSvnMultiRange As SvnMultiRange)
    Dim eachRange As Range
    For Each eachRange In inSvnMultiRange.Areas
        AddRange eachRange
    Next
End Sub

Public Function Offset(ByVal Rows As Long, ByVal Cols As Long) As SvnMultiRange
    Set Offset = multiRangeFromInstantiate(pWorkbook)
    Dim eachArea As Range
    For Each eachArea In Me.Areas
        Offset.AddRanges rangeGetAddress(eachArea.Offset(Rows, Cols), Equals:=False)
    Next
End Function

Public Function RelativeResize(ByVal Rows As Long, ByVal Cols As Long) As SvnMultiRange
    Set RelativeResize = multiRangeFromInstantiate(pWorkbook)
    Dim eachArea As Range
    For Each eachArea In Me.Areas
        If eachArea.Rows.count + Rows > 0 And eachArea.Columns.count + Cols > 0 Then
            RelativeResize.AddRanges rangeGetAddress(eachArea.Resize(eachArea.Rows.count + Rows, eachArea.Columns.count + Cols), Equals:=False)
        Else
            RelativeResize.AddRanges rangeGetAddress(eachArea, Equals:=False)
        End If
    Next
End Function

Public Function SetRanges(ByVal inFormula As String) As SvnMultiRange
    Set pSheetRangeCollections = New Collection
    If inFormula <> vbNullString Then AddRanges inFormula
    Set SetRanges = Me
End Function

Public Function Edit(ByVal inWindowTitle As String, ByVal inDimensions As SvnDimensions) As Boolean
    With New SvnRangeEditor
        .Initialize Me, inWindowTitle, pWorkbook, inDimensions
        .Show
        Edit = .Dirty
        If Edit Then Me.SetRanges .Range.ToString
    End With
End Function

'' ** Signature Mirrored in SvnSovereign ** ''
Public Function Instantiate(ByVal inWorkbook As Workbook, Optional ByVal inFormula As String = vbNullString, Optional ByVal inUnionize As Boolean = False, Optional ByVal inEntireRows As Boolean = False, Optional ByVal inSingleCell As Boolean = False) As SvnMultiRange
    Set pSheetRangeCollections = New Collection
    Set pWorkbook = inWorkbook
    If inFormula <> vbNullString Then AddRanges inFormula
    pUnionize = inUnionize
    pEntireRows = inEntireRows
    pSingleCell = inSingleCell
    Set Instantiate = Me
End Function

