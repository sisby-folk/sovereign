Attribute VB_Name = "SvnOffice"
'@Folder("Sovereign")
Option Explicit

Public Function pathFromPickerFolder(ByVal title As String, ByVal ButtonName As String, ByVal InitialPath As String) As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = title
        .ButtonName = ButtonName
        .AllowMultiSelect = False
        .InitialView = msoFileDialogViewDetails
        
        If Len(InitialPath) > 0 Then
            .InitialFileName = InitialPath & IIf(Right(InitialPath, 1) <> "\", "\", vbNullString)
        End If
        If .Show = -1 Then                       ' if OK is pressed
            pathFromPickerFolder = .SelectedItems(1)
        End If
    End With
End Function

Public Sub openInExplorer(ByVal FolderPath As String)
    Shell "explorer """ & FolderPath & "", vbNormalFocus
End Sub

Public Function nameOrNothing(ByVal inWorkbook As Workbook, ByVal inName As String) As Range
    On Error Resume Next
    Set nameOrNothing = inWorkbook.Names(inName).RefersToRange
    On Error GoTo 0
End Function

Public Function collectionAppendSheets(ByVal toCollection As Collection, ByVal fromCollection As Object, Optional ByVal nameOnly As Boolean = False) As Collection
    Dim currentSheet As Worksheet
    For Each currentSheet In fromCollection
        If Not nameOnly Then collectionTryAdd toCollection, currentSheet, currentSheet.Name
        If nameOnly Then collectionTryAddVar toCollection, currentSheet.Name, currentSheet.Name
    Next
    Set collectionAppendSheets = toCollection
End Function

Public Function collectionRemoveSheets(ByVal toCollection As Collection, ByVal fromCollection As Object) As Collection
    Dim currentSheet As Worksheet
    For Each currentSheet In fromCollection
        collectionTryRemove toCollection, currentSheet.Name
    Next
    Set collectionRemoveSheets = toCollection
End Function

Public Function collectionGetOrderedSheets(ByVal sheetCollection As Collection, ByVal inWorkbook As Workbook) As Collection
    Dim outCollection As Collection
    Set outCollection = New Collection
    Dim currentSheet As Variant
    For Each currentSheet In inWorkbook.Sheets
        If Not collectionGetOrNothing(sheetCollection, currentSheet.Name) Is Nothing Then outCollection.Add sheetCollection(currentSheet.Name), currentSheet.Name
    Next
    Set collectionGetOrderedSheets = outCollection
End Function

Public Function rangeInsertPicture(ByVal Path As String, ByVal inRange As Range, ByVal title As String) As Shape
    Dim rangeWidth As Double
    Dim rangeHeight As Double
    If inRange.MergeCells Then
        inRange.MergeArea.EntireRow.Hidden = False
        rangeWidth = inRange.MergeArea.Width
        rangeHeight = inRange.MergeArea.Height
    Else
        inRange.EntireRow.Hidden = False
        rangeWidth = inRange.Width
        rangeHeight = inRange.Height
    End If

    Set rangeInsertPicture = inRange.Worksheet.Shapes.AddPicture2(FileName:=Path, LinkToFile:=msoTrue, SaveWithDocument:=msoFalse, Left:=inRange.Left + rangeWidth * 0.1, Top:=inRange.Top + rangeWidth * 0.1, Width:=200, Height:=200, Compress:=msoPictureCompressFalse)
    rangeInsertPicture.LockAspectRatio = msoTrue
    rangeInsertPicture.title = title
    rangeInsertPicture.Width = rangeWidth * 0.8
    If rangeInsertPicture.Height > rangeHeight * 0.8 Then rangeInsertPicture.Height = rangeHeight * 0.8
End Function

Public Function rangeGetOffsetByStep(ByVal inRange As Range, ByVal Rows As Long) As Range
    Dim currentArea As Range
    Dim rangeString As String
    For Each currentArea In inRange.Areas
        rangeString = rangeString & IIf(rangeString = vbNullString, vbNullString, ",") & currentArea.Worksheet.Rows(currentArea.row + Rows).Columns(currentArea.column).Address
    Next
    Set rangeGetOffsetByStep = inRange.Worksheet.Range(rangeString)
End Function

Public Function rangeOrNothing(ByVal inFormula As Variant, ByVal inWorkbook As Workbook) As Range
    Dim inWorksheet As Worksheet
    On Error Resume Next
    If InStr(inFormula, "!") <> 0 Then Set inWorksheet = inWorkbook.Sheets(stringTrim(stringTrim(stringTrim(getWord(inFormula, "!", 0), "="), " "), "'"))
    If inWorksheet Is Nothing Then Set rangeOrNothing = ActiveSheet.Range(inFormula)
    If Not inWorksheet Is Nothing Then Set rangeOrNothing = inWorksheet.Range(inFormula)
    On Error GoTo 0
End Function

Public Function rangeGetMatchOrEmpty(ByVal inValue As Variant, ByVal inRange As Range) As Variant
    On Error Resume Next
    rangeGetMatchOrEmpty = WorksheetFunction.Match(inValue, inRange, 0)
    On Error GoTo 0
End Function

Public Function rangeGetText(ByVal Rng As Range, Optional ByVal Delimiter As String = ",", Optional ByVal EOL As String = "@") As String
    Dim outString As String
    Dim currentRow As Range
    Dim currentCell As Range
    For Each currentRow In Rng.Rows
        For Each currentCell In currentRow.Columns
            outString = outString & currentCell.Value2 & Delimiter
        Next
        outString = outString & EOL
    Next
    rangeGetText = IIf(Len(outString) > 1, Left$(outString, Len(outString) - 2), outString)
End Function

Public Function rangeGetAddress(ByVal Source As Range, Optional ByVal Equals As Boolean = True, Optional ByVal Comma As Boolean = False) As String
    Dim outString As String
    Dim currentArea As Range
    If Source Is Nothing Then Exit Function
    For Each currentArea In Source.Areas
        outString = outString & "'" & currentArea.Worksheet.Name & "'!" & currentArea.Address & ", "
    Next
    If Len(outString) > 1 Then
        outString = IIf(Equals, "=", vbNullString) & Left$(outString, Len(outString) - IIf(Comma, 0, 2))
    End If
    rangeGetAddress = outString
End Function

Public Function sheetFromMakeOrClear(ByVal inWorkbook As Workbook, ByVal SheetName As String) As Worksheet
    Set sheetFromMakeOrClear = collectionGetOrNothing(workbookGetSheets(inWorkbook), SheetName)
    If Not sheetFromMakeOrClear Is Nothing Then
        sheetFromMakeOrClear.Cells.Clear
    Else
        inWorkbook.Activate
        inWorkbook.Sheets(1).Select              ' Deselect Sheets to avoid double sheets
        Set sheetFromMakeOrClear = inWorkbook.Sheets.Add
        sheetFromMakeOrClear.Name = SheetName
    End If
End Function

Public Function workbookFromPicker(ByVal title As String, Optional ByVal IgnoreSelf As Boolean = False) As Workbook

    ' Get Excel Application
    Dim excelApp As Object
    On Error Resume Next
    Set excelApp = GetObject(Class:="Excel.Application")
    On Error GoTo 0
    If excelApp Is Nothing Then
        MsgBox "Excel Is Not Open", vbOKOnly, "Error"
        Exit Function
    End If
    
    ' Get Workbook Names
    Dim nameString As Collection
    Set nameString = New Collection
    
    Dim currentWorkbook As Workbook
    For Each currentWorkbook In excelApp.Workbooks
        If currentWorkbook.Name <> ThisWorkbook.Name Or Not IgnoreSelf Then nameString.Add currentWorkbook.Name
    Next
    
    ' Choose Workbook
    Dim chosenWorkbook As String
    With New SvnPicker
        .Initialize title, nameString
        .Show
        If .Success Then chosenWorkbook = nameString(.Index)
    End With
    
    For Each currentWorkbook In excelApp.Workbooks
        If chosenWorkbook = currentWorkbook.Name Then
            Set workbookFromPicker = currentWorkbook
        End If
    Next
End Function

Public Function stringFromRefEdit(ByVal inWorkbook As Workbook, ByVal InitialText As String, ByVal title As String, Optional ByVal Dimensions As SvnDimensions = Nothing, Optional ByVal Unionize As Boolean = False, Optional ByVal EntireRows As Boolean = False, Optional ByVal SingleCell As Boolean = False) As String
    Dim newRange As SvnMultiRange
    Set newRange = multiRangeFromRefedit(inWorkbook:=inWorkbook, InitialText:=InitialText, title:=title, Dimensions:=Dimensions, Unionize:=Unionize, EntireRows:=EntireRows, SingleCell:=SingleCell)
    If Not newRange Is Nothing Then stringFromRefEdit = newRange.ToString
End Function

Public Function stringFromFormulaEdit(ByVal InitialText As String, ByVal title As String) As String
    With New SvnFormulaEditor
        .Initialize InitialText, title
        .Show vbModal
        If Not IsEmpty(varOrEmpty(.Formula)) Then
            stringFromFormulaEdit = IIf(Left(.Formula, 1) = "=", vbNullString, "=") & IIf(.Formula = vbNullString, Empty, .Formula)
        End If
    End With
End Function

Public Function rangeGetURLs(ByVal inRange As Range) As Variant
    Dim outArray As Variant
    outArray = arrayForce2D(inRange.Value2)
    Dim i As Long
    Dim j As Long
    For i = 1 To inRange.Rows.count
        For j = 1 To inRange.Columns.count
            If inRange.Rows(i).Columns(j).Hyperlinks.count > 0 Then outArray(i, j) = inRange.Rows(i).Columns(j).Hyperlinks(1).Address
        Next
    Next
    rangeGetURLs = outArray
End Function

Public Function documentFromPicker() As Document
    ' Get Word Application
    Dim wordApp As Object
    On Error Resume Next
    Set wordApp = GetObject(Class:="Word.Application")
    On Error GoTo 0
    If wordApp Is Nothing Then
        MsgBox "Word Is Not Open", vbOKOnly, "Error"
        Exit Function
    End If
    
    ' Get Document Names
    Dim nameString As Collection
    Set nameString = New Collection
    
    Dim currentDocument As Document
    For Each currentDocument In wordApp.Documents
        nameString.Add currentDocument.Name
    Next
    
    If nameString.count = 0 Then
        MsgBox "Fatal Error - Word is Open but no document could be found (" & CStr(wordApp.Documents.count) & " documents) - Possibly a windows permission error?"
        Exit Function
    End If
    
    ' Choose Word Document
    Dim chosenDocument As String
    With New SvnPicker
        .Initialize "Please Choose Report to Autofill:", nameString
        .Show
        If .Success Then chosenDocument = nameString(.Index)
    End With
    
    For Each currentDocument In wordApp.Documents
        If chosenDocument = currentDocument.Name Then
            Set documentFromPicker = currentDocument
        End If
    Next
End Function

Public Function sheetFromPicker(ByVal inWorkbook As Workbook) As Worksheet
    With New SvnPicker
        .Initialize " Please choose a sheet:", sheetsToNames(workbookGetSheets(inWorkbook))
        .Show
        If .Success Then Set sheetFromPicker = inWorkbook.Sheets(.Text)
    End With
End Function

Public Sub documentAddComment(ByVal Report As Document, ByVal tableTitle As String, ByVal contents As String)
    Dim foundTable As Object
    Dim currentTable As Object
    For Each currentTable In Report.Tables
        If currentTable.title = tableTitle Then Set foundTable = currentTable
    Next
    If foundTable Is Nothing Then Exit Sub
    On Error Resume Next
    Report.Comments.Add foundTable.cell(1, 1).Range, contents
    On Error GoTo 0
End Sub

Public Function documentGetTable(ByVal Report As Document, ByVal tableTitle As String) As Object
    Dim currentTable As Object
    For Each currentTable In Report.Tables
        If currentTable.title = tableTitle Then Set documentGetTable = currentTable
    Next
End Function

Public Sub documentClearComments(ByVal Report As Document, ByVal tableTitle As String)
    Dim foundTable As Object
    Dim currentTable As Object
    For Each currentTable In Report.Tables
        If currentTable.title = tableTitle Then Set foundTable = currentTable
    Next
    If foundTable Is Nothing Then Exit Sub
    On Error Resume Next
    Dim currentComment As Object
    For Each currentComment In foundTable.cell(1, 1).Range.Comments
        currentComment.DeleteRecursively
    Next
    On Error GoTo 0
End Sub

Public Sub rangeCopyURLs(ByVal DestRange As Range, ByVal SourceRange As Range)
    Dim currentRow As Long
    Dim destCell As Range
    Dim sourceCell As Range
    Dim currentLink As Hyperlink
    For currentRow = 1 To SourceRange.Rows.count
        Set destCell = DestRange.Rows(currentRow)
        Set sourceCell = SourceRange.Rows(currentRow)
        For Each currentLink In sourceCell.Hyperlinks
            destCell.Hyperlinks.Add destCell, currentLink.Address, TextToDisplay:=currentLink.TextToDisplay
        Next
    Next
End Sub

Public Function collectionGetSheetVariant(ByVal inCol As Collection) As Variant
    collectionGetSheetVariant = WorksheetFunction.Index(collectionToArray(sheetsToNames(inCol)), 1, 0)
End Function

Public Function sheetsToNames(ByVal inCol As Collection) As Collection
    Dim outCol As Collection
    Set outCol = New Collection
    Dim currentSheet As Worksheet
    For Each currentSheet In inCol
        outCol.Add currentSheet.Name
    Next
    Set sheetsToNames = outCol
End Function

Public Function workbookGetSheets(ByVal inWorkbook As Workbook) As Collection
    Set workbookGetSheets = New Collection
    Dim currentSheet As Worksheet
    For Each currentSheet In inWorkbook.Sheets
        workbookGetSheets.Add currentSheet, currentSheet.Name
    Next
End Function

Public Function workbookGetNames(ByVal inWorkbook As Workbook) As Collection
    Set workbookGetNames = New Collection
    Dim currentName As Name
    For Each currentName In inWorkbook.Names
        workbookGetNames.Add currentName, currentName.Name
    Next
End Function

Public Sub labelSetContent(ByVal inLabel As MSForms.Label, ByVal inCaption As String, ByVal inColour As Long)
    inLabel.Caption = inCaption
    inLabel.ForeColor = inColour
End Sub

Public Function stringGetCleanSheetName(ByVal SheetName As String) As String
    Const specialChars As String = "\,/,*,?,:,[,]"
    Dim specialChar As Variant
    Dim outName As String
    outName = SheetName
    ' Remove special characters
    For Each specialChar In Split(specialChars, ",")
        outName = replace(outName, specialChar, vbNullString)
    Next
    
    ' Limit Length
    outName = Left(SheetName, 31)
    
    ' Ensure uniqueness
    Dim currentSheet As Worksheet
    For Each currentSheet In ActiveWorkbook.Worksheets
        If currentSheet.Name = outName Then Exit Function ' Will return null string
    Next
    
    stringGetCleanSheetName = outName
End Function

