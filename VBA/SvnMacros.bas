Attribute VB_Name = "SvnMacros"
'@Folder("Sovereign")
'@IgnoreModule ProcedureNotUsed
Option Explicit

Public Sub unhideAllWorksheets()
    Dim inSheet As Worksheet
    For Each inSheet In ActiveWorkbook.Sheets
        inSheet.Visible = xlSheetVisible
    Next inSheet
End Sub

Public Sub openPropertyExplorer()
    ''' Macro: Open UserForm '''
    Application.EnableCancelKey = xlInterrupt
    SvnPropertyExplorer.Instantiate(ActiveWorkbook).Show
    Application.EnableCancelKey = xlInterrupt
End Sub

Public Sub replaceAddressesFromNames()
    Application.Calculation = xlCalculationManual
    Dim currentName As Name
    With ActiveSheet.Cells
        For Each currentName In ActiveWorkbook.Names
            On Error Resume Next
            If (currentName.RefersToRange.Worksheet.Name = ActiveSheet.Name) Then .replace What:=currentName.Name, Replacement:=currentName.RefersToRange.Address(External:=False)
            On Error GoTo 0
        Next currentName
    End With
    Application.Calculation = xlCalculationAutomatic
End Sub

Public Sub replaceReferencesFromValues()
    Dim replaceRange As SvnMultiRange
    Dim searchRange As SvnMultiRange
    Dim replaceString As String
    Dim searchString As String
    replaceString = stringFromRefEdit(ActiveWorkbook, "Enter Range to Replace", "Replace References", Dimensions:=dimensionsFromRowsColsRepeatable(True, -1, -1))
    If replaceString = vbNullString Then Exit Sub
    searchString = stringFromRefEdit(ActiveWorkbook, "Enter Range to Search", "Replace References", Dimensions:=dimensionsFromRowsCols(-1, -1))
    
    Dim replaceArea As Range
    Dim replaceCell As Range
    
    For Each replaceArea In replaceRange.Areas
        For Each replaceCell In replaceArea.Cells
            On Error GoTo UDFErr
            Dim foundCell As Range
            Set foundCell = searchRange.Areas(1).Cells(WorksheetFunction.Match(replaceCell.Value2, searchRange, 0))
            replaceCell.Formula = "='" & foundCell.Parent.Name & "'!" & foundCell.Address(External:=False)
            GoTo NextCell
UDFErr:
            replaceCell.Font.Color = vbRed
            Resume NextCell
NextCell:
        Next
    Next
End Sub

Public Sub removeShapesByTitle()
    Dim title As String
    title = InputBox("Enter Title")
    Dim curShape As Shape
    For Each curShape In ActiveSheet.Shapes
        If curShape.title = title Then curShape.Delete
    Next
End Sub
