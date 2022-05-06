VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SvnPropertyExplorer 
   Caption         =   "Property Explorer"
   ClientHeight    =   6120
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15360
   OleObjectBlob   =   "SvnPropertyExplorer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SvnPropertyExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Sovereign")
Option Explicit

Private pWorkbook As Workbook
Private pProperties As Collection
Private pSheetProperties As Collection
Private pSheets As Collection
Private pTypes As Collection

Private Sub Refresh()
    ' Load Properties
    Set pProperties = workbookGetProperties(pWorkbook)
    Set pSheetProperties = New Collection
    Dim currentSheet As Worksheet
    For Each currentSheet In pSheets
        pSheetProperties.Add worksheetGetProperties(currentSheet), currentSheet.Name
    Next

    ' Preserve Selections
    Dim listPropertyIndex As Long
    listPropertyIndex = ListPropBook.ListIndex
    Dim listSheetIndex As Long
    listSheetIndex = ListSheet.ListIndex
    Dim listSheetPropIndex As Long
    listSheetPropIndex = ListPropSheet.ListIndex
    
    
    ' Show Workbook Properties
    ListPropBook.Clear
    Dim currentProperty As SvnProperty
    For Each currentProperty In pProperties
        ListPropBook.AddItem
        ListPropBook.list(ListPropBook.ListCount - 1, 0) = IIf(VarType(currentProperty.value) = 8, "String", IIf(VarType(currentProperty.value) = 3, "Long", IIf(VarType(currentProperty.value) = 11, "Boolean", VarType(currentProperty.value))))
        ListPropBook.list(ListPropBook.ListCount - 1, 1) = currentProperty.Name
        ListPropBook.list(ListPropBook.ListCount - 1, 2) = CStr(currentProperty.value)
    Next
    If listPropertyIndex < ListPropBook.ListCount Then ListPropBook.ListIndex = listPropertyIndex
    
    ' Show Worksheets
    ListSheet.Clear
    For Each currentSheet In pSheets
        ListSheet.AddItem currentSheet.Name
    Next
    If listSheetIndex < ListSheet.ListCount Then ListSheet.ListIndex = listSheetIndex
    ListSheet_Change
    If listSheetPropIndex < ListPropSheet.ListCount Then ListPropSheet.ListIndex = listSheetPropIndex
End Sub

Private Sub ListPropBook_Click()
    If ListPropBook.ListIndex = -1 Then
    LabelWorkbookPropertyName = vbNullString
    LabelWorkbookPropertyValue = vbNullString
    Else
        Dim selectedProperty As SvnProperty
        Set selectedProperty = pProperties(ListPropBook.ListIndex + 1)
    LabelWorkbookPropertyName = selectedProperty.Name
    LabelWorkbookPropertyValue = CStr(selectedProperty.value)
    End If
End Sub

Private Sub ListPropSheet_Click()
    If ListSheet.ListIndex = -1 Or ListPropSheet.ListIndex = -1 Then
    LabelSheetPropertyName = vbNullString
    LabelSheetPropertyValue = vbNullString
    Else
    Dim selectedProperty As SvnProperty
    Set selectedProperty = pSheetProperties(ListSheet.ListIndex + 1)(ListPropSheet.ListIndex + 1)
    LabelSheetPropertyName = selectedProperty.Name
    LabelSheetPropertyValue = CStr(selectedProperty.value)
    End If
End Sub

Private Sub ListSheet_Change()
    ListPropSheet.Clear
    Dim currentProperty As SvnProperty
    If ListSheet.ListIndex <> -1 Then
        For Each currentProperty In pSheetProperties(pSheets(ListSheet.ListIndex + 1).Name)
            ListPropSheet.AddItem
            ListPropSheet.list(ListPropSheet.ListCount - 1, 0) = currentProperty.Name
            ListPropSheet.list(ListPropSheet.ListCount - 1, 1) = currentProperty.value
        Next
    End If
    ListPropSheet_Click
End Sub

Private Sub AddPropertyDialog(ByVal PropType As MsoDocProperties, Optional ByVal sheet As Worksheet = Nothing)
    Dim newPropName As String
    newPropName = InputBox("Enter Property Name:", "Add New Property")
    If newPropName = vbNullString Then Exit Sub
    
    Dim newPropValue As Variant
    newPropValue = InputBox("Enter Property Value:", "Add New Property")
    If newPropValue = vbNullString Then Exit Sub
    
    If Not sheet Is Nothing Then worksheetSetProperty pWorkbook, propertyFromInstantiate(newPropName, newPropValue, PropType, sheet)
    If sheet Is Nothing Then workbookSetProperty pWorkbook, propertyFromInstantiate(newPropName, newPropValue, PropType, sheet)
    
    Refresh
End Sub

Private Sub ButtonEditPropBook_Click()
    If ListPropBook.ListIndex = -1 Then Exit Sub
    Dim selectedProperty As SvnProperty
    Set selectedProperty = pProperties(ListPropBook.ListIndex + 1)

    Dim inputString As String
    inputString = InputBox(selectedProperty.value, "Edit Workbook Property: " & selectedProperty.Name)
    If inputString = vbNullString Then Exit Sub
    selectedProperty.value = inputString
    
    If Not IsEmpty(selectedProperty.value) Then
        workbookSetProperty pWorkbook, selectedProperty
    End If
    
    Refresh
End Sub

Private Sub ButtonEditPropSheet_Click()
    If ListSheet.ListIndex = -1 Or ListPropSheet.ListIndex = -1 Then Exit Sub
    Dim selectedProperty As SvnProperty
    Set selectedProperty = pSheetProperties(ListSheet.ListIndex + 1)(ListPropSheet.ListIndex + 1)
    
    Dim inputString As String
    inputString = InputBox(selectedProperty.value, "Edit Worksheet Property: " & selectedProperty.sheet.Name & " - " & selectedProperty.Name)
    If inputString = vbNullString Then Exit Sub
    selectedProperty.value = inputString
    
    If Not IsEmpty(selectedProperty.value) Then
        worksheetSetProperty pWorkbook, selectedProperty
    End If

    Refresh
End Sub

Private Sub ButtonRemovePropBook_Click()
    If ListPropBook.ListIndex = -1 Then Exit Sub
    Dim selectedProperty As SvnProperty
    Set selectedProperty = pProperties(ListPropBook.ListIndex + 1)
    
    workbookRemoveProperty pWorkbook, selectedProperty.Name
    Refresh
End Sub

Private Sub ButtonRemovePropSheet_Click()
    If ListPropSheet.ListIndex = -1 Or ListSheet.ListIndex = -1 Then Exit Sub
    Dim selectedProperty As SvnProperty
    Set selectedProperty = pSheetProperties(ListSheet.ListIndex + 1)(ListPropSheet.ListIndex + 1)
    
    worksheetRemoveProperty selectedProperty.sheet, selectedProperty.Name
    Refresh
End Sub

Private Sub ButtonAddPropBook_Click()
    AddPropertyDialog pTypes(ComboType.ListIndex + 1)
End Sub

Private Sub ButtonAddPropSheet_Click()
    If ListSheet.ListIndex = -1 Then Exit Sub
    AddPropertyDialog MsoDocProperties.msoPropertyTypeString, pSheets(ListSheet.ListIndex + 1)
End Sub

Public Function Instantiate(ByVal inWorkbook As Workbook) As SvnPropertyExplorer
    Set pWorkbook = inWorkbook
    Set pSheets = workbookGetSheets(inWorkbook)
    Set pTypes = New Collection
    pTypes.Add msoPropertyTypeString, "String"
    pTypes.Add msoPropertyTypeBoolean, "Boolean"
    pTypes.Add msoPropertyTypeDate, "Date"
    pTypes.Add msoPropertyTypeFloat, "Float"
    pTypes.Add msoPropertyTypeNumber, "Number"
    ComboType.AddItem "String"
    ComboType.AddItem "Boolean"
    ComboType.AddItem "Date"
    ComboType.AddItem "Float"
    ComboType.AddItem "Number"
    Refresh
    Set Instantiate = Me
End Function
