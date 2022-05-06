Attribute VB_Name = "SvnProperties"
'@Folder("Sovereign")
Option Explicit

Public Function workbookGetProperties(ByVal inWorkbook As Workbook, Optional ByVal Prefix As String = vbNullString, Optional ByVal Name As String = vbNullString) As Collection '<SvnProperty>
    Set workbookGetProperties = New Collection
    Dim currentProperty As DocumentProperty
    For Each currentProperty In inWorkbook.CustomDocumentProperties
        If Left$(currentProperty.Name, Len(Prefix)) = Prefix And (Name = vbNullString Or Prefix & Name = currentProperty.Name) Then
        addPropertyWithKey workbookGetProperties, Right(currentProperty.Name, Len(currentProperty.Name) - Len(Prefix)), currentProperty.value, currentProperty.Type
        End If
    Next
End Function

Public Function worksheetGetProperties(ByVal inSheet As Worksheet, Optional ByVal Prefix As String = vbNullString, Optional ByVal Name As String = vbNullString) As Collection  '<SvnProperty>
    Set worksheetGetProperties = New Collection
    Dim currentProperty As CustomProperty
    For Each currentProperty In inSheet.CustomProperties
        If Left$(currentProperty.Name, Len(Prefix)) = Prefix And (Name = vbNullString Or Prefix & Name = currentProperty.Name) Then addPropertyWithKey worksheetGetProperties, Right(currentProperty.Name, Len(currentProperty.Name) - Len(Prefix)), currentProperty.value, inType:=msoPropertyTypeString, inSheet:=currentProperty.Parent
    Next
End Function

Public Function workbookGetSheetProperties(ByVal inWorkbook As Workbook, Optional ByVal Prefix As String = vbNullString) As Collection '<SvnProperty>
    Set workbookGetSheetProperties = New Collection
    Dim currentSheet As Worksheet
    For Each currentSheet In workbookGetSheets(inWorkbook)
        Set workbookGetSheetProperties = collectionAppend(workbookGetSheetProperties, worksheetGetProperties(currentSheet, Prefix))
    Next
End Function

Public Function workbookGetNameProperties(ByVal inWorkbook As Workbook, Optional ByVal Prefix As String = vbNullString) As Collection '<SvnProperty>
    Set workbookGetNameProperties = New Collection
    Dim currentName As Name
    For Each currentName In workbookGetNames(inWorkbook)
        If Left$(currentName.Name, Len(Prefix)) = Prefix Then
            addPropertyWithKey workbookGetNameProperties, inName:=Right(currentName.Name, Len(currentName.Name) - Len(Prefix)), inValue:=currentName.RefersTo, inType:=msoPropertyTypeString
        End If
    Next
End Function

' Batch Overwriters - XXX: Needs a better backup system
Public Sub workbookOverwriteProperties(ByVal inWorkbook As Workbook, ByVal properties As Collection, Optional ByVal Prefix As String)
    Dim currentProperty As DocumentProperty
    For Each currentProperty In inWorkbook.CustomDocumentProperties
        If Left$(currentProperty.Name, Len(Prefix)) = Prefix Then currentProperty.Delete
    Next
    
    Dim remainingProperty As SvnProperty
    For Each remainingProperty In properties
        inWorkbook.CustomDocumentProperties.Add Name:=Prefix & remainingProperty.Name, LinkToContent:=False, Type:=remainingProperty.PropType, value:=remainingProperty.value
    Next
End Sub

Public Sub workbookOverwriteSheetProperties(ByVal inWorkbook As Workbook, ByVal properties As Collection, Optional ByVal Prefix As String)
    Dim inSheet As Worksheet
    For Each inSheet In workbookGetSheets(inWorkbook)
        Dim currentProperty As CustomProperty
        For Each currentProperty In inSheet.CustomProperties
            If Prefix <> vbNullString And Left$(currentProperty.Name, Len(Prefix)) = Prefix Then currentProperty.Delete
        Next
    Next
    
    Dim remainingProperty As SvnProperty
    For Each remainingProperty In properties
        If Prefix = vbNullString Then
            If Not IsEmpty(collectionGetPropertyValueOrEmpty(worksheetGetProperties(remainingProperty.sheet), remainingProperty.Name)) Then worksheetGetProperties(remainingProperty.sheet)(remainingProperty.Name).Delete
        End If
        remainingProperty.sheet.CustomProperties.Add Prefix & remainingProperty.Name, remainingProperty.value
    Next
End Sub

Public Sub workbookOverwriteNamedRanges(ByVal inWorkbook As Workbook, ByVal properties As Collection, Optional ByVal Prefix As String)
    Dim currentName As Name
    For Each currentName In workbookGetNames(inWorkbook)
        If Left$(currentName.Name, Len(Prefix)) = Prefix Then currentName.Delete
    Next
    
    Dim remainingProperty As SvnProperty
    For Each remainingProperty In properties
        If Len(remainingProperty.value) > 255 Then inWorkbook.Names.Add Prefix & remainingProperty.Name, remainingProperty.Range.GetUnionRange
        If Len(remainingProperty.value) <= 255 Then inWorkbook.Names.Add Prefix & remainingProperty.Name, remainingProperty.value
    Next
End Sub

Public Sub worksheetSetProperty(ByVal inWorkbook As Workbook, ByVal inProp As SvnProperty, Optional ByVal Prefix As String = vbNullString)
    workbookOverwriteSheetProperties inWorkbook, collectionFromArgs(propertyFromInstantiate(Prefix & inProp.Name, inProp.value, inProp.PropType, inProp.sheet)), vbNullString
End Sub

Public Sub workbookSetProperty(ByVal inWorkbook As Workbook, ByVal inProp As SvnProperty, Optional ByVal Prefix As String = vbNullString) ' Uses funky behaviour
    workbookOverwriteProperties inWorkbook, collectionFromArgs(propertyFromInstantiate(vbNullString, inProp.value, inProp.PropType)), Prefix & inProp.Name
End Sub

' Remove Properties (doesn't use name)

Public Sub worksheetRemoveProperty(ByVal inSheet As Worksheet, ByVal Name As String)
    Dim currentProperty As CustomProperty
    For Each currentProperty In inSheet.CustomProperties
        If currentProperty.Name = Name Then currentProperty.Delete
    Next
End Sub

Public Sub workbookRemoveProperty(ByVal inWorkbook As Workbook, ByVal Name As String)
    Dim currentProperty As DocumentProperty
    For Each currentProperty In inWorkbook.CustomDocumentProperties
        If currentProperty.Name = Name Then currentProperty.Delete
    Next
End Sub


'''' Word Document Properties (Not bound by SvnProperty) ''''

Public Function documentSetCustomProperty(ByVal inDocument As Document, ByVal Name As String, ByVal newValue As Variant, Optional ByVal ReplaceZero As Boolean = False) As Boolean
    Dim currentProperty As Object
    documentSetCustomProperty = False
    For Each currentProperty In inDocument.CustomDocumentProperties
        If currentProperty.Name = Name Then
            If currentProperty.Type = msoPropertyTypeDate Then
                currentProperty.value = CLng(CDate("1970-01-01")) ' Insert Bad Date
                On Error Resume Next
                currentProperty.value = CLng(CDate(newValue))
                On Error GoTo 0
            Else
                currentProperty.value = CStr(IIf(ReplaceZero And newValue = 0, " ", newValue))
            End If
            documentSetCustomProperty = True
        End If
    Next
End Function

Public Function documentSetBuiltinProperty(ByVal inDocument As Document, ByVal Name As String, ByVal newValue As Variant, Optional ByVal ReplaceZero As Boolean = False) As Boolean
    Dim currentProperty As Object
    documentSetBuiltinProperty = False
    For Each currentProperty In inDocument.BuiltinDocumentProperties
        If currentProperty.Name = Name Then
            currentProperty.value = CStr(IIf(ReplaceZero And newValue = 0, " ", newValue))
            documentSetBuiltinProperty = True
        End If
    Next
End Function

Public Function documentGetCustomProperty(ByVal inDocument As Document, ByVal Name As String, Optional ByVal ReplaceZero As Boolean = False) As Variant
    Dim currentProperty As Object
    For Each currentProperty In inDocument.CustomDocumentProperties
        If currentProperty.Name = Name Then
            documentGetCustomProperty = IIf(ReplaceZero And currentProperty.value = " ", 0, currentProperty.value)
        End If
    Next
End Function

Public Function documentGetBuiltinProperty(ByVal inDocument As Document, ByVal Name As String) As Variant
    Dim currentProperty As Object
    For Each currentProperty In inDocument.BuiltinDocumentProperties
        If currentProperty.Name = Name Then
            documentGetBuiltinProperty = IIf(currentProperty.value = " ", vbNullString, currentProperty.value)
        End If
    Next
End Function

