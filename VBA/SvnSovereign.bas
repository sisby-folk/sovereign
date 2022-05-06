Attribute VB_Name = "SvnSovereign"
'@Folder("Sovereign")
Option Explicit

'*********************************'
'' ** Freehand Instantiations ** ''
'*********************************'

Public Function multiRangeFromInstantiate(ByVal inWorkbook As Workbook, Optional ByVal inFormula As String = vbNullString, Optional ByVal inUnionize As Boolean = False, Optional ByVal inEntireRows As Boolean = False, Optional ByVal inSingleCell As Boolean = False) As SvnMultiRange
    Set multiRangeFromInstantiate = New SvnMultiRange
    multiRangeFromInstantiate.Instantiate inWorkbook:=inWorkbook, inFormula:=inFormula, inUnionize:=inUnionize, inEntireRows:=inEntireRows, inSingleCell:=inSingleCell
End Function

Public Function propertyFromInstantiate(ByVal inName As String, ByVal inValue As Variant, ByVal inType As MsoDocProperties, Optional ByVal inSheet As Worksheet = Nothing, Optional ByVal inRange As SvnMultiRange = Nothing) As SvnProperty
    Set propertyFromInstantiate = New SvnProperty
    propertyFromInstantiate.Instantiate inName:=inName, inValue:=inValue, inType:=inType, inSheet:=inSheet, inRange:=inRange
End Function

Public Function resultFromInstantiate(ByVal inResult As SvnResultType, Optional ByVal inMessage As String, Optional ByVal inErrorNum As Long, Optional ByVal inErrorDesc As String) As SvnResult
    Set resultFromInstantiate = New SvnResult
    resultFromInstantiate.Instantiate inResult:=inResult, inMessage:=inMessage, inErrorNum:=inErrorNum, inErrorDesc:=inErrorDesc
End Function

Public Function dimensionsFromRange(ByVal inRange As Range) As SvnDimensions
    Set dimensionsFromRange = New SvnDimensions
    dimensionsFromRange.FromRange inRange
End Function

Public Function dimensionsFromRowsCols(ParamArray var() As Variant) As SvnDimensions
    Set dimensionsFromRowsCols = New SvnDimensions
    dimensionsFromRowsCols.FromRowsCols var
End Function

Public Function dimensionsFromRowsColsRepeatable(ByVal crossSheet As Boolean, ByVal Rows As Long, ByVal Cols As Long) As SvnDimensions
    Set dimensionsFromRowsColsRepeatable = New SvnDimensions
    dimensionsFromRowsColsRepeatable.FromRowsColsRepeatable crossSheet, Rows, Cols
End Function

Public Function dimensionsFromRowsColsAreas(ByVal Rows As Long, ByVal Cols As Long, ByVal Areas As Long) As SvnDimensions
    Set dimensionsFromRowsColsAreas = New SvnDimensions
    dimensionsFromRowsColsAreas.FromRowsColsAreas Rows, Cols, Areas
End Function

'********************************'
'' ** Special Instantiations ** ''
'********************************'

Public Function multiRangeFromRefedit(ByVal inWorkbook As Workbook, ByVal InitialText As String, ByVal title As String, Optional ByVal Dimensions As SvnDimensions = Nothing, Optional ByVal Unionize As Boolean = False, Optional ByVal EntireRows As Boolean = False, Optional ByVal SingleCell As Boolean = False) As SvnMultiRange
    With New SvnRefEditor
        .Initialize inWorkbook, InitialText, title
        .Show vbModal
        Set multiRangeFromRefedit = multiRangeFromInstantiate(inWorkbook, inFormula:=.Ref, inEntireRows:=EntireRows, inUnionize:=Unionize, inSingleCell:=SingleCell)
        If Not Dimensions Is Nothing Then
            If Not Dimensions.Match(multiRangeFromRefedit) Then Set multiRangeFromRefedit = Nothing
        End If
    End With
End Function

Public Sub addPropertyWithKey(ByVal outCollection As Collection, ByVal inName As String, ByVal inValue As Variant, ByVal inType As MsoDocProperties, Optional ByVal inSheet As Worksheet = Nothing, Optional ByVal inRange As SvnMultiRange = Nothing)
    outCollection.Add propertyFromInstantiate(inName:=inName, inValue:=inValue, inType:=inType, inSheet:=inSheet, inRange:=inRange), key:=inName
End Sub
