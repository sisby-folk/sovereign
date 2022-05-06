VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SvnDimensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Sovereign")
Option Explicit

Private pRangeAreas As Collection
Private pRepeatableAreas As Boolean
Private pCrossSheet As Boolean

Property Get crossSheet() As Boolean
    crossSheet = pCrossSheet
End Property

Public Function Match(ByVal inRange As SvnMultiRange) As Boolean
    Match = False
    If Not pRepeatableAreas And inRange.Areas.count <> pRangeAreas.count Then Exit Function
    If Not pCrossSheet And inRange.Worksheets.count > 1 Then Exit Function
    Dim cArea As Long
    For cArea = 1 To inRange.Areas.count
        If pRepeatableAreas Then
            If (inRange.Areas(cArea).Rows.count <> pRangeAreas(1)(0) And pRangeAreas(1)(0) <> -1) Or (inRange.Areas(cArea).Columns.count <> pRangeAreas(1)(1) And pRangeAreas(1)(1) <> -1) Then Exit Function
        Else
            If (inRange.Areas(cArea).Rows.count <> pRangeAreas(cArea)(0) And pRangeAreas(cArea)(0) <> -1) Or (inRange.Areas(cArea).Columns.count <> pRangeAreas(cArea)(1) And pRangeAreas(cArea)(1) <> -1) Then Exit Function
        End If
    Next
    Match = True
End Function

Public Function Pretty() As String
    Pretty = IIf(pCrossSheet, "(Multi-Sheet) ", vbNullString) & IIf(pRepeatableAreas, "Repeatable:", "Non-Repeating:")
    Dim rangeArea As Variant
    For Each rangeArea In pRangeAreas
        Pretty = Pretty & " (" & rangeArea(0) & ", " & rangeArea(1) & ")"
    Next
End Function

'' **         - Instantiations -          ** ''
'' ** Signatures Mirrored in SvnSovereign ** ''

Public Function FromRange(ByVal inRange As Range) As SvnDimensions
    Set pRangeAreas = New Collection
    pRepeatableAreas = False
    pCrossSheet = False
    Dim cArea As Range
    For Each cArea In inRange.Areas
        pRangeAreas.Add Array(cArea.Rows.count, cArea.Columns.count)
    Next
    Set FromRange = Me
End Function

' Not Enforceable cross-sheet unless MultiRange follows area order cross-sheet
Public Function FromRowsCols(ParamArray var() As Variant) As SvnDimensions
    Set pRangeAreas = New Collection
    pRepeatableAreas = False
    pCrossSheet = False
    Dim i As Long
    For i = LBound(var(0)) + 1 To UBound(var(0)) Step 2 ' Each Area
        pRangeAreas.Add Array(var(0)(i - 1), var(0)(i))
    Next
    Set FromRowsCols = Me
End Function

Public Function FromRowsColsRepeatable(ByVal inCrossSheet As Boolean, ByVal Rows As Long, ByVal Cols As Long) As SvnDimensions
    Set pRangeAreas = New Collection
    pCrossSheet = inCrossSheet
    pRepeatableAreas = True
    pRangeAreas.Add Array(Rows, Cols)
    Set FromRowsColsRepeatable = Me
End Function

Public Function FromRowsColsAreas(ByVal Cols As Long, ByVal Rows As Long, ByVal Areas As Long) As SvnDimensions
    Set pRangeAreas = New Collection
    pRepeatableAreas = False
    pCrossSheet = False
    Dim i As Long
    For i = 1 To Areas
        pRangeAreas.Add Array(Rows, Cols)
    Next
    Set FromRowsColsAreas = Me
End Function

