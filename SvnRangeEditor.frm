VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SvnRangeEditor 
   Caption         =   "Range Editor"
   ClientHeight    =   3870
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8745.001
   OleObjectBlob   =   "SvnRangeEditor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SvnRangeEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@Folder("Sovereign")
Option Explicit

Private pWorkbook As Workbook
Private pCurrentRange As SvnMultiRange
Private pDimensions As SvnDimensions
Private pDirty As Boolean ' whether there have been changes
Private pValidRange As Boolean
Private pValidDimensions As Boolean

' Colour Constants
Const UnlockedText As Long = 0
Const LockedText As Long = 7895160
Const ValidText As Long = &H8000&
Const InvalidText As Long = &HFF&

Public Property Get Dirty() As Boolean
    Dirty = pDirty
End Property

Public Property Get Range() As SvnMultiRange
    Set Range = pCurrentRange
End Property

Public Property Get Union() As Range
    
End Property

Private Sub DirtyEditor(Optional ByVal doLoadForm As Boolean = True)
    pDirty = True
    FrameCurrentRange.Caption = "Edit Range*"
    If doLoadForm Then LoadForm
End Sub

Private Sub LoadForm()
    TextBoxRange.Text = pCurrentRange.ToString
    ValidateTextRange
End Sub

Private Sub LockControl(ByVal ctrl As Object, ByVal newLocked As Boolean)
    ctrl.Enabled = Not newLocked
    ctrl.ForeColor = IIf(newLocked, LockedText, UnlockedText)
End Sub

Private Sub ValidRange(ByVal valid As Boolean, ByVal testRange As SvnMultiRange, ByVal unionRange As SvnMultiRange)
    LabelValidRange.Caption = Len(testRange.ToString) & "/2048 - "
    pValidRange = valid And (Len(testRange.ToString) < 255 Or testRange.ToString = unionRange.ToString)
    If pValidRange Then
        Set pCurrentRange = testRange
        LabelValidRange.Caption = LabelValidRange.Caption & "Valid with " & CStr(pCurrentRange.Worksheets.count) & " Worksheets and " & CStr(pCurrentRange.Areas.count) & " Areas"
        LabelValidRange.ForeColor = ValidText
    Else
        LabelValidRange.Caption = LabelValidRange.Caption & "Invalid Range [Recover]" & IIf(valid, " Unions do not match.", "")
        LabelValidRange.ForeColor = InvalidText
    End If
End Sub

Private Sub ValidDimensions(ByVal valid As Boolean)
    If valid Then
            LabelValidDimensions.Caption = "Valid Dimensions for this object"
            LabelValidDimensions.ForeColor = ValidText
    Else
            LabelValidDimensions.Caption = "Invalid. Should be " & pDimensions.Pretty
            LabelValidDimensions.ForeColor = InvalidText
    End If
    pValidDimensions = valid
End Sub

Private Sub ValidateTextRange()
    Dim testRange As SvnMultiRange
    Set testRange = multiRangeFromInstantiate(pWorkbook, inUnionize:=pCurrentRange.Unionize, inEntireRows:=pCurrentRange.EntireRows, inSingleCell:=pCurrentRange.SingleCell)
    
    ValidRange testRange.AddRanges(TextBoxRange.Text) And Len(testRange.ToString) < 2048, testRange, multiRangeFromInstantiate(pWorkbook, rangeGetAddress(testRange.GetUnionRange), inUnionize:=pCurrentRange.Unionize, inEntireRows:=pCurrentRange.EntireRows, inSingleCell:=pCurrentRange.SingleCell)

    ValidDimensions pDimensions.Match(testRange)

    LockControl ButtonAdd, Not pValidRange
    LockControl ButtonOK, Not (pValidRange And pValidDimensions)
    LockControl ButtonExtend, Not (pValidRange And pCurrentRange.Worksheets.count = 1)
End Sub

Public Sub Initialize(ByVal inRange As SvnMultiRange, ByVal inWindowTitle As String, ByVal inWorkbook As Workbook, ByVal inDimensions As SvnDimensions)
    pDirty = False
    Set pCurrentRange = inRange
    Me.Caption = "Range Editor: " & inWindowTitle
    Set pWorkbook = inWorkbook
    Set pDimensions = inDimensions
    LoadForm
End Sub

Private Sub AddRange()
    Dim newRange As SvnMultiRange
    Set newRange = multiRangeFromRefedit(pWorkbook, "Enter Range", "Add Text", Dimensions:=Nothing, Unionize:=pCurrentRange.Unionize, EntireRows:=pCurrentRange.EntireRows, SingleCell:=pCurrentRange.SingleCell)
    If Not newRange Is Nothing Then
        pCurrentRange.Union newRange
        DirtyEditor
    End If
End Sub


Private Sub ButtonAdd_Click()
    Me.Hide
        AddRange
    Me.Show
End Sub

Private Sub ButtonClear_Click()
    Set pCurrentRange = multiRangeFromInstantiate(pWorkbook, inUnionize:=pCurrentRange.Unionize, inEntireRows:=pCurrentRange.EntireRows, inSingleCell:=pCurrentRange.SingleCell)
    DirtyEditor
End Sub

Private Sub ButtonExtend_Click()
    Dim currentSheetAreas As Collection
    Set currentSheetAreas = pCurrentRange.Areas

    Dim equidistance As Long
    equidistance = currentSheetAreas(2).row - currentSheetAreas(1).row
    
    Dim newRange As SvnMultiRange
    Set newRange = pCurrentRange.Offset(equidistance * (currentSheetAreas.count), 0)
    
    pCurrentRange.AddRanges newRange.ToString
    DirtyEditor
End Sub

Private Sub ButtonOK_Click()
    Me.Hide
End Sub

Private Sub ButtonCancel_Click()
    pDirty = False
    Me.Hide
End Sub

Private Sub LabelValidRange_Click()
    LoadForm
End Sub

Private Sub TextBoxRange_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ValidateTextRange
    DirtyEditor False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ButtonCancel_Click
    Cancel = True
End Sub



