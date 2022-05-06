VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SvnRefEditor 
   Caption         =   "RefEdit"
   ClientHeight    =   1230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6315
   OleObjectBlob   =   "SvnRefEditor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SvnRefEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Sovereign")
Option Explicit

Private pRangeText As String
Private pWorkbook As Workbook
Const ValidText As Long = &H8000&
Const InvalidText As Long = &HFF&

Public Property Get Ref() As String
    Ref = pRangeText
End Property

Private Sub LoadForm()
    RefEditor.Text = pRangeText
End Sub

Private Sub UpdateLabels()
    Dim testRange As Range
    Set testRange = rangeOrNothing(RefEditor.Text, pWorkbook)
    If Not testRange Is Nothing Then
        LabelValid.Caption = "Valid with " & CStr(testRange.Areas.count) & " Areas"
        LabelValid.ForeColor = ValidText
    Else
        LabelValid.Caption = "Invalid Range"
        LabelValid.ForeColor = InvalidText
    End If
    LabelCharacters.Caption = CStr(Len(RefEditor.Text)) & "/255"
End Sub

Private Sub RefEditor_Change()
    UpdateLabels
End Sub

Private Sub RefEditor_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        pRangeText = RefEditor.Text
        CloseMe
    End If
End Sub

Public Sub Initialize(ByVal inWorkbook As Workbook, ByVal inRef As String, ByVal title As String)
    Set pWorkbook = inWorkbook
    pRangeText = inRef
    Me.Caption = "RefEdit: " & title
    LoadForm
End Sub

Private Sub CloseMe()
    TextBoxVoid.SetFocus
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Cancel = True
    CloseMe
End Sub

