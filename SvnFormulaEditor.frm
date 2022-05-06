VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SvnFormulaEditor 
   Caption         =   "Formula Editor"
   ClientHeight    =   1170
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6270
   OleObjectBlob   =   "SvnFormulaEditor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SvnFormulaEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Sovereign")
Option Explicit

Private pFormulaText As String
Const ValidText As Long = &H8000&
Const InvalidText As Long = &HFF&

Public Property Get Formula() As String
    Formula = pFormulaText
End Property

Private Sub LoadForm()
    TextBoxEditor.Text = pFormulaText
End Sub

Private Sub UpdateLabels()
    Dim testVar As Variant
    testVar = varOrEmpty(TextBoxEditor.Text)
    If Not IsEmpty(testVar) Then
        LabelValid.Caption = "Valid. Evaluates to " & variantGetType(testVar) & " with value " & variantGetPretty(testVar)
        LabelValid.ForeColor = ValidText
    Else
        LabelValid.Caption = "Invalid Condition"
        LabelValid.ForeColor = InvalidText
    End If
    LabelCharacters.Caption = CStr(Len(TextBoxEditor.Text))
End Sub

Private Sub TextBoxEditor_Change()
    UpdateLabels
End Sub

Public Sub Initialize(ByVal inFormula As String, ByVal title As String)
    pFormulaText = inFormula
    Me.Caption = "Formula Edit: " & title
    LoadForm
    TextBoxEditor.SetFocus
End Sub

Private Sub CloseMe()
    TextBoxVoid.SetFocus
    Me.Hide
End Sub

Private Sub TextBoxEditor_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        pFormulaText = TextBoxEditor.Text
        CloseMe
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Cancel = True
    CloseMe
End Sub


