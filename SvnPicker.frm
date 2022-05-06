VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SvnPicker 
   Caption         =   "Picker"
   ClientHeight    =   2175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6750
   OleObjectBlob   =   "SvnPicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SvnPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Sovereign")
Option Explicit

Private pickedIndex As Long
Private pickedText As String
Private pSuccess As Boolean

Public Property Get Index() As Long
    Index = pickedIndex
End Property

Public Property Get Text() As String
    Text = pickedText
End Property

Public Property Get Success() As Boolean
    Success = pSuccess
End Property

Private Sub ButtonCancel_Click()
    Me.Hide
End Sub

Private Sub ButtonSelect_Click()
    If ListNames.ListIndex <> -1 Then
        pickedIndex = ListNames.ListIndex + 1
        pickedText = ListNames.list(ListNames.ListIndex)
        pSuccess = True
        Me.Hide
    End If
End Sub

Public Sub Initialize(ByVal title As String, ByVal inputList As Collection)
    Me.Caption = title
    pSuccess = False
    Dim currentName As Variant
    For Each currentName In inputList
        ListNames.AddItem currentName
    Next
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ButtonCancel_Click
    Cancel = True
End Sub

