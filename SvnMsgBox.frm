VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SvnMsgBox 
   Caption         =   "Message Box"
   ClientHeight    =   2445
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5385
   OleObjectBlob   =   "SvnMsgBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SvnMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Sovereign")
Option Explicit

Public Enum svnMessageType
    svnLeft
    svnCentre
    svnRight
    svnClose
End Enum

Private pReturnType As svnMessageType

Public Property Get ReturnType() As svnMessageType
    ReturnType = pReturnType
End Property

Public Sub Initialize(ByVal inTitle As String, ByVal inMessage As String, ByVal inButtonLeftText As String, ByVal inButtonCentreText As String, ByVal inButtonRightText As String)
    Me.Caption = inTitle
    Me.LabelMessage = inMessage
    Me.ButtonLeft.Caption = inButtonLeftText
    Me.ButtonCentre.Caption = inButtonCentreText
    Me.ButtonRight.Caption = inButtonRightText
    If inButtonLeftText = vbNullString Then Me.ButtonLeft.Visible = False
    If inButtonCentreText = vbNullString Then Me.ButtonCentre.Visible = False
    If inButtonRightText = vbNullString Then Me.ButtonRight.Visible = False
End Sub

Private Sub ButtonLeft_Click()
    pReturnType = svnLeft
    Me.Hide
End Sub

Private Sub ButtonCentre_Click()
    pReturnType = svnCentre
    Me.Hide
End Sub

Private Sub ButtonRight_Click()
    pReturnType = svnRight
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(ByRef Cancel As Integer, ByRef CloseMode As Integer)
    pReturnType = svnClose
    Me.Hide
    Cancel = True
End Sub



