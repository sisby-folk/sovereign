VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SvnResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Sovereign")
Option Explicit

Public Enum SvnResultType
    svnSuccess
    svnSkip
    svnCancel
    SvnSubErr
End Enum

Private pType As SvnResultType
Private pMessage As String
Private pErrorDesc As String
Private pErrorNum As Long

Public Property Get ResultType() As SvnResultType
    ResultType = pType
End Property

Public Property Get Message() As String
    Message = pMessage
End Property

Public Property Get ErrorDesc() As String
    ErrorDesc = pErrorDesc
End Property

Public Property Get ErrorNum() As Long
    ErrorNum = pErrorNum
End Property

Public Function ShowResult(ByVal title As String) As Boolean
    If pMessage <> vbNullString Then MsgBox pMessage, vbOKOnly, title
    ShowResult = pType = svnSuccess
End Function

Public Function ComboResult(ByRef buildString As String, Optional ByVal Del As String = vbNewLine) As Boolean
    If (pType = svnCancel Or pType = SvnSubErr) Then MsgBox buildString & Del & pMessage, vbOKOnly, "Issue encountered"
    buildString = buildString & Del & pMessage
    ComboResult = pType = svnSuccess Or pType = svnSkip
End Function

Public Function assertionFails(ByVal Assertion As Boolean, ByVal inResult As SvnResultType, ByVal Message As String) As Boolean
    assertionFails = Not Assertion
    If assertionFails Then Instantiate inResult, Message
End Function

'' ** Signature Mirrored in SvnSovereign ** ''
Public Function Instantiate(ByVal inResult As SvnResultType, Optional ByVal inMessage As String, Optional ByVal inErrorNum As Long, Optional ByVal inErrorDesc As String) As SvnResult
    pType = inResult
    pMessage = inMessage
    pErrorNum = inErrorNum
    pErrorDesc = inErrorDesc
    Set Instantiate = Me
End Function

