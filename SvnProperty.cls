VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SvnProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Sovereign")
Option Explicit

' XXX: Make type sensitive

Private pSheet As Worksheet
Private pRange As SvnMultiRange
Private pType As MsoDocProperties
Private pName As String
Private pValue As Variant

Public Property Get sheet() As Worksheet
    Set sheet = pSheet
End Property

Public Property Get Name() As String
    Name = pName
End Property

Public Property Get value() As Variant
    value = pValue
End Property

Public Property Let value(ByVal inValue As Variant)
    On Error GoTo Handler
    If pType = msoPropertyTypeBoolean Then pValue = CBool(inValue)
    If pType = msoPropertyTypeNumber Then pValue = CLng(inValue)
    If pType = msoPropertyTypeString Then pValue = CStr(inValue)
    On Error GoTo 0
    Exit Property
Handler:
     Err.Raise 5000, Description:="Couldn't Coerce Property Value"
End Property

Public Property Get PropType() As MsoDocProperties
    PropType = pType
End Property

Public Property Get Range() As SvnMultiRange
    Set Range = pRange
End Property

'' ** Signature Mirrored in SvnSovereign ** ''
Public Function Instantiate(ByVal inName As String, ByVal inValue As Variant, ByVal inType As MsoDocProperties, Optional ByVal inSheet As Worksheet = Nothing, Optional ByVal inRange As SvnMultiRange = Nothing) As SvnProperty
    pName = inName
    pType = inType
    Me.value = inValue ' Calls Setter
    Set pSheet = inSheet
    If Not pSheet Is Nothing And (pType <> msoPropertyTypeString Or pValue = vbNullString) Then Err.Raise 5000, Description:="Non-String or Empty Worksheet Property"
    Set pRange = inRange
    Set Instantiate = Me
End Function

