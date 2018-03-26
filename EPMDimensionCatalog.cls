VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EPMDimensionCatalog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'require Tools->References->"Microsoft Scripting Runtime"
'require EPMDimension

Option Explicit

Private DimensionDict As Scripting.Dictionary

'================PUBLIC API=====================

Public Sub Sheets(ParamArray SheetNames() As Variant)
Dim SheetName As Variant
Dim ThisEPMDimension As EPMDimension

Set DimensionDict = New Scripting.Dictionary

For Each SheetName In SheetNames
    Set ThisEPMDimension = New EPMDimension
    ThisEPMDimension.Sheet = SheetName

    DimensionDict.Add _
        key:=SheetName, _
        Item:=ThisEPMDimension
Next
End Sub

Public Function ValueFirst( _
    DimensionName As Variant, _
    LookupColumn As Variant, _
    LookupValue As Variant, _
    ResoultColumn As Variant _
) As Variant
Dim EPMDimension As EPMDimension

If Not DimensionDict.Exists(DimensionName) Then
    Err.Raise _
        Number:=12345, _
        Description:="Dimension """ & DimensionName & _
            """ not found in this EPMDimensionCatalog"
    Exit Function
End If

Set EPMDimension = DimensionDict.Item(DimensionName)

ValueFirst = EPMDimension.FindFirstValue( _
    LookupColumn, _
    LookupValue, _
    ResoultColumn _
)
End Function

Public Function SaveValueFirst( _
    DimensionName As Variant, _
    LookupColumn As Variant, _
    LookupValue As Variant, _
    ResoultColumn As Variant _
) As Variant

On Error GoTo NullExit

SaveValueFirst = ValueFirst( _
    DimensionName, _
    LookupColumn, _
    LookupValue, _
    ResoultColumn _
)

Exit Function

NullExit:
SaveValueFirst = Null
End Function

'==============PRIVATE METHODS==================

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
'Set DimensionDict = Nothing
End Sub
