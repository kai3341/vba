VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EPMTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'require Tools->References->"Microsoft VBScript Regular Expressions 5.5"
'require Tools->References->"Microsoft Scripting Runtime"
'require PivotTable
'require PivotTable:RangeTools

Option Explicit

Private Sheet As Worksheet

Private EPMMemberOParser As Scripting.Dictionary

Private R1C1RegEx As RegExp

Private DataRectangle As TRectangle

Private FormulaRectangleHead As TRectangle
Private FormulaRectangleLeft As TRectangle

Private Type TCell
    Row As Integer
    Col As Integer
End Type

Private Type TRectangle
    TopLeft As TCell
    BottomRight As TCell
End Type

'==================== PUBLIC METHODS ====================

Public Sub ConfigureEPMOLAPMember( _
    Optional DimensionName As Variant = ".*", _
    Optional HierarchyName As Variant = ".*", _
    Optional MemberID As Variant = ".*", _
    Optional ConnectionName As Variant = ".*", _
    Optional Description1 As Variant = ".*", _
    Optional Description2 As Variant = ".*", _
    Optional TableID As Variant = ".*" _
)

Dim EPMRegEx As New RegExp

EPMRegEx.Pattern = "EPMOlapMember[OC]?\(""\[(" & _
    DimensionName & ")\]\.\[(" & _
    HierarchyName & ")\]\.\[(" & _
    MemberID & ")\]""[;,]""(" & _
    ConnectionName & ")""[;,]""(" & _
    Description1 & ")""[;,]""(" & _
    Description2 & ")""[;,]""(" & _
    TableID & ")""\)"

EPMRegEx.IgnoreCase = False

If EPMMemberOParser.Exists("Olap") Then
    EPMMemberOParser.Remove "Olap"
End If

EPMMemberOParser.Add _
    Key:="Olap", _
    Item:=EPMRegEx

End Sub

Public Sub ConfigureEPMLocalMember( _
    Optional Description As Variant = ".*", _
    Optional MemberID As Variant = ".*", _
    Optional TableID As Variant = ".*" _
)
Dim EPMRegEx As New RegExp

EPMRegEx.Pattern = "EPMLocalMember[OC]?\(""(" & _
    Description & ")""[;,]""(" & _
    MemberID & ")""[;,]""(" & _
    TableID & ")""\)"

EPMRegEx.IgnoreCase = False

If EPMMemberOParser.Exists("Local") Then
    EPMMemberOParser.Remove "Local"
End If

EPMMemberOParser.Add _
    Key:="Local", _
    Item:=EPMRegEx
End Sub

Public Sub ConfigureEPMBlankMember( _
    Optional ConnectionName As Variant = ".*", _
    Optional Description1 As Variant = ".*", _
    Optional Description2 As Variant = ".*", _
    Optional TableID As Variant = ".*" _
)
Dim EPMRegEx As New RegExp

EPMRegEx.Pattern = "EPMOlapMember[OC]?\(""\[Blank Member\]""[;,]""(" & _
    ConnectionName & ")""[;,]""(" & _
    Description1 & ")""[;,]""(" & _
    Description2 & ")""[;,]""(" & _
    TableID & ")""\)"

EPMRegEx.IgnoreCase = False

If EPMMemberOParser.Exists("Blank") Then
    EPMMemberOParser.Remove "Blank"
End If

EPMMemberOParser.Add _
    Key:="Blank", _
    Item:=EPMRegEx
End Sub

Public Property Get FormulaRangeHead() As Range
Set FormulaRangeHead = Rectangle2Area(FormulaRectangleHead)
End Property

'Public Property Let FormulaRangeHead(NewVal As Range)
'FormulaRectangleHead = Area2Rectangle(NewVal)
'End Property

Public Property Get FormulaRangeLeft() As Range
Set FormulaRangeLeft = Rectangle2Area(FormulaRectangleLeft)
End Property

'Public Property Let FormulaRangeLeft(NewVal As Range)
'FormulaRectangleLeft = Area2Rectangle(NewVal)
'End Property

Public Property Let TopLeftDataCell(ByVal TheCell As Range)
Dim Row, Col As Integer
Dim BottomRightCell As Range

Set Sheet = TheCell.Worksheet

DataRectangle.TopLeft.Row = TheCell.Row
DataRectangle.TopLeft.Col = TheCell.column

FormulaRectangleHead.TopLeft.Col = TheCell.column
FormulaRectangleLeft.TopLeft.Row = TheCell.Row

Set BottomRightCell = Sheet.Cells.SpecialCells(xlCellTypeLastCell)

' This values will be redefined
' "Íî ýòî íå òî÷íî"
DataRectangle.BottomRight.Row = BottomRightCell.Row
DataRectangle.BottomRight.Col = BottomRightCell.column

FormulaRectangleHead.BottomRight.Col = BottomRightCell.column
FormulaRectangleLeft.BottomRight.Row = BottomRightCell.Row

Col = DataRectangle.TopLeft.Col     '<=== CONST

For Row = DataRectangle.TopLeft.Row To 1 Step -1
    Set TheCell = Sheet.Cells(Row, Col)
    If isValidEPMFormula(TheCell.Formula) Then
        Exit For
    End If
Next

FormulaRectangleHead.BottomRight.Row = Row

For Row = Row To 1 Step -1
    Set TheCell = Sheet.Cells(Row, Col)
    If isValidEPMFormula(TheCell.Formula) Then
        FormulaRectangleHead.TopLeft.Row = Row
    End If
Next

Row = DataRectangle.TopLeft.Row     '<=== CONST

For Col = DataRectangle.TopLeft.Col To 1 Step -1
    Set TheCell = Sheet.Cells(Row, Col)
    If isValidEPMFormula(TheCell.Formula) Then
        Exit For
    End If
Next

FormulaRectangleLeft.BottomRight.Col = Col

For Col = Col To 1 Step -1
    Set TheCell = Sheet.Cells(Row, Col)
    If isValidEPMFormula(TheCell.Formula) Then
        FormulaRectangleLeft.TopLeft.Col = Col
    End If
Next

Row = FormulaRectangleHead.BottomRight.Row      '<=== CONST

For Col = FormulaRectangleHead.TopLeft.Col To FormulaRectangleHead.BottomRight.Col
    Set TheCell = Sheet.Cells(Row, Col)
    If isValidEPMFormula(TheCell.Formula) Then
        DataRectangle.BottomRight.Col = Col
    End If
Next

FormulaRectangleHead.BottomRight.Col = DataRectangle.BottomRight.Col

Col = FormulaRectangleLeft.BottomRight.Col      '<=== CONST

For Row = FormulaRectangleLeft.TopLeft.Row To FormulaRectangleLeft.BottomRight.Row
    Set TheCell = Sheet.Cells(Row, Col)
    If isValidEPMFormula(TheCell.Formula) Then
        DataRectangle.BottomRight.Row = Row
    End If
Next

FormulaRectangleLeft.BottomRight.Row = DataRectangle.BottomRight.Row

'Set BottomRightCell = Rectangle2Area(FormulaRectangleHead)
'BottomRightCell.Select

'Set BottomRightCell = Rectangle2Area(FormulaRectangleLeft)
'BottomRightCell.Select

'Set BottomRightCell = Rectangle2Area(DataRectangle)
'BottomRightCell.Select
End Property

Public Property Get TopLeftDataCell() As Range
Set TopLeftDataCell = Sheet.Cells( _
    DataRectangle.TopLeft.Row, _
    DataRectangle.TopLeft.Col _
)
End Property

Public Property Let InnerDataRange(DataRange As Range)  'FIXME: NOT TESTED
Dim TopLeftCell As Range
Dim BottomRightCell As Range
Dim Row, Col As Integer

Set Sheet = DataRange.Worksheet


Set TopLeftCell = DataRange.Cells(1, 1)
Set BottomRightCell = DataRange.Cells.SpecialCells(xlCellTypeLastCell)

DataRectangle.TopLeft.Col = TopLeftCell.column
DataRectangle.TopLeft.Row = TopLeftCell.Row
DataRectangle.BottomRight.Col = BottomRightCell.column
DataRectangle.BottomRight.Row = BottomRightCell.Row

'Head
FormulaRectangleHead.TopLeft.Col = DataRectangle.TopLeft.Col
FormulaRectangleHead.BottomRight.Col = DataRectangle.BottomRight.Col

'Left
FormulaRectangleLeft.TopLeft.Row = DataRectangle.TopLeft.Row
FormulaRectangleLeft.BottomRight.Row = DataRectangle.BottomRight.Row

Col = DataRectangle.TopLeft.Col     '<=== CONST

For Row = DataRectangle.TopLeft.Row To 1 Step -1
    Set TheCell = Sheet.Cells(Row, Col)
    If isValidEPMFormula(TheCell.Formula) Then
        Exit For
    End If
Next

FormulaRectangleHead.BottomRight.Row = Row

For Row = Row To 1 Step -1
    Set TheCell = Sheet.Cells(Row, Col)
    If isValidEPMFormula(TheCell.Formula) Then
        FormulaRectangleHead.TopLeft.Row = Row
    End If
Next

Row = DataRectangle.TopLeft.Row     '<=== CONST

For Col = DataRectangle.TopLeft.Col To 1 Step -1
    Set TheCell = Sheet.Cells(Row, Col)
    If isValidEPMFormula(TheCell.Formula) Then
        Exit For
    End If
Next

FormulaRectangleLeft.BottomRight.Col = Col

For Col = Col To 1 Step -1
    Set TheCell = Sheet.Cells(Row, Col)
    If isValidEPMFormula(TheCell.Formula) Then
        FormulaRectangleLeft.TopLeft.Col = Col
    End If
Next
End Property

Public Property Get InnerDataRange() As Range
Set InnerDataRange = Rectangle2Area(DataRectangle)
End Property

'Public Property Let OuterTableRange(TableRange As Range)
'End Property

'Public Property Get OuterDataRange() As Range
'End Property

Public Property Get Dictionary() As Scripting.Dictionary
Set Dictionary = New Scripting.Dictionary

Dictionary.Add _
    Key:="Head", _
    Item:=Rectangle2Area(FormulaRectangleHead)

Dictionary.Add _
    Key:="Left", _
    Item:=Rectangle2Area(FormulaRectangleLeft)

Dictionary.Add _
    Key:="Data", _
    Item:=Rectangle2Area(DataRectangle)
End Property

Public Property Let Dictionary(NewVal As Scripting.Dictionary)
Dim DataRange As Range

Set DataRange = NewVal.Item("Data")

Set Sheet = DataRange.Worksheet

FormulaRectangleHead = Area2Rectangle(NewVal.Item("Head"))
FormulaRectangleLeft = Area2Rectangle(NewVal.Item("Left"))

DataRectangle = Area2Rectangle(DataRange)
End Property

Public Property Get PivotTable() As PivotTable
Set TargetTable = New PivotTable
ScanTable
Set PivotTable = TargetTable
End Property

Public Function ParseEPMFormula( _
    FormulaName As Variant, _
    line As String _
) As Object

Dim Matches As Object

If Not EPMMemberOParser.Exists(FormulaName) Then
    Err.Raise _
        Number:=12345, _
        Description:="EPM Formula """ & _
        FormulaName & """ does not exists."
End If

Set Matches = EPMMemberOParser.Item(FormulaName).Execute(line)

If Matches.Count = 0 Then
    Exit Function
End If

Set ParseEPMFormula = Matches.Item(0).SubMatches
End Function

Public Function isValidEPMFormula(line As String) As Boolean
Dim Key As Variant
Dim CurrentRegEx As RegExp

For Each Key In EPMMemberOParser.Keys
    Set CurrentRegEx = EPMMemberOParser(Key)
    isValidEPMFormula = CurrentRegEx.test(line)
    If isValidEPMFormula Then
        Exit Function
    End If
Next
End Function

'==================== PRIVATE METHODS ===================

Private Sub Class_Initialize()
Dim Patterns As Scripting.Dictionary

Set EPMMemberOParser = New Scripting.Dictionary

ConfigureEPMBlankMember
ConfigureEPMLocalMember
ConfigureEPMOLAPMember

Set R1C1RegEx = New RegExp

R1C1RegEx.Pattern = "R(\d+)C(\d+)\:R(\d+)C(\d+)"
R1C1RegEx.Global = True
R1C1RegEx.IgnoreCase = False

End Sub

Private Sub Class_Terminate()
    Set EPMMemberOParser = Nothing
    Set Sheet = Nothing
End Sub

Private Function Rectangle2Area(ThisRectangle As TRectangle) As Range
Dim TopLeftCell As Range
Dim BottomRightCell As Range

Set TopLeftCell = Sheet.Cells( _
    ThisRectangle.TopLeft.Row, _
    ThisRectangle.TopLeft.Col _
)

Set BottomRightCell = Sheet.Cells( _
    ThisRectangle.BottomRight.Row, _
    ThisRectangle.BottomRight.Col _
)

Set Rectangle2Area = Range(TopLeftCell, BottomRightCell)
End Function

Private Function Area2Rectangle(ThisArea As Range) As TRectangle
Dim Matches As Object

If Not ThisArea.Areas.Count = 1 Then
    Err.Raise _
        Number:=12345, _
        Description:="Only monolitic area can be converted to TRectangle unambigously"
End If

Set Matches = R1C1RegEx.Execute( _
    ThisArea.Address(ReferenceStyle:=xlR1C1) _
)

If Not Matches.Count = 1 Then
    Err.Raise _
        Number:=12345, _
        Description:="Invalid area R1C1 address. This exception could not occured"
End If

Set Matches = Matches.Item(0).SubMatches

Area2Rectangle.TopLeft.Row = CInt(Matches.Item(0))
Area2Rectangle.TopLeft.Col = CInt(Matches.Item(1))
Area2Rectangle.BottomRight.Row = CInt(Matches.Item(2))
Area2Rectangle.BottomRight.Col = CInt(Matches.Item(3))
End Function
                                                        
