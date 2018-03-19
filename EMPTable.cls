VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EMPTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'require Tools->References->"Microsoft VBScript Regular Expressions 5.5"
'require Tools->References->"Microsoft Scripting Runtime"
'require PivotTable
'require PivotTable:RangeTools

Option Explicit

Private TargetTable As PivotTable

Private Sheet As Worksheet

Private EPMTableId As Variant

Private EPMMemberOParser As Scripting.Dictionary

Private R1C1RegEx As RegExp

Private DataRectangle As TRectangle

Private FormulaCount As TCell

Private Type TEPMOlapMemberO
    'MemberType As TEPMMemberType
    Valid As Boolean
    DimestionName As String
    MemberId As String
End Type

Private Enum TEPMMemberType
    EMPOlap = 0
    EPMLocal = 1
    EPMBlank = 2
End Enum

Private Type TCell
    Row As Integer
    Col As Integer
End Type

Private Type TRectangle
    Start As TCell
    End As TCell
End Type

Private Sub Class_Initialize()
Dim EPMRegEx As RegExp
Dim Patterns As Scripting.Dictionary
Dim PatternName As Variant

Set EPMMemberOParser = New Scripting.Dictionary
Set Patterns = New Scripting.Dictionary


'Patterns.Add key:="Olap", Item:="EPMOlapMemberO\(""\[([\w\.]+)\]\.\[([\w\.]*)\]\.\[([\w\.]+)\]"";"".*"";"".*"";(\d{3})"
Patterns.Add key:="Olap", Item:="EPMOlapMemberO\(""\[([\w\.]+)\]\.\[([\w\.]*)\]\.\[([\w\.]+)\]"
Patterns.Add key:="Local", Item:="EPMLocalMember(O){0,1}\("""
Patterns.Add key:="Blank", Item:="EPMOlapMemberO\(""\[Blank Member\]"""


For Each PatternName In Patterns.Keys
    Set EPMRegEx = New RegExp
    EPMRegEx.Pattern = Patterns(PatternName)
    EPMRegEx.IgnoreCase = False
    
    EPMMemberOParser.Add _
        key:=PatternName, _
        Item:=EPMRegEx
Next

Set R1C1RegEx = New RegExp
R1C1RegEx.Pattern = "R(\d+)C(\d+):R(\d+)C(\d+)"

End Sub

Private Sub Class_Terminate()
    Set EPMMemberOParser = Nothing
    Set Sheet = Nothing
End Sub

Public Property Get TableID() As Variant
TableID = EPMTableId
End Property

Public Property Let TableID(TableID As Variant)
Dim EPMRegEx As New RegExp
Dim PatternName As Variant

PatternName = "Olap"

EPMTableId = TableID

EPMRegEx.Pattern = "EPMOlapMemberO\(""\[([\w\.]+)\]\.\[([\w\.]*)\]\.\[" & TableID & "\]"""

Debug.Print EPMRegEx.Pattern

EPMRegEx.IgnoreCase = False

EPMMemberOParser.Remove PatternName

EPMMemberOParser.Add _
    key:=PatternName, _
    Item:=EPMRegEx
End Property

Private Function isValidEPMFormula(line As String) As Boolean
Dim key As Variant
Dim AllMatches As Object
Dim CurrentRegEx As RegExp

For Each key In EPMMemberOParser.Keys
    Set CurrentRegEx = EPMMemberOParser(key)
    Set AllMatches = CurrentRegEx.Execute(line)
    isValidEPMFormula = (AllMatches.Count > 0)
    If isValidEPMFormula Then
        Exit Function
    End If
Next
End Function

Private Function ScanTableByCols() As Integer
Dim _
    StartFormulaRow, _
    EndFormulaRow, _
    FormulaRow _
As Integer
Dim ThisRange As Range
Dim ThisFormulaCell As Range
Dim EPMOlapMember As TEPMOlapMemberO

StartFormulaRow = DataRectangle.Start.Row - 1
EndFormulaRow = DataRectangle.Start.Row - FormulaCount.Row

For ScanTableByCols = DataRectangle.Start.Col To DataRectangle.End.Col
    Set ThisRange = Range( _
        Sheet.Cells(DataRectangle.Start.Row, ScanTableByCols), _
        Sheet.Cells(DataRectangle.End.Row, ScanTableByCols) _
    )

    For FormulaRow = StartFormulaRow To EndFormulaRow Step -1
        Set ThisFormulaCell = Sheet.Cells(FormulaRow, ScanTableByCols)
        
        EPMOlapMember = EPMOlapFormula(ThisFormulaCell.formula)
        If EPMOlapMember.Valid Then
            TargetTable.Add _
                DimensionName:=EPMOlapMember.DimestionName, _
                MemberId:=EPMOlapMember.MemberId, _
                ItemRange:=ThisRange
        End If
    Next FormulaRow
Next ScanTableByCols

ScanTableByCols = ScanTableByCols - 1
End Function

Private Function ScanTableByRows() As Integer
Dim _
    StartFormulaCol, _
    EndFormulaCol, _
    FormulaCol _
As Integer
Dim ThisRange As Range
Dim ThisFormulaCell As Range
Dim EPMOlapMember As TEPMOlapMemberO

StartFormulaCol = DataRectangle.Start.Col - 1
EndFormulaCol = DataRectangle.Start.Col - FormulaCount.Col

For ScanTableByRows = DataRectangle.Start.Row To DataRectangle.End.Row
    Set ThisRange = Range( _
        Sheet.Cells(ScanTableByRows, DataRectangle.Start.Col), _
        Sheet.Cells(ScanTableByRows, DataRectangle.End.Col) _
    )

    For FormulaCol = StartFormulaCol To EndFormulaCol Step -1
        Set ThisFormulaCell = Sheet.Cells(ScanTableByRows, FormulaCol)
        
        EPMOlapMember = EPMOlapFormula(ThisFormulaCell.formula)
        If EPMOlapMember.Valid Then
            TargetTable.Add _
                DimensionName:=EPMOlapMember.DimestionName, _
                MemberId:=EPMOlapMember.MemberId, _
                ItemRange:=ThisRange
        End If
    Next FormulaCol
Next ScanTableByRows

ScanTableByRows = ScanTableByRows - 1
End Function

Private Sub ScanTable()
'FIXME: Заменить порядок сканирования на Z-стратегию
'(если это возможно)
Call ScanTableByCols
Call ScanTableByRows
End Sub

Private Sub CalcOLAPFormulaInnerLeft()
Dim Col As Integer
Dim ScanRow As Integer
Dim CurrentCell As Range

ScanRow = DataRectangle.Start.Row

For Col = DataRectangle.Start.Col - 1 To 1 Step -1
    Set CurrentCell = Sheet.Cells( _
        ScanRow, _
        Col _
    )
    
    If Not isValidEPMFormula(CurrentCell.formula) Then
        'CurrentCell.Select
        Exit For
    End If
Next Col

FormulaCount.Col = DataRectangle.Start.Col - Col - 1
End Sub

Private Sub CalcOLAPFormulaInnerHead()
Dim Row As Integer
Dim ScanCol As Integer
Dim CurrentCell As Range

ScanCol = DataRectangle.Start.Col

For Row = DataRectangle.Start.Row - 1 To 1 Step -1
    Set CurrentCell = Sheet.Cells( _
        Row, _
        ScanCol _
    )
    
    If Not isValidEPMFormula(CurrentCell.formula) Then
        'CurrentCell.Select
        Exit For
    End If
Next Row

FormulaCount.Row = DataRectangle.Start.Row - Row - 1
End Sub

Private Sub CalcOLAPFormulaInner()
'V[e]+ry stupid scan method
CalcOLAPFormulaInnerHead
CalcOLAPFormulaInnerLeft
End Sub

Private Sub CalcOLAPFormulaOuterRight( _
    FirstOuterCell As Range, _
    LastCell As Range _
)

Dim Col As Integer
Dim ScanRow As Integer
Dim CurrentCell As Range

ScanRow = FirstOuterCell.Row

For Col = FirstOuterCell.column To LastCell.column
    Set CurrentCell = Sheet.Cells( _
        ScanRow, _
        Col _
    )
    
    If isValidEPMFormula(CurrentCell.formula) Then
        'CurrentCell.Select
        Exit For
    End If
Next Col

DataRectangle.Start.Col = Col
FormulaCount.Col = Col - FirstOuterCell.column
End Sub

Private Sub CalcOLAPFormulaOuterBottom( _
    FirstOuterCell As Range, _
    LastCell As Range _
)

Dim Row As Integer
Dim ScanCol As Integer
Dim CurrentCell As Range

ScanCol = FirstOuterCell.column

For Row = FirstOuterCell.Row To LastCell.Row
    Set CurrentCell = Sheet.Cells( _
        Row, _
        ScanCol _
    )
    
    If isValidEPMFormula(CurrentCell.formula) Then
        'CurrentCell.Select
        Exit For
    End If
Next Row

DataRectangle.Start.Row = Row
FormulaCount.Row = Row - FirstOuterCell.Row
End Sub

Private Sub CalcOLAPFormulaOuter( _
    FirstOuterCell As Range, _
    LastCell As Range _
)

'V[e]+ry stupid scan method
'FirstOuterCell.Select
Call CalcOLAPFormulaOuterRight(FirstOuterCell, LastCell)
Call CalcOLAPFormulaOuterBottom(FirstOuterCell, LastCell)
End Sub

Private Sub CalcPivotTableLengthRight(LastCell As Range)
Dim Col As Integer
Dim ScanRow As Integer
Dim CurrentCell As Range

If DataRectangle.Start.Row = 1 Then
    DataRectangle.End.Col = DataRectangle.Start.Col
    Exit Sub
End If

ScanRow = DataRectangle.Start.Row - 1

For Col = DataRectangle.Start.Col To LastCell.column
    Set CurrentCell = Sheet.Cells( _
        ScanRow, _
        Col _
    )
    
    If Not isValidEPMFormula(CurrentCell.formula) Then
        'CurrentCell.Select
        Exit For
    End If
Next Col

DataRectangle.End.Col = Col - 1

End Sub

Private Sub CalcPivotTableLengthBottom(LastCell As Range)
Dim Row As Integer
Dim ScanCol As Integer
Dim CurrentCell As Range

If DataRectangle.Start.Col = 1 Then
    DataRectangle.End.Row = DataRectangle.Start.Row
    Exit Sub
End If

ScanCol = DataRectangle.Start.Col - 1

For Row = DataRectangle.Start.Row To LastCell.Row
    Set CurrentCell = Sheet.Cells( _
        Row, _
        ScanCol _
    )
    If Not isValidEPMFormula(CurrentCell.formula) Then
        'CurrentCell.Select
        Exit For
    End If
Next Row

DataRectangle.End.Row = Row - 1

End Sub

Private Sub CalcPivotTableLength(LastCell As Range)
'V[e]+ry stupid scan method
Call CalcPivotTableLengthRight(LastCell)
Call CalcPivotTableLengthBottom(LastCell)
End Sub

Private Function EPMOlapFormula(line As String) As TEPMOlapMemberO
Dim Matches As Object
Dim MatchItems As Object

Set Matches = EPMMemberOParser("Olap").Execute(line)

EPMOlapFormula.Valid = (Matches.Count > 0)

If Not EPMOlapFormula.Valid Then
    Exit Function
End If

Set MatchItems = Matches.Item(0).SubMatches

EPMOlapFormula.Valid = (MatchItems.Count > 2)

If Not EPMOlapFormula.Valid Then
    Exit Function
End If

EPMOlapFormula.DimestionName = MatchItems.Item(0)
EPMOlapFormula.MemberId = MatchItems.Item(MatchItems.Count - 1)
End Function

Public Function fromFirstDataCell(FirstDataCell As Range) As PivotTable
Dim LastCell As Range

Set Sheet = FirstDataCell.Worksheet
Set LastCell = Sheet.Cells.SpecialCells(xlLastCell)
Set TargetTable = New PivotTable

DataRectangle.Start.Row = FirstDataCell.Row
DataRectangle.Start.Col = FirstDataCell.column
Call CalcPivotTableLength(LastCell)
CalcOLAPFormulaInner

ScanTable

Set fromFirstDataCell = TargetTable

End Function

Public Function fromInnerDataRange(DataRange As Range) As PivotTable
'FIXME: Maybe works INCORRECTLY!
Dim LastCell As Range

Set Sheet = DataRange.Worksheet
Set LastCell = DataRange.SpecialCells(xlLastCell)
Set TargetTable = New PivotTable

DataRectangle.Start.Row = DataRange.Cells(1, 1).Row
DataRectangle.Start.Col = DataRange.Cells(1, 1).column
DataRectangle.End.Row = LastCell.Row
DataRectangle.End.Col = LastCell.column

CalcOLAPFormulaInner

ScanTable

Set fromInnerDataRange = TargetTable

End Function

Public Function fromOuterTableRange(TableRange As Range) As PivotTable
'FIXME: Works ugly.

Dim LastCell As Range

Dim RangeAddress As String
Dim Matches As Object

Dim StartRangeCell As TCell
Dim EndRangeCell As TCell

Set Sheet = TableRange.Worksheet

Set TargetTable = New PivotTable

RangeAddress = TableRange.Address(ReferenceStyle:=xlR1C1)

Set Matches = R1C1RegEx.Execute(RangeAddress)

If Matches.Count <> 1 Then
    Err.Raise _
        Number:=12345, _
        Description:="Range Address """ & RangeAddress & """ looks invalid"
End If

Set Matches = Matches.Item(0).SubMatches

If Matches.Count <> 4 Then
    Err.Raise _
        Number:=12345, _
        Description:="Range Address """ & RangeAddress & """ looks invalid"
End If

StartRangeCell.Row = Matches.Item(0)
StartRangeCell.Col = Matches.Item(1)

DataRectangle.End.Row = Matches.Item(2)
DataRectangle.End.Col = Matches.Item(3)

Set LastCell = Sheet.Cells( _
    DataRectangle.End.Row, _
    DataRectangle.End.Col _
)

Call CalcOLAPFormulaOuter(TableRange.Cells(1, 1), LastCell)
ScanTable

Set fromOuterTableRange = TargetTable

End Function

'Public Function fromAnalyzerDynReport(AnalyzerDynReport As Object) As PivotTable
'End Function
