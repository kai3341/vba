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

Private TargetTable As PivotTable

Private Sheet As Worksheet

Private EPMTableId As Variant

Private EPMMemberOParser As Scripting.Dictionary

Private R1C1RegEx As RegExp

Private DataRectangle As TRectangle

Private FormulaCount As TCell

Private Type TCell
    Row As Integer
    Col As Integer
End Type

Private Type TRectangle
    Start As TCell
    End As TCell
End Type

'==================== PUBLIC METHODS ====================

Public Sub SetEPMOLAPMemberParser( _
    Optional DimensionName As Variant = ".*", _
    Optional HierarchyName As Variant = ".*", _
    Optional MemberID As Variant = ".*", _
    Optional ConnectionName As Variant = ".*", _
    Optional Description1 As Variant = ".*", _
    Optional Description2 As Variant = ".*", _
    Optional TableID As Variant = ".*" _
)

Dim EPMRegEx As New RegExp

EPMRegEx.Pattern = "EPMOlapMember[OC]\(""\[(" & _
        DimensionName & ")\]\.\[(" & _
        HierarchyName & ")\]\.\[(" & _
        MemberID & ")\]""[;,]""(" & _
        ConnectionName & ")""[;,]""(" & _
        Description1 & ")""[;,]""(" & _
        Description2 & ")""[;,]""(" & _
        TableID & ")""\)\)"

EPMRegEx.IgnoreCase = False

If EPMMemberOParser.Exists("Olap") Then
    EPMMemberOParser.Remove "Olap"
End If

EPMMemberOParser.Add _
    key:="Olap", _
    Item:=EPMRegEx

End Sub

Public Function FormulaRange() As Range
Dim Head As Range
Dim Left As Range
Dim TopLeftCell As Range
Dim BottomRightCell As Range

Set TopLeftCell = Sheet.Cells( _
    DataRectangle.Start.Row - FormulaCount.Row - 1, _
    DataRectangle.Start.Col _
)

Set BottomRightCell = Sheet.Cells( _
    DataRectangle.Start.Row - 1, _
    DataRectangle.End.Col _
)

Set Head = Range(TopLeftCell, BottomRightCell)

Set TopLeftCell = Sheet.Cells( _
    DataRectangle.Start.Row, _
    DataRectangle.Start.Col - FormulaCount.Col - 1 _
)

Set BottomRightCell = Sheet.Cells( _
    DataRectangle.End.Row, _
    DataRectangle.Start.Col - 1 _
)

Set Left = Range(TopLeftCell, BottomRightCell)

Set FormulaRange = Application.Union(Head, Left)
End Function

Public Property Let FirstDataCell(FirstDataCell As Range)
Dim LastCell As Range

Set Sheet = FirstDataCell.Worksheet
Set LastCell = Sheet.Cells.SpecialCells(xlLastCell)

DataRectangle.Start.Row = FirstDataCell.Row
DataRectangle.Start.Col = FirstDataCell.column
Call CalcPivotTableLength(LastCell)
CalcOLAPFormulaInner
End Property

Public Property Get FirstDataCell() As Range
Set FirstDataCell = Sheet.Cells( _
    DataRectangle.Start.Row, _
    DataRectangle.Start.Col _
)
End Property

Public Property Let InnerDataRange(DataRange As Range)
'FIXME: Maybe works INCORRECTLY!
Dim LastCell As Range

Set Sheet = DataRange.Worksheet
Set LastCell = DataRange.SpecialCells(xlLastCell)

DataRectangle.Start.Row = DataRange.Cells(1, 1).Row
DataRectangle.Start.Col = DataRange.Cells(1, 1).column
DataRectangle.End.Row = LastCell.Row
DataRectangle.End.Col = LastCell.column

CalcOLAPFormulaInner
End Property

Public Property Get InnerDataRange() As Range
Dim TopLeftCell As Range
Dim BottomRightCell As Range

Set TopLeftCell = Sheet.Cells( _
    DataRectangle.Start.Row, _
    DataRectangle.Start.Col _
)

Set BottomRightCell = Sheet.Cells( _
    DataRectangle.End.Row, _
    DataRectangle.End.Col _
)

Set InnerDataRange = Range(TopLeftCell, BottomRightCell)
End Property

Public Property Let OuterTableRange(TableRange As Range)
'FIXME: Works ugly.

Dim LastCell As Range

Dim RangeAddress As String
Dim Matches As Object

Dim StartRangeCell As TCell
Dim EndRangeCell As TCell

Set Sheet = TableRange.Worksheet

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
End Property

Public Property Get OuterDataRange() As Range
Dim TopLeftCell As Range
Dim BottomRightCell As Range

Set TopLeftCell = Sheet.Cells( _
    DataRectangle.Start.Row - FormulaCount.Row, _
    DataRectangle.Start.Col - FormulaCount.Col _
)

Set BottomRightCell = Sheet.Cells( _
    DataRectangle.End.Row, _
    DataRectangle.End.Col _
)

Set InnerDataRange = Range(TopLeftCell, BottomRightCell)
End Property

Public Property Get PivotTable() As PivotTable
Set TargetTable = New PivotTable
ScanTable
Set PivotTable = TargetTable
End Property

Public Function ParseEPMOlapFormula(line As String) As Object
Dim Matches As Object
Dim MatchItems As Object

Set Matches = EPMMemberOParser("Olap").Execute(line)

If Matches.Count = 0 Then
    Exit Function
End If

Set ParseEPMOlapFormula = Matches.Item(0).SubMatches
End Function

Public Function isValidEPMFormula(line As String) As Boolean
Dim key As Variant
Dim CurrentRegEx As RegExp

For Each key In EPMMemberOParser.Keys
    Set CurrentRegEx = EPMMemberOParser(key)
    isValidEPMFormula = CurrentRegEx.Test(line)
    If isValidEPMFormula Then
        Exit Function
    End If
Next
End Function

'==================== PRIVATE METHODS ===================

Private Sub Class_Initialize()
Dim EPMRegEx As RegExp
Dim Patterns As Scripting.Dictionary
Dim PatternName As Variant

Set EPMMemberOParser = New Scripting.Dictionary
Set Patterns = New Scripting.Dictionary

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

SetEPMOLAPMemberParser

Set R1C1RegEx = New RegExp
R1C1RegEx.Pattern = "R(\d+)C(\d+):R(\d+)C(\d+)"

End Sub

Private Sub Class_Terminate()
    Set EPMMemberOParser = Nothing
    Set Sheet = Nothing
End Sub

'==== Здесь и далее идут костыли ====
'==== Перепишу, как совсем допечет. Работает -- и ладно :)

Private Function ScanTableByCols() As Integer
Dim _
    StartFormulaRow, _
    EndFormulaRow, _
    FormulaRow _
As Integer
Dim ThisRange As Range
Dim ThisFormulaCell As Range
Dim EPMOlapMember As Object

StartFormulaRow = DataRectangle.Start.Row - 1
EndFormulaRow = DataRectangle.Start.Row - FormulaCount.Row

For ScanTableByCols = DataRectangle.Start.Col To DataRectangle.End.Col
    Set ThisRange = Range( _
        Sheet.Cells(DataRectangle.Start.Row, ScanTableByCols), _
        Sheet.Cells(DataRectangle.End.Row, ScanTableByCols) _
    )

    For FormulaRow = StartFormulaRow To EndFormulaRow Step -1
        Set ThisFormulaCell = Sheet.Cells(FormulaRow, ScanTableByCols)
        
        Set EPMOlapMember = ParseEPMOlapFormula(ThisFormulaCell.formula)
        If Not EPMOlapMember Is Nothing Then
            TargetTable.Add _
                DimensionName:=EPMOlapMember.Item(0), _
                MemberID:=EPMOlapMember.Item(2), _
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
Dim EPMOlapMember As Object

StartFormulaCol = DataRectangle.Start.Col - 1
EndFormulaCol = DataRectangle.Start.Col - FormulaCount.Col

For ScanTableByRows = DataRectangle.Start.Row To DataRectangle.End.Row
    Set ThisRange = Range( _
        Sheet.Cells(ScanTableByRows, DataRectangle.Start.Col), _
        Sheet.Cells(ScanTableByRows, DataRectangle.End.Col) _
    )

    For FormulaCol = StartFormulaCol To EndFormulaCol Step -1
        Set ThisFormulaCell = Sheet.Cells(ScanTableByRows, FormulaCol)
        
        Set EPMOlapMember = ParseEPMOlapFormula(ThisFormulaCell.formula)
        If Not EPMOlapMember Is Nothing Then
            TargetTable.Add _
                DimensionName:=EPMOlapMember.Item(0), _
                MemberID:=EPMOlapMember.Item(2), _
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
