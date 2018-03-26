VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EPMDimension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'require Tools->References->"Microsoft VBScript Regular Expressions 5.5"
'require Tools->References->"Microsoft Scripting Runtime"

Option Explicit

Private Columns As Scripting.Dictionary
Private ColumnRanges As Scripting.Dictionary

Private EPMOLAPMember_re As RegExp
Private DimensionSheet As Worksheet
Private TopLeftCell As Range
Private BottomRightCell As Range

'================PUBLIC API=====================

Public Property Let Sheet(NewVal As Variant)
Set Columns = New Scripting.Dictionary
Set ColumnRanges = New Scripting.Dictionary
Set DimensionSheet = ThisWorkbook.Sheets(NewVal)

Set TopLeftCell = DimensionSheet.Cells(1, 1)  'Not so hardcode. Can be redefined.
Set BottomRightCell = DimensionSheet.Cells.SpecialCells(xlLastCell)
Call SheetBorders
Call MakeColRanges
End Property

Public Property Get Sheet() As Variant
Sheet = DimensionSheet.Name
End Property

Public Function FindFirstValue( _
    LookupColumnName As Variant, _
    LookupValue As Variant, _
    ResoultColumnName As Variant _
) As Variant

Dim RowIndex As Integer
Dim LookupVector As Range
Dim ResoultColumn As Integer
Dim FoundCell As Range
Dim ResoultCell As Range

If Not ColumnRanges.Exists(LookupColumnName) Then
    Err.Raise _
        Number:=12345, _
        Description:="Column """ & LookupColumnName & _
        """ not found in this EMPDimension"
End If

Set LookupVector = ColumnRanges.Item(LookupColumnName)

Set FoundCell = LookupVector.Find(LookupValue)

If FoundCell Is Nothing Then
    FindFirstValue = Null
    Exit Function
End If

RowIndex = FoundCell.Row

ResoultColumn = Columns.Item(ResoultColumnName)

Set ResoultCell = DimensionSheet.Cells( _
    RowIndex, _
    ResoultColumn _
)

FindFirstValue = ResoultCell.Value
End Function

'==============PRIVATE METHODS==================

Private Sub Class_Initialize()
Set EPMOLAPMember_re = New RegExp
EPMOLAPMember_re.Pattern = "EPMOlapMemberO\("
End Sub

Private Sub Class_Terminate()
Set Columns = Nothing
Set ColumnRanges = Nothing
Set EPMOLAPMember_re = Nothing
End Sub

Private Sub SheetBorders()
Dim TheCell As Range
Dim Col, Row As Integer

'Determine BottomRightCell more correctly
Dim RightCol, BottomRow As Integer

Row = TopLeftCell.Row

For Col = TopLeftCell.column To BottomRightCell.column
    Set TheCell = DimensionSheet.Cells(Row, Col)
    
    If EPMOLAPMember_re.Test(TheCell.formula) Then
        Exit For
    End If
    
    Columns.Add _
        key:=TheCell.Value, _
        Item:=Col
Next Col

RightCol = Col - 1
Col = TopLeftCell.column

For Row = TopLeftCell.Row + 1 To BottomRightCell.Row
    Set TheCell = DimensionSheet.Cells(Row, Col)

    If Not EPMOLAPMember_re.Test(TheCell.formula) Then
        Exit For
    End If
Next Row

BottomRow = Row - 1

Set BottomRightCell = DimensionSheet.Cells( _
    BottomRow, _
    RightCol _
)
End Sub

Private Sub MakeColRanges()
Dim ThisColumnName As Variant
Dim ThisColumnRange As Range
Dim Col As Integer

Dim TopRow, BottomRow As Integer

TopRow = TopLeftCell.Row + 1
BottomRow = BottomRightCell.Row

For Each ThisColumnName In Columns
    Col = Columns.Item(ThisColumnName)

    Set ThisColumnRange = Range( _
        DimensionSheet.Cells(TopRow, Col), _
        DimensionSheet.Cells(BottomRow, Col) _
    )

    ColumnRanges.Add _
        key:=ThisColumnName, _
        Item:=ThisColumnRange
Next
End Sub
