Attribute VB_Name = "RangeTools"
Option Explicit

Function VVectorExtract( _
    input_range As Range, _
    Row As Integer _
) As Range
    Set VVectorExtract = VSliceExtract( _
        input_range, _
        Row, _
        Row _
    )
End Function

Function HVectorExtract( _
    input_range As Range, _
    column As Integer _
) As Range
    Set HVectorExtract = HSliceExtract( _
        input_range, _
        column, _
        column _
    )
End Function

Function VSliceExtract( _
    input_range As Range, _
    start_row As Integer, _
    end_row As Integer _
) As Range
'VSliceExtract: Range -> Int -> Int -> Range

Dim m1, m2 As Range
Dim Mask As Range

Dim LastCell As Range

Set LastCell = input_range.Cells.SpecialCells(xlLastCell)

If _
    start_row = input_range.Cells(1, 1).Row And _
    end_row = LastCell.Row _
Then
    VSliceExtract = Nothing
    Exit Function

ElseIf start_row = input_range.Cells(1, 1).Row Then
    Set Mask = Range( _
        input_range.Worksheet.Cells(end_row + 1, 1), _
        input_range.Worksheet.Cells( _
            LastCell.Row, _
            LastCell.column _
        ) _
    )
ElseIf end_row = LastCell.Row Then
    Set Mask = Range( _
        input_range.Worksheet.Cells(1, 1), _
        input_range.Worksheet.Cells( _
            start_row - 1, _
            1 _
        ) _
    )
Else
    Set m1 = Range( _
        input_range.Worksheet.Cells(1, 1), _
        input_range.Worksheet.Cells( _
            start_row - 1, _
            LastCell.column _
        ) _
    )
    
    Set m2 = Range( _
        input_range.Worksheet.Cells(1, end_column + 1), _
        input_range.Worksheet.Cells( _
            LastCell.Row, _
            LastCell.column _
        ) _
    )
    
    Set Mask = Application.Union(m1, m2)
End If

'mask.Select

Set VSliceExtract = Application.Intersect( _
    input_range, _
    Mask _
)

End Function

Function HSliceExtract( _
    input_range As Range, _
    start_column As Integer, _
    end_column As Integer _
) As Range
'HSliceExtract: Range -> Int -> Int -> Range

Dim m1, m2 As Range
Dim Mask As Range

Dim LastCell As Range

Set LastCell = input_range.Cells.SpecialCells(xlLastCell)

If _
    start_column = input_range.Cells(1, 1).column And _
    end_column = LastCell.column _
Then
    HSliceExtract = Nothing
    Exit Function

ElseIf start_column = input_range.Cells(1, 1).column Then
    Set Mask = Range( _
        input_range.Worksheet.Cells(1, end_column + 1), _
        input_range.Worksheet.Cells( _
            LastCell.Row, _
            LastCell.column _
        ) _
    )
ElseIf end_column = LastCell.column Then
    Set Mask = Range( _
        input_range.Worksheet.Cells(1, 1), _
        input_range.Worksheet.Cells( _
            LastCell.Row, _
            start_column - 1 _
        ) _
    )
Else
    Set m1 = Range( _
        input_range.Worksheet.Cells(1, 1), _
        input_range.Worksheet.Cells( _
            LastCell.Row, _
            start_column - 1 _
        ) _
    )
    
    Set m2 = Range( _
        input_range.Worksheet.Cells(1, end_column + 1), _
        input_range.Worksheet.Cells( _
            LastCell.Row, _
            LastCell.column _
        ) _
    )
    
    Set Mask = Application.Union(m1, m2)
End If

'mask.Select

Set HSliceExtract = Application.Intersect( _
    input_range, _
    Mask _
)

'HSliceExtract.Select

End Function

Function HVector( _
    ByRef input_range As Range, _
    ByVal Row As Integer _
) As Range
    Set HVector = VSlice( _
        input_range, _
        Row, _
        Row _
    )
    'HVector.Select
End Function

Function VVector( _
    ByRef input_range As Range, _
    ByVal column As Integer _
) As Range
    Set VVector = HSlice( _
        input_range, _
        column, _
        column _
    )
    'VVector.Select
End Function

Function VSlice( _
    ByRef input_range As Range, _
    ByVal start_row As Integer, _
    ByVal end_row As Integer _
) As Range

Dim input_worksheet As Worksheet
Dim Mask As Range
Dim LastCell As Range

Set input_worksheet = input_range.Worksheet
Set LastCell = input_worksheet.Cells.SpecialCells(xlLastCell)

Set Mask = input_worksheet.Range( _
    input_worksheet.Cells(start_row, 1), _
    input_worksheet.Cells( _
        end_row, _
        LastCell.column _
    ) _
)

Set VSlice = Application.Intersect( _
    input_range, _
    Mask _
)

End Function

Function HSlice( _
    ByRef input_range As Range, _
    ByVal start_column As Integer, _
    ByVal end_column As Integer _
) As Range
'HSlice: Range -> Int -> Int -> Range

Dim input_worksheet As Worksheet
Dim Mask As Range
Dim LastCell As Range

Set input_worksheet = input_range.Worksheet
Set LastCell = input_worksheet.Cells.SpecialCells(xlLastCell)

Set Mask = Range( _
    input_worksheet.Cells(1, start_column), _
    input_worksheet.Cells( _
        LastCell.Row, _
        end_column _
    ) _
)

Set HSlice = Application.Intersect( _
    input_range, _
    Mask _
)

End Function

Function VFilter( _
    ByVal filter_value As Variant, _
    ByRef lookup_range As Range, _
    ByVal lookup_column As Integer _
) As Range

Dim lookup_vector As Range

Dim found_cell As Range
Dim FirstFoundRow As Integer

Set lookup_vector = VVector(lookup_range, lookup_column)

Set found_cell = lookup_vector.Find( _
    filter_value, _
    LookIn:=xlValues _
)

If found_cell Is Nothing Then
    Set VFilter = Nothing
    Exit Function
End If

FirstFoundRow = found_cell.Row

Set VFilter = HVector( _
    lookup_range, _
    FirstFoundRow _
)

Do
    Set found_cell = lookup_vector.FindNext(found_cell)
    Set VFilter = Application.Union( _
        VFilter, _
        HVector( _
            lookup_range, _
            found_cell.Row _
        ) _
    )
Loop While found_cell.Row <> FirstFoundRow

'VFilter.Select

End Function

Function HFilter( _
    ByVal lookup_value As Variant, _
    ByRef lookup_range As Range, _
    ByVal lookup_row As Integer _
) As Range

Dim lookup_vector As Range
Dim resoult_range As Range

Dim found_cell As Range
Dim FirstFoundRow As Integer

Set lookup_vector = HSlice(lookup_range, lookup_column, lookup_column)

Set found_cell = lookup_vector.Find( _
    lookup_value, _
    LookIn:=xlValues _
)

If found_cell Is Nothing Then
    Set VFilter = Nothing
    Exit Function
End If

FirstFoundRow = found_cell.Row

Set VFilter = VSlice( _
    lookup_value, _
    FirstFoundRow, _
    FirstFoundRow _
)

Do
    Set found_cell = lookup_vector.FindNext(found_cell)
    Set VFilter = Application.Union( _
        VFilter, _
        VSlice( _
            lookup_value, _
            found_cell.Row, _
            found_cell.Row _
        ) _
    )
Loop While found_cell.Row <> FirstFoundRow

End Function

Function VLookupRange( _
    ByVal lookup_value As Variant, _
    ByRef lookup_range As Range, _
    ByVal lookup_row As Integer _
) As Range
'Works like VLookup, but returns Range all matching resoults

Set VLookupRange = HVectorExtract( _
    VFilter( _
        lookup_value, _
        lookup_range, _
        lookup_row _
    ), _
    lookup_row _
)

End Function

Function HLookupRange( _
    ByVal lookup_value As Variant, _
    ByRef lookup_range As Range, _
    ByVal lookup_column As Integer _
) As Range
'Works like HLookup, but returns Range all matching resoults

Set HLookupRange = VVectorExtract( _
    HFilter( _
        lookup_value, _
        lookup_range, _
        lookup_column _
    ), _
    lookup_column _
)

End Function

Function VLookupChain( _
    input_range As Range, _
    ParamArray filter_values() As Variant _
) As Range

Dim filter_item As Variant

Set VLookupChain = input_range

For Each filter_item In filter_values
    Set VLookupChain = VLookupRange( _
        filter_item, _
        VLookupChain, _
        VLookupChain.Cells(1, 1).column _
    )
Next filter_item

End Function

Function CopyRange(input_range As Range) As Range

Set CopyRange = Application.Intersect( _
    input_range, _
    input_range _
)

End Function
