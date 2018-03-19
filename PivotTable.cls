VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PivotTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'require RangeTools
'require Tools->References->"Microsoft Scripting Runtime"
Option Explicit

Private Dimensions As Scripting.Dictionary

Private Aggregate As Boolean

'==================== PUBLIC METHODS ====================

Public Property Get DimensionArray() As Variant
DimensionArray = Dimensions.Keys
End Property

Public Property Get DimensionCount() As Integer
DimensionCount = Dimensions.Count
End Property

Public Property Get Intersection(ParamArray Params() As Variant) As PivotTable
Dim Pairs() As Variant
Pairs = Params      ' I don't know why I can't pass Params, but
                    ' passing Pairs (the same reference) works ok
Set Intersection = RealIntersection(Pairs)
End Property

Public Function hasDimension(DimensionName) As Boolean
hasDimension = Dimensions.Exists(DimensionName)
End Function

Public Function hasMember(DimensionName, MemberId) As Boolean

If Not Dimensions.Exists(DimensionName) Then
    Err.Raise _
        Number:=12345, _
        Description:="Dimension """ & DimensionName & """ not found in this PivotTable." & _
            Chr(13) & Chr(10) & _
            "Maybe this Intersection already done."
End If

hasMember = Dimensions.Item(DimensionName).Exists(MemberId)
End Function

Public Sub MakeAggregate()
'There is NO path back
Aggregate = True
End Sub

Public Function Union(Other As PivotTable) As PivotTable
'FIXME: STUB
End Function

Public Sub Add( _
    DimensionName As Variant, _
    MemberId As Variant, _
    ItemRange As Range _
)
Dim ThisDimension As Object
Dim ThisRange As Range

If ItemRange Is Nothing Then
    Exit Sub
End If

If Dimensions.Exists(DimensionName) Then
    Set ThisDimension = Dimensions.Item(DimensionName)
Else
    Set ThisDimension = New Scripting.Dictionary
    Dimensions.Add _
        key:=DimensionName, _
        Item:=ThisDimension
End If

If ThisDimension.Exists(MemberId) Then
    Set ThisRange = ThisDimension.Item(MemberId)
    ThisDimension.Remove (MemberId)
    Set ThisRange = Application.Union( _
        ThisRange, _
        ItemRange _
    )
Else
    Set ThisRange = ItemRange
End If

ThisDimension.Add _
    key:=MemberId, _
    Item:=ThisRange

End Sub

Public Property Get MemberArray(DimensionName As Variant) As Variant

If Not Dimensions.Exists(DimensionName) Then
    Err.Raise _
        Number:=12345, _
        Description:="Dimension """ & DimensionName & """ not found in this PivotTable." & _
            Chr(13) & Chr(10) & _
            "Maybe this Intersection already done."
End If

MemberArray = Dimensions.Item(DimensionName).Keys
End Property

Public Property Get Value( _
    DimensionName As Variant, _
    MemberId As Variant _
) As Variant

Dim ThisDimension As Object
Dim MemberCell As Range

If Not Dimensions.Exists(DimensionName) Then
    Err.Raise _
        Number:=12345, _
        Description:="Dimension """ & DimensionName & """ not found in this PivotTable."
End If

Set ThisDimension = Dimensions.Item(DimensionName)

If Not ThisDimension.Exists(MemberId) Then
    Err.Raise _
        Number:=12345, _
        Description:="Dimension """ & DimensionName & """ does not contains" & _
            Chr(13) & Chr(10) & _
            "Member ID = """ & MemberId & """ in this PivotTable."

End If

Set MemberCell = ThisDimension.Item(MemberId)

'MemberCell.Select

If MemberCell Is Nothing Then
    Value = Null
ElseIf Aggregate Then
    Value = WorksheetFunction.Sum(MemberCell)
Else
    Value = MemberCell.Value
End If

End Property

Public Property Let Value( _
    DimensionName As Variant, _
    MemberId As Variant, _
    NewVal As Variant _
)

Dim ThisDimension As Object
Dim MemberCell As Range

If Aggregate Then
    Err.Raise _
        Number:=12345, _
        Description:="This table created as aggregation. Can not" & _
            Chr(13) & Chr(10) & _
            "assign values to aggregation tables because" & _
            Chr(13) & Chr(10) & _
            "target cell can not be determined unambiguously."
End If

If Not Dimensions.Exists(DimensionName) Then
    Err.Raise _
        Number:=12345, _
        Description:="Dimension """ & DimensionName & """ not found in this PivotTable."
End If

Set ThisDimension = Dimensions.Item(DimensionName)

If Not ThisDimension.Exists(MemberId) Then
    Err.Raise _
        Number:=12345, _
        Description:="Dimension """ & DimensionName & """ does not contains" & _
            Chr(13) & Chr(10) & _
            "Member ID = """ & MemberId & """ in this PivotTable."

End If

Set MemberCell = ThisDimension.Item(MemberId)

MemberCell.Value = NewVal

End Property

Public Function RequireIntersections(Target As PivotTable) As Variant()
Dim DimensionName As Variant
Dim GreatestIndex As Integer
Dim Out() As Variant
GreatestIndex = 0
For Each DimensionName In Target.DimensionArray
    If Dimensions.Exists(DimensionName) Then
        GoTo NextDimensionName
    End If
    ReDim Preserve Out(GreatestIndex) As Variant
    Out(GreatestIndex) = DimensionName
    GreatestIndex = GreatestIndex + 1
NextDimensionName:
Next
RequireIntersections = Out
End Function

Public Property Get MinimalDimension() As Variant
Dim DimensionCount As Integer
Dim DimensionArray() As Variant
Dim ThisDimension As Variant
Dim ThisDimensionCount As Integer
Dim i As Integer

DimensionArray = Dimensions.Keys

MinimalDimension = DimensionArray(0)
DimensionCount = Dimensions.Item(MinimalDimension).Count()

For i = 1 To UBound(DimensionArray)
    ThisDimension = DimensionArray(i)
    ThisDimensionCount = Dimensions.Item(ThisDimension).Count()
    If ThisDimensionCount < DimensionCount Then
        MinimalDimension = ThisDimension
        DimensionCount = ThisDimensionCount
    End If
Next i

End Property

Public Function IntersectOne( _
    DimensionName As Variant, _
    MemberId As Variant _
) As PivotTable

Dim ThisDimension As Object
Dim ThisRange As Range
Dim IntersectionRange As Range
Dim NewRange As Range

Dim ThisDimensionName As Variant
Dim ThisMemberId As Variant


If Not Dimensions.Exists(DimensionName) Then
    Err.Raise _
        Number:=12345, _
        Description:="Dimension """ & DimensionName & """ not found in this PivotTable." & _
            Chr(13) & Chr(10) & _
            "Maybe this Intersection already done."
End If

Set ThisDimension = Dimensions.Item(DimensionName)

If Not ThisDimension.Exists(MemberId) Then
    'Maybe we shall return empty PivotTable if member not found
    'Anyway, it's not regular case: something gone wrong
    Err.Raise _
        Number:=12345, _
        Description:="Dimension """ & DimensionName & """ does not contains" & _
            Chr(13) & Chr(10) & _
            "Member ID = """ & MemberId & """ in this PivotTable."

End If

Set IntersectionRange = ThisDimension.Item(MemberId)

Set IntersectOne = New PivotTable

For Each ThisDimensionName In Dimensions.Keys
    If ThisDimensionName = DimensionName Then
        'Continue
        GoTo NextDimension
    End If
    
    Set ThisDimension = Dimensions.Item(ThisDimensionName)
    
    For Each ThisMemberId In ThisDimension.Keys
        
        Set ThisRange = ThisDimension.Item(ThisMemberId)
        Set NewRange = Application.Intersect( _
            ThisRange, _
            IntersectionRange _
        )
        
        If Not NewRange Is Nothing Then
            IntersectOne.Add _
                DimensionName:=ThisDimensionName, _
                MemberId:=ThisMemberId, _
                ItemRange:=NewRange
        End If
    Next

NextDimension:
Next

If Aggregate Then
    IntersectOne.MakeAggregate
End If

End Function

Public Function UnionRange() As Range
Dim DimensionName As Variant
Dim MemberId As Variant
Dim ThisMember As Range
Dim ThisDimension As Object

For Each DimensionName In Dimensions.Keys

    Set ThisDimension = Dimensions.Item(DimensionName)
    For Each MemberId In ThisDimension.Keys

        Set ThisMember = ThisDimension.Item(MemberId)
        If UnionRange Is Nothing Then
            Set UnionRange = ThisMember
        Else
            Set UnionRange = Application.Union( _
                UnionRange, _
                ThisMember _
            )
        End If
    Next
Next
End Function

Public Function AggregateDimensions( _
    ParamArray AggregateDimensionNames() As Variant _
) As PivotTable

Dim MyDimensionName As Variant
Dim AggregateDimensionName As Variant
Dim MemberId As Variant

Dim MyDimension As Object
Dim MemberRange As Range

Set AggregateDimensions = New PivotTable

For Each MyDimensionName In Dimensions.Keys
    For Each AggregateDimensionName In AggregateDimensionNames
        If MyDimensionName = AggregateDimensionName Then
            GoTo NextMyDimensionName
        End If
    Next

    Set MyDimension = Dimensions.Item(MyDimensionName)

    For Each MemberId In MyDimension.Keys
        Set MemberRange = MyDimension.Item(MemberId)

        AggregateDimensions.Add _
            DimensionName:=MyDimensionName, _
            MemberId:=MemberId, _
            ItemRange:=MemberRange
    Next
NextMyDimensionName:
Next

If Aggregate Or UBound(AggregateDimensionNames) > -1 Then
    AggregateDimensions.MakeAggregate
End If

End Function

Public Property Let Intersection( _
    ParamArray Params() As Variant, _
    NewVal As PivotTable _
) ' В этой функции происходит какая-то дичь. Переосмыслить рекурсию.

Dim IntersectDimensionName As Variant
Dim IntersectMemberId As Variant
Dim Pairs() As Variant

Dim MeIntersection As PivotTable

Dim NewCellValue As Variant

Pairs = Params      ' I don't know why I can't pass Params, but
                    ' passing Pairs (the same reference) works ok

Set MeIntersection = RealIntersection(Pairs)

If MeIntersection.DimensionCount > 1 Then
    IntersectDimensionName = MeIntersection.MinimalDimension
    For Each IntersectMemberId In MeIntersection.MemberArray(IntersectDimensionName)
        If Not NewVal.hasMember(IntersectDimensionName, IntersectMemberId) Then
            GoTo NextIntersectMemberIdRecursion
        End If

        ' A(i, j) = B(i, j)
        MeIntersection.Intersection( _
            IntersectDimensionName, _
            IntersectMemberId _
        ) = NewVal.Intersection( _
            IntersectDimensionName, _
            IntersectMemberId _
        )
NextIntersectMemberIdRecursion:
    Next
    Exit Property
End If

IntersectDimensionName = MeIntersection.DimensionArray(0)

For Each IntersectMemberId In MeIntersection.MemberArray(IntersectDimensionName)
    If Not NewVal.hasMember(IntersectDimensionName, IntersectMemberId) Then
        GoTo NextIntersectMemberId
    End If
    
    ' A(i, j) = B(i, j)
    MeIntersection.Value( _
        IntersectDimensionName, _
        IntersectMemberId _
    ) = NewVal.Value( _
        IntersectDimensionName, _
        IntersectMemberId _
    )
NextIntersectMemberId:
Next

End Property

Public Function FilterEquals( _
    FilterDimensionName As Variant, _
    ParamArray FilterMemberIds() As Variant _
) As PivotTable      'FIXME: Not Tested

Dim MyDimensionName As Variant
Dim FilterMemberId As Variant
Dim MyMemberId As Variant

Dim MyDimension As Object
Dim MemberRange As Range
Dim Mask As Range

Set FilterEquals = New PivotTable

Set MyDimension = Dimensions.Item(FilterDimensionName)

For Each FilterMemberId In FilterMemberIds
    Set MemberRange = MyDimension.Item(FilterMemberId)

    If Mask Is Nothing Then
        Set Mask = MemberRange
    Else
        Set Mask = Application.Union( _
            Mask, _
            MemberRange _
        )
    End If

    FilterEquals.Add _
        DimensionName:=MyDimensionName, _
        MemberId:=FilterMemberId, _
        ItemRange:=MemberRange
Next

For Each MyDimensionName In Dimensions.Keys
    If MyDimensionName = FilterDimensionName Then
        GoTo NextMyDimensionName
    End If

    Set MyDimension = Dimensions.Item(MyDimensionName)

    For Each MyMemberId In MyDimension.Keys
        Set MemberRange = MyDimension.Item(MyMemberId)

        Set MemberRange = Application.Intersect( _
            MemberRange, _
            Mask _
        )

        FilterEquals.Add _
            DimensionName:=MyDimensionName, _
            MemberId:=MyMemberId, _
            ItemRange:=MemberRange
    Next
NextMyDimensionName:
Next

End Function

Public Function FilterNotEquals( _
    FilterDimensionName As Variant, _
    ParamArray FilterMemberIds() As Variant _
) As PivotTable      'FIXME: Not Tested

Dim MyDimensionName As Variant
Dim FilterMemberId As Variant
Dim MyMemberId As Variant

Dim MyDimension As Object
Dim MemberRange As Range

Set FilterEquals = New PivotTable

For Each MyDimensionName In Dimensions.Keys
    Set MyDimension = Dimensions.Item(MyDimensionName)
    
    If MyDimensionName = FilterDimensionName Then
        For Each MyMemberId In MyDimension.Keys
            For Each FilterMemberId In FilterMemberIds
                If MyMemberId = FilterMemberId Then
                    GoTo NextMemberId
                End If
            Next
            
            Set MemberRange = MyDimension.Item(MyMemberId)
    
            FilterEquals.Add _
                DimensionName:=MyDimensionName, _
                MemberId:=MyMemberId, _
                ItemRange:=MemberRange
NextMemberId:
        Next
    Else
        For Each MyMemberId In MyDimension.Keys
            Set MemberRange = MyDimension.Item(MyMemberId)
    
            FilterEquals.Add _
                DimensionName:=MyDimensionName, _
                MemberId:=MyMemberId, _
                ItemRange:=MemberRange
        Next
    End If
NextMyDimensionName:
Next

End Function

'==================== PRIVATE METHODS ===================

Private Sub Class_Initialize()
    Set Dimensions = New Scripting.Dictionary
    Aggregate = False
End Sub

Private Sub Class_Terminate()
    Set Dimensions = Nothing
End Sub

Private Function RealIntersection(Pairs() As Variant) As PivotTable
Dim i As Integer
Dim Dimension As Variant
Dim MemberId As Variant

Set RealIntersection = Me

If Not UBound(Pairs) And 1 Then
    ' Lengtn(Pairs) must be even
    ' Проверка длины Pairs на четность
    Err.Raise _
        Number:=12345, _
        Description:="Intersrction creates by pairs Dimension:MemberId." & _
        Chr(13) & Chr(10) & _
        "That's why count of args must be even."
End If

For i = LBound(Pairs) To UBound(Pairs) Step 2
    Dimension = Pairs(i)
    MemberId = Pairs(i + 1)
    Set RealIntersection = RealIntersection.IntersectOne(Dimension, MemberId)
Next i

End Function
