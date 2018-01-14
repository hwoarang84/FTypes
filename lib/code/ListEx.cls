VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'INTERFACES

Implements TypesInterface

'VARIABLES

Private p_Array As TSAFEARRAY

'EVENTS

Private Sub Class_Terminate()
    
    ArrayDestroy p_Array

End Sub

'PROPERTIES

Private Property Get TypesInterface_Pointer() As Long
    
    TypesInterface_Pointer = p_Array.lPointer

End Property

Public Property Get Count() As Long
    
    If p_Array.lPointer Then Count = p_Array.uBounds(0).lElements

End Property

Public Property Get Data() As Long
    
    Data = p_Array.lData

End Property

Public Property Get ElementSize() As Long
    
    ElementSize = p_Array.lElementSize

End Property

Public Property Get IndexOf(ByRef Value As Variant) As Long
    
    Dim i As Long
    Dim v As Variant
    
    IndexOf = L_NG
    
    If p_Array.lPointer Then
        
        VariantSetType v, p_Array.lVarType
        
        For i = 0& To p_Array.uBounds(0).lElements + L_NG
            
            If ArrayElementGet(p_Array, i, v) Then
                
                If v = Value Then
                    
                    IndexOf = i
                    
                    Exit For
                
                End If
            
            End If
        
        Next i
        
        VariantZero v
    
    End If

End Property

Public Property Get Item(ByVal Index As Long) As Variant
    
    If Index >= 0& And Index < Count Then
        
        VariantSetType Item, p_Array.lVarType
        
        ArrayElementGet p_Array, Index, Item
    
    End If

End Property

Public Property Let Item(ByVal Index As Long, ByRef NewValue As Variant)
    
    If Index >= 0& And Index < Count Then
        
        If ArrayElementSet(p_Array, Index, NewValue) Then p_Array.bSorted = False
    
    End If

End Property

Public Property Get Pointer() As Long
    
    Pointer = TypesInterface_Pointer

End Property

Public Property Get Sorted() As Boolean
    
    Sorted = p_Array.bSorted

End Property

Public Property Get Type_() As VbVarType
    
    Type_ = p_Array.lVarType

End Property

Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "200"
    
    If p_Array.lPointer Then ArrayCopy p_Array, Value

End Property

Public Property Let Value(ByRef NewValue As Variant)
    
    Dim a As TSAFEARRAY
    
    ArrayPtr a, VarPtr(NewValue) * (((VarType(NewValue) And vbArray) = vbArray) * L_NG), True
    
    If a.lPointer > 0& And a.iDims = 1 Then
        
        If ArrayCreateVector(p_Array, a.lVarType, a.uBounds(0).lElements) Then ArrayCopy p_Array, a.lPointer
    
    End If

End Property

'METHODS

Private Function TypesInterface_Clone() As TypesInterface
    
    Set TypesInterface_Clone = New ListEx
    
    If p_Array.lPointer Then TypesInterface_Clone.Parse Me.Value

End Function

Private Function TypesInterface_Parse(Value As Variant) As TypesInterface
    
    Set TypesInterface_Parse = Me
    
    If (VarType(Value) And vbArray) = vbArray Then
        
        Me.Value = Value
    
    Else
        
        If ArrayCreateVector(p_Array, VarType(Value), 1&) Then ArrayElementSet p_Array, 0&, Value
    
    End If

End Function

Public Function Add(ByRef Item As Variant) As ListEx
    
    Set Add = Me
    
    If (VarType(Item) And vbArray) <> vbArray Then
        
        If p_Array.lPointer Then ArrayRedim p_Array, p_Array.uBounds(0).lElements + 1& Else ArrayCreateVector p_Array, VarType(Item), 1&
        
        ArrayElementSet p_Array, p_Array.uBounds(0).lElements + L_NG, Item
    
    End If

End Function

Public Function AddRange(ByRef Range As Variant) As ListEx
    
    Dim a As TSAFEARRAY
    Dim i As Long
    Dim x As Long
    
    Set AddRange = Me
    
    If p_Array.lPointer Then
        
        ArrayPtr a, VarPtr(Range) * (((VarType(Range) And vbArray) = vbArray) * L_NG), True
        
        If a.lPointer > 0& And a.iDims = 1& Then
            
            x = p_Array.uBounds(0).lElements
            
            ArrayRedim p_Array, x + a.uBounds(0).lElements
            
            For i = a.uBounds(0).lLowest To a.uBounds(0).lLowest + a.uBounds(0).lElements + L_NG
                
                x = x + 1&
                
                ArrayElementSet p_Array, x + L_NG, Range(i)
            
            Next i
        
        End If
    
    Else
        
        Me.Value = Range
    
    End If

End Function

Public Function Clone() As ListEx
    
    Set Clone = TypesInterface_Clone

End Function

Public Function Create(ByVal ArrayType As VbVarType, ByVal Count As Long) As ListEx
    
    Set Create = Me
    
    ArrayCreateVector p_Array, ArrayType, Count

End Function

Public Function Destroy() As ListEx
    
    Set Destroy = Me
    
    ArrayDestroy p_Array

End Function

Public Function Distinct() As ListEx
    
    Dim i As Long
    Dim j As Long
    Dim v1 As Variant
    Dim v2 As Variant
    Dim u As Long
    Dim x As Long
    
    Set Distinct = Me
    
    If p_Array.lPointer Then
        
        Sort
        
        u = p_Array.uBounds(0).lElements + L_NG
        
        VariantSetType v1, p_Array.lVarType
        VariantSetType v2, p_Array.lVarType
        
        Do
            
            ArrayElementGet p_Array, i, v1
            
            For j = i + 1& To u
                
                i = j
                
                ArrayElementGet p_Array, j, v2
                
                If v1 <> v2 Then
                    
                    x = x + 1&
                    
                    ArrayElementSet p_Array, x, v2
                    
                    Exit For
                
                End If
            
            Next j
        
        Loop While i < u
        
        VariantZero v1
        VariantZero v2
        
        ArrayRedim p_Array, x + 1&, True
    
    End If

End Function

Public Function Filter(ByRef Values As Variant) As ListEx
    
    Set Filter = Me
    
    If (VarType(Values) And vbArray) = vbArray Then FilterInternal Values, False, False Else FilterInternal Array(Values), False, False

End Function

Public Function FilterAt(ByRef Indexes As Variant) As ListEx
    
    Set FilterAt = Me
    
    If (VarType(Indexes) And vbArray) = vbArray Then FilterInternal Indexes, False, True Else FilterInternal Array(Indexes), False, True

End Function

Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    
    Set NewEnum = NewEnumerator(Me, Count + L_NG)

End Function

Public Function Parse(ByRef Value As Variant) As ListEx
    
    Set Parse = TypesInterface_Parse(Value)

End Function

Public Function Remove(ByRef Values As Variant) As ListEx
    
    Set Remove = Me
    
    If (VarType(Values) And vbArray) = vbArray Then FilterInternal Values, True, False Else FilterInternal Array(Values), True, False

End Function

Public Function RemoveAt(ByRef Indexes As Variant) As ListEx
    
    Set RemoveAt = Me
    
    If (VarType(Indexes) And vbArray) = vbArray Then FilterInternal Indexes, True, True Else FilterInternal Array(Indexes), True, True

End Function

Public Function Resize(ByVal NewUpperBound As Long) As ListEx
    
    Set Resize = Me
    
    If p_Array.lPointer > 0& And NewUpperBound >= 0& Then ArrayRedim p_Array, NewUpperBound + 1&

End Function

Public Function Sort(Optional ByVal Descending As Boolean, Optional ByVal StartIndex As Long = L_NG, Optional ByVal EndIndex As Long = L_NG) As ListEx
    
    On Error GoTo ErrHandler
    
    Dim e As Long
    Dim i As Long
    Dim j As Long
    Dim p As Variant
    Dim s As Long
    Dim v1 As Variant
    Dim v2 As Variant
    Dim u As Long
    
    Set Sort = Me
    
    If p_Array.lPointer Then
        
        e = p_Array.uBounds(0).lElements + L_NG
        
        If StartIndex > 0& And StartIndex <= e Then s = StartIndex
        If EndIndex >= s And EndIndex <= e Then e = EndIndex
        
        u = (s + e) \ 2&
        
        VariantSetType p, p_Array.lVarType
        
        If ArrayElementGet(p_Array, u, p) Then
            
            VariantSetType v1, p_Array.lVarType
            VariantSetType v2, p_Array.lVarType
            
            i = s
            j = e
            
            While i <= j
                
                Do
                    
                    u = i
                    
                    ArrayElementGet p_Array, u, v1
                    
                    If ((Descending = False And v1 < p) Or (Descending = True And v1 > p)) And i < e Then i = i + 1& Else Exit Do
                
                Loop
                
                Do
                    
                    u = j
                    
                    ArrayElementGet p_Array, u, v2
                    
                    If ((Descending = False And p < v2) Or (Descending = True And p > v2)) And j > s Then j = j + L_NG Else Exit Do
                
                Loop
                
                If i <= j Then
                    
                    ArrayElementSet p_Array, u, v1
                    u = i
                    ArrayElementSet p_Array, u, v2
                    
                    i = i + 1&
                    j = j + L_NG
                
                End If
            
            Wend
            
            VariantZero v1
            VariantZero v2
            
            If s < j Then
                u = s
                Sort Descending, u, j
            End If
            
            If i < e Then
                u = i
                Sort Descending, u, e
            End If
        
        End If
        
        VariantZero p
        
        p_Array.bSorted = True
    
    End If

ErrHandler:

End Function

Public Function ToString(Optional ByRef Delimeter As String, Optional ByVal StartIndex As Long = L_NG, Optional ByVal EndIndex As Long = L_NG) As StringEx
    
    Dim i As Long
    Dim j As Long
    Dim v As Variant
    
    Set ToString = New StringEx
    
    If p_Array.lPointer Then
        
        j = p_Array.uBounds(0).lElements + L_NG
        
        If StartIndex > 0& And StartIndex <= j Then i = StartIndex
        If EndIndex >= i And EndIndex <= j Then j = EndIndex
        
        VariantSetType v, p_Array.lVarType
        
        For i = i To j
            
            If ArrayElementGet(p_Array, i, v) Then ToString.Concat((v)).Concat Delimeter
        
        Next i
        
        VariantZero v
    
    End If

End Function

Private Sub FilterInternal(ByRef Values As Variant, ByVal Remove As Boolean, ByVal ValuesAreIndexes As Boolean)
    
    Dim f As Boolean
    Dim i As Long
    Dim j As Long
    Dim x As Long
    Dim u As Long
    Dim v As Variant
    
    u = UBound(Values)
    
    If u > L_NG Then
        
        VariantSetType v, p_Array.lVarType
        
        For i = 0& To Count + L_NG
            
            ArrayElementGet p_Array, i, v
            
            f = False
            
            For j = 0& To u
                
                If ValuesAreIndexes Then
                    
                    If i = Values(j) Then
                        
                        f = True
                        
                        Exit For
                    
                    End If
                
                Else
                    
                    If v = Values(j) Then
                        
                        f = True
                        
                        Exit For
                    
                    End If
                
                End If
            
            Next j
            
            If Remove <> f Then
                
                If i <> x Then ArrayElementSet p_Array, x, v
                
                x = x + 1&
            
            End If
        
        Next i
        
        VariantZero v
        
        ArrayRedim p_Array, x, True
    
    End If

End Sub
