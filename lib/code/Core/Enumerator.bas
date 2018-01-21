Attribute VB_Name = "Enumerator"
Option Explicit

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'
' Copyright © 2017 Dexter Freivald. All Rights Reserved. DEXWERX.COM
'
' See original source at http://www.vbforums.com/showthread.php?854963-VB6-IEnumVARIANT-For-Each-support-without-a-typelib
'
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'TYPES

Private Type TENUMERATOR
    lVTablePtr As Long
    lReferences As Long
    uEnumerable As Object
    lIndex As Long
    lUpper As Long
End Type

'CONSTANTS

Private Const E_NOTIMPL As Long = &H80004001

'VARIABLES

Private m_Table(6) As Long

'WINAPI

Private Declare Function CopyBytesZero Lib "msvbvm60" Alias "__vbaCopyBytesZero" (ByVal lLen As Long, ByVal lTarget As Long, uSource As Any) As Long
Private Declare Function CoTaskMemAlloc Lib "ole32" (ByVal lMem As Long) As Long
Private Declare Function GetMem4 Lib "msvbvm60" (lSource As Long, uTarget As Any) As Long
Private Declare Function VariantCopy Lib "oleaut32" (ByVal lTarget As Long, ByRef uSource As Variant) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal lMem As Long)

'METHODS

Private Function IEnumVARIANT_Clone(ByRef This As TENUMERATOR, ByVal lEnum As Long) As Long
    
    IEnumVARIANT_Clone = E_NOTIMPL

End Function

Private Function IEnumVARIANT_Next(ByRef This As TENUMERATOR, ByVal lCelt As Long, ByVal lVar As Long, ByVal lFetched As Long) As Long
    
    With This
        
        If .lIndex > .lUpper Then
            
            IEnumVARIANT_Next = 1&
        
        Else
            
            VariantCopy lVar, .uEnumerable.Item(.lIndex)
            
            .lIndex = .lIndex + 1&
        
        End If
    
    End With

End Function

Private Function IEnumVARIANT_Reset(ByRef This As TENUMERATOR) As Long
    
    IEnumVARIANT_Reset = E_NOTIMPL

End Function

Private Function IEnumVARIANT_Skip(ByRef This As TENUMERATOR, ByVal lCelt As Long) As Long
    
    IEnumVARIANT_Skip = E_NOTIMPL

End Function

Private Function IUnknown_AddRef(ByRef This As TENUMERATOR) As Long
    
    With This
        
        .lReferences = .lReferences + 1&
         
         IUnknown_AddRef = .lReferences
    
    End With

End Function

Private Function IUnknown_QueryInterface(ByRef This As TENUMERATOR, ByVal lRiid As Long, ByVal lObject As Long) As Long
    
    GetMem4 VarPtr(This), ByVal lObject
    
    This.lReferences = This.lReferences + 1&

End Function

Private Function IUnknown_Release(ByRef This As TENUMERATOR) As Long
    
    With This
        
        .lReferences = .lReferences - 1&
         
         IUnknown_Release = .lReferences
        
        If .lReferences = 0& Then
            
            Set .uEnumerable = Nothing
            
            CoTaskMemFree VarPtr(This)
        
        End If
    
    End With

End Function

Public Function NewEnumerator(ByRef uEnumerable As Object, ByVal lUpper As Long) As IEnumVARIANT
    
    Dim e As TENUMERATOR
    Dim p As Long
    
    If m_Table(0) = 0& Then
        
        RtlMoveMemory m_Table(0), AddressOf IUnknown_QueryInterface, 4&
        RtlMoveMemory m_Table(1), AddressOf IUnknown_AddRef, 4&
        RtlMoveMemory m_Table(2), AddressOf IUnknown_Release, 4&
        RtlMoveMemory m_Table(3), AddressOf IEnumVARIANT_Next, 4&
        RtlMoveMemory m_Table(4), AddressOf IEnumVARIANT_Skip, 4&
        RtlMoveMemory m_Table(5), AddressOf IEnumVARIANT_Reset, 4&
        RtlMoveMemory m_Table(6), AddressOf IEnumVARIANT_Clone, 4&
    
    End If
    
    With e
        .lVTablePtr = VarPtr(m_Table(0))
        .lUpper = lUpper
        .lReferences = 1&
         Set .uEnumerable = uEnumerable
    End With
    
    p = CoTaskMemAlloc(LenB(e))
    
    CopyBytesZero LenB(e), p, e
    
    GetMem4 p, NewEnumerator

End Function
