VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TypesFactory"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'METHODS

Public Function NewArray(ByRef Value As Variant) As ArrayEx
    
    Set NewArray = New ArrayEx
    
    NewArray.Parse Value

End Function

Public Function NewBoolean(ByRef Value As Variant) As BooleanEx
    
    Set NewBoolean = New BooleanEx
    
    NewBoolean.Parse Value

End Function

Public Function NewByte(ByRef Value As Variant) As ByteEx
    
    Set NewByte = New ByteEx
    
    NewByte.Parse Value

End Function

Public Function NewCurrency(ByRef Value As Variant) As CurrencyEx
    
    Set NewCurrency = New CurrencyEx
    
    NewCurrency.Parse Value

End Function

Public Function NewDateTime(ByRef Value As Variant) As DateTimeEx
    
    Set NewDateTime = New DateTimeEx
    
    NewDateTime.Parse Value

End Function

Public Function NewDecimal(ByRef Value As Variant) As DecimalEx
    
    Set NewDecimal = New DecimalEx
    
    NewDecimal.Parse Value

End Function

Public Function NewDouble(ByRef Value As Variant) As DoubleEx
    
    Set NewDouble = New DoubleEx
    
    NewDouble.Parse Value

End Function

Public Function NewInteger(ByRef Value As Variant) As IntegerEx
    
    Set NewInteger = New IntegerEx
    
    NewInteger.Parse Value

End Function

Public Function NewList(ByRef Value As Variant) As ListEx
    
    Set NewList = New ListEx
    
    NewList.Parse Value

End Function

Public Function NewLong(ByRef Value As Variant) As LongEx
    
    Set NewLong = New LongEx
    
    NewLong.Parse Value

End Function

Public Function NewSingle(ByRef Value As Variant) As SingleEx
    
    Set NewSingle = New SingleEx
    
    NewSingle.Parse Value

End Function

Public Function NewString(ByRef Value As Variant) As StringEx
    
    Set NewString = New StringEx
    
    NewString.Parse Value

End Function
