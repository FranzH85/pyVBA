VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPyDicList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==============================================================================
' CLASS: cPyDicList
' DEVELOPER: Franz Hardinger
' LICENSE: MIT License (Free to use)
'
' DESCRIPTION:
' An extremely fast and flexible data container for VBA that brings the elegant
' syntax of Python dictionaries directly into Excel. Works completely standalone
' without requiring any installation on Windows and Mac.
'
' IMPORTANT NOTE FOR PYTHON SYNTAX:
' For bracket-chaining to work seamlessly without appending ".Value", the
' "Property Get Value" property must have the hidden attribute 'VB_UserMemId = 0'!
' (Either import the native .cls file or inject it via a text editor in the source code)
'
' YOUTUBE VIDEO: "Python dic list stile in vba"
' GITHUB: [Insert your link here]
'
' REAL-WORLD USAGE (Compatible from Excel 2007 up to 2026+):
'   Dim PyDicList As New cPyDicList
'   PyDicList.Append "Ship", New cPyDicList
'   PyDicList.LastValue.Append "Name", "Enterprise"
'   Debug.Print PyDicList("Ship")("Name") ' Output: Enterprise
' ==============================================================================


' ==========================
' Basis-Arrays & Einstellungen
' ==========================
Dim KeyArray() As String
Dim ValueArray() As Variant
Dim localImax As Long
Public CaseSensitive As Boolean
Dim localTyp As Variant
Dim localNewPyXMLSyndax As Boolean

 
' ==========================
' Initialisierung & Speicher
' ==========================
Private Sub Class_Initialize()
    localImax = -1
    ReDim KeyArray(0)
    ReDim ValueArray(0)
    CaseSensitive = True
End Sub

Public Property Get MyTyp() As Variant
    If IsObject(localTyp) Then
        Set MyTyp = localTyp
    Else
        MyTyp = localTyp
    End If
End Property
 
Public Property Let MyTyp(Typ_ As Variant)
    If IsObject(Typ_) Then
        Set localTyp = Typ_
    Else
        localTyp = Typ_
    End If
End Property

Public Property Set MyTyp(Typ_ As Variant)
    If IsObject(Typ_) Then
        Set localTyp = Typ_
    Else
        localTyp = Typ_
    End If
End Property

Public Property Get NewPyXMLSyndax() As Boolean
    NewPyXMLSyndax = localNewPyXMLSyndax
End Property
 
Public Property Let NewPyXMLSyndax(NewPyXMLSyndax_ As Boolean)
    localNewPyXMLSyndax = NewPyXMLSyndax_
End Property
 
Sub ReDimValue()
    localImax = localImax + 1
    ReDim Preserve KeyArray(localImax)
    ReDim Preserve ValueArray(localImax)
End Sub
 
Sub SetNewMaxIndex(NewMaxIndex_ As Long)
    localImax = NewMaxIndex_
    ReDim Preserve KeyArray(localImax)
    ReDim Preserve ValueArray(localImax)
End Sub
 
Sub DeleteAll()
    localImax = -1
    ReDim KeyArray(0)
    ReDim ValueArray(0)
End Sub

' ==========================
' Python-Style Key-Wert-Funktionen
' ==========================
Function GetKeyIndex(Key_ As String) As Long
    Dim I As Long
    Dim CompareMethod As Long
    
    If CaseSensitive Then
        CompareMethod = vbBinaryCompare
    Else
        CompareMethod = vbTextCompare
    End If
    
    For I = 0 To localImax
        If StrComp(KeyArray(I), Key_, CompareMethod) = 0 Then
            GetKeyIndex = I
            Exit Function
        End If
    Next
    GetKeyIndex = -1
End Function

Sub UpdateWithIndex(I_ As Long, NewKey_ As String, Value_ As Variant)
    KeyArray(I_) = NewKey_
    If IsObject(Value_) Then
        Set ValueArray(I_) = Value_
    Else
        ValueArray(I_) = Value_
    End If
End Sub
 
Sub Update(Key_ As String, Value_ As Variant)
    Dim I As Long
    I = GetKeyIndex(Key_)
    
    If I = -1 Then
        ReDimValue
        I = localImax
    End If
 
    KeyArray(I) = Key_
    If IsObject(Value_) Then
        Set ValueArray(I) = Value_
    Else
        If localNewPyXMLSyndax Then
            ValueArray(I).Value = Value_
        Else
            ValueArray(I) = Value_
        End If
    End If
End Sub
 
Sub Append(Key_ As String, Value_ As Variant)
    ReDimValue
    KeyArray(localImax) = Key_
    If IsObject(Value_) Then
        Set ValueArray(localImax) = Value_
    Else
        ValueArray(localImax) = Value_
    End If
End Sub

Function GetDefault(Key_ As String, Default_ As Variant) As Variant
    Dim I As Long
    I = GetKeyIndex(Key_)
    If I > -1 Then
        If IsObject(ValueArray(I)) Then
            Set GetDefault = ValueArray(I)
        Else
            GetDefault = ValueArray(I)
        End If
    Else
        If IsObject(Default_) Then
            Set GetDefault = Default_
        Else
            GetDefault = Default_
        End If
    End If
End Function

Function Pop(Key_ As String) As Variant
    Dim I As Long
    I = GetKeyIndex(Key_)
    If I > -1 Then
        If IsObject(ValueArray(I)) Then
            Set Pop = ValueArray(I)
        Else
            Pop = ValueArray(I)
        End If
        KeyDelete Key_
    Else
        Pop = Empty
    End If
End Function

' ==========================
' Properties (Eigenschaften)
' ==========================
Public Property Get Value(Key_ As String) As Variant
Attribute Value.VB_UserMemId = 0
    Dim I As Long
    I = GetKeyIndex(Key_)
    If I > -1 Then
        If IsObject(ValueArray(I)) Then
            Set Value = ValueArray(I)
        Else
            Value = ValueArray(I)
        End If
        Exit Property
    End If
    Value = Empty
End Property
 
Public Property Let Value(Key_ As String, Value_ As Variant)
Attribute Value.VB_UserMemId = 0
    Dim I As Long
    I = GetKeyIndex(Key_)
    If I > -1 Then
        If IsObject(Value_) Then
            Set ValueArray(I) = Value_
        Else
            If localNewPyXMLSyndax Then
                ValueArray(I).Value = Value_
            Else
                ValueArray(I) = Value_
            End If
        End If
    End If
End Property

Public Property Set Value(Key_ As String, Value_ As Variant)
Attribute Value.VB_UserMemId = 0
    Dim I As Long
    I = GetKeyIndex(Key_)
    If I > -1 Then
        If IsObject(Value_) Then
            Set ValueArray(I) = Value_
        Else
            ValueArray(I) = Value_
        End If
    End If
End Property

Public Property Get LastValue() As Variant
    If IsObject(ValueArray(localImax)) Then
        Set LastValue = ValueArray(localImax)
    Else
        LastValue = ValueArray(localImax)
    End If
End Property
 
Public Property Let LastValue(Value_ As Variant)
    If IsObject(Value_) Then
        Set ValueArray(localImax) = Value_
    Else
        If localNewPyXMLSyndax Then
            ValueArray(localImax).Value = Value_
        Else
            ValueArray(localImax) = Value_
        End If
    End If
End Property

Public Property Set LastValue(Value_ As Variant)
    If IsObject(Value_) Then
        Set ValueArray(localImax) = Value_
    Else
        ValueArray(localImax) = Value_
    End If
End Property
 
Public Property Get ValueIndex(I_ As Long) As Variant
    If IsObject(ValueArray(I_)) Then
        Set ValueIndex = ValueArray(I_)
    Else
        ValueIndex = ValueArray(I_)
    End If
End Property
 
Public Property Let ValueIndex(I_ As Long, Value_ As Variant)
    If IsObject(Value_) Then
        Set ValueArray(I_) = Value_
    Else
        If localNewPyXMLSyndax Then
            ValueArray(I_).Value = Value_
        Else
            ValueArray(I_) = Value_
        End If
    End If
End Property

Public Property Set ValueIndex(I_ As Long, Value_ As Variant)
    If IsObject(Value_) Then
        Set ValueArray(I_) = Value_
    Else
        ValueArray(I_) = Value_
    End If
End Property
 
Function Exists(Key_ As String) As Boolean
    Dim Index As Long
    Index = GetKeyIndex(Key_)
    
    If Index > -1 Then
        Exists = True
    Else
        Exists = False
    End If
End Function

Function ValueExists(Value_ As Variant) As Boolean
    Dim I As Long
    Dim IsObjectValue As Boolean
    
    IsObjectValue = IsObject(Value_)
    
    For I = 0 To localImax
        If IsObjectValue Then
            If IsObject(ValueArray(I)) Then
                If ValueArray(I) Is Value_ Then
                    ValueExists = True
                    Exit Function
                End If
            End If
        Else
            If Not IsObject(ValueArray(I)) Then
                If ValueArray(I) = Value_ Then
                    ValueExists = True
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Function FindIndicesByKey(ByVal Key_ As String) As cPyDicList
    Dim I As Long
    Dim CompareMethod As Long
    Dim TempList As New cPyDicList
    
    If CaseSensitive Then
        CompareMethod = vbBinaryCompare
    Else
        CompareMethod = vbTextCompare
    End If
    
    For I = 0 To localImax
        If StrComp(KeyArray(I), Key_, CompareMethod) = 0 Then
            TempList.Append Key_, I
        End If
    Next I
    
    Set FindIndicesByKey = TempList
End Function

Function FindIndicesByValue(ByVal Value_ As String) As cPyDicList
    Dim I As Long
    Dim CompareMethod As Long
    Dim TempList As New cPyDicList
    
    For I = 0 To localImax
        If ValueArray(I) = Value_ Then
            TempList.Append Key_, I
        End If
    Next I
    
    Set FindIndicesByKey = TempList
End Function

Function getListByKey(ByVal Key_ As String) As cPyDicList
    Dim I As Long
    Dim CompareMethod As Long
    Dim TempList As New cPyDicList
    Dim Imax As Long
    Dim ITempList As Long
    
    If CaseSensitive Then
        CompareMethod = vbBinaryCompare
    Else
        CompareMethod = vbTextCompare
    End If
    
    Imax = -1
    For I = 0 To localImax
        If StrComp(KeyArray(I), Key_, CompareMethod) = 0 Then
            Imax = Imax + 1
        End If
    Next I
    
    If Imax = -1 Then
        Set getListByKey = TempList
        Exit Function
    End If
    
    TempList.SetNewMaxIndex Imax
    
    For I = 0 To localImax
        If StrComp(KeyArray(I), Key_, CompareMethod) = 0 Then
            TempList.UpdateWithIndex ITempList, KeyArray(I), ValueArray(I)
            ITempList = ITempList + 1
        End If
    Next I
    
    Set getListByKey = TempList
End Function

Public Property Get Key(I_ As Long) As String
    Key = KeyArray(I_)
End Property
 
Public Property Let Key(I_ As Long, NewKey_ As String)
    KeyArray(I_) = NewKey_
End Property

Public Property Get Imax() As Long
    Imax = localImax
End Property

Public Property Get Count() As Long
    Count = localImax + 1
End Property
 
' ==========================
' Datenmanipulation & Debug
' ==========================
Sub KeyDelete(Key_ As String)
    Dim I As Long, j As Long
    I = GetKeyIndex(Key_)
    If I = -1 Then Exit Sub
    
    For j = I To localImax - 1
        KeyArray(j) = KeyArray(j + 1)
        If IsObject(ValueArray(j + 1)) Then
            Set ValueArray(j) = ValueArray(j + 1)
        Else
            ValueArray(j) = ValueArray(j + 1)
        End If
    Next j
    
    localImax = localImax - 1
    If localImax >= 0 Then
        ReDim Preserve KeyArray(localImax)
        ReDim Preserve ValueArray(localImax)
    Else
        ReDim KeyArray(0)
        ReDim ValueArray(0)
    End If
End Sub
 
Sub DebugPrintAllKey()
    Dim I As Long
    For I = 0 To localImax
        Debug.Print KeyArray(I)
    Next
End Sub
