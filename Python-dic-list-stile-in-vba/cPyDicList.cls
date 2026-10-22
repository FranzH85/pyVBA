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
' KLASSE: cPyDicList
' ENTWICKLER: Franz Hardinger
' LIZENZ: Frei verwendbar (MIT License)
'
' BESCHREIBUNG:
' Ein extrem schneller und flexibler Daten-Container fuer VBA, der die elegante
' Syntax von Python-Dictionaries direkt nach Excel bringt. Funktioniert komplett
' ohne Installation auf Windows und Mac.
'
' WICHTIGER HINWEIS FUER DIE PYTHON-SYNTAX:
' Damit die Klammer-Kettung ohne ".Value" funktioniert, muss die Eigenschaft
' "Property Get Value" das versteckte Attribut 'VB_UserMemId = 0' besitzen!
' (Entweder die .cls-Datei importieren oder per Texteditor im Code hinterlegen)
'
' YOUTUBE VIDEO: "Python dic list stile in vba"
' GITHUB: [Hier kommt dein Link hin]
'
' REALE ANWENDUNG (Kompatibel von Excel 2007 bis 2026+):
'   Dim PyDicList As New cPyDicList
'   PyDicList.Append "Schiff", New cPyDicList
'   PyDicList.LastValue.Append "Name", "Enterprise"
'   Debug.Print PyDicList("Schiff")("Name") ' Ausgabe: Enterprise
' ==============================================================================


' ==========================
' Basis-Arrays & Einstellungen
' ==========================
Dim KeyArray() As String
Dim ValueArray() As Variant
Dim localImax As Long
Public CaseSensitive As Boolean
Public MyTyp As Variant
 
' ==========================
' Initialisierung & Speicher
' ==========================
Private Sub Class_Initialize()
    localImax = -1
    ReDim KeyArray(0)
    ReDim ValueArray(0)
    CaseSensitive = True
End Sub
 
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
    Dim i As Long
    Dim CompareMethod As Long
    
    If CaseSensitive Then
        CompareMethod = vbBinaryCompare
    Else
        CompareMethod = vbTextCompare
    End If
    
    For i = 0 To localImax
        If StrComp(KeyArray(i), Key_, CompareMethod) = 0 Then
            GetKeyIndex = i
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
    Dim i As Long
    i = GetKeyIndex(Key_)
    
    If i = -1 Then
        ReDimValue
        i = localImax
    End If
 
    KeyArray(i) = Key_
    If IsObject(Value_) Then
        Set ValueArray(i) = Value_
    Else
        ValueArray(i) = Value_
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
    Dim i As Long
    i = GetKeyIndex(Key_)
    If i > -1 Then
        If IsObject(ValueArray(i)) Then
            Set GetDefault = ValueArray(i)
        Else
            GetDefault = ValueArray(i)
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
    Dim i As Long
    i = GetKeyIndex(Key_)
    If i > -1 Then
        If IsObject(ValueArray(i)) Then
            Set Pop = ValueArray(i)
        Else
            Pop = ValueArray(i)
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
    Dim i As Long
    i = GetKeyIndex(Key_)
    If i > -1 Then
        If IsObject(ValueArray(i)) Then
            Set Value = ValueArray(i)
        Else
            Value = ValueArray(i)
        End If
        Exit Property
    End If
    Value = Empty
End Property
 
Public Property Let Value(Key_ As String, Value_ As Variant)
Attribute Value.VB_UserMemId = 0
    Dim i As Long
    i = GetKeyIndex(Key_)
    If i > -1 Then
        ValueArray(i) = Value_
    End If
End Property

Public Property Set Value(Key_ As String, Value_ As Variant)
Attribute Value.VB_UserMemId = 0
    Dim i As Long
    i = GetKeyIndex(Key_)
    If i > -1 Then
        Set ValueArray(i) = Value_
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
    ValueArray(localImax) = Value_
End Property

Public Property Set LastValue(Value_ As Variant)
    Set ValueArray(localImax) = Value_
End Property
 
Public Property Get ValueIndex(I_ As Long) As Variant
    If IsObject(ValueArray(I_)) Then
        Set ValueIndex = ValueArray(I_)
    Else
        ValueIndex = ValueArray(I_)
    End If
End Property
 
Public Property Let ValueIndex(I_ As Long, Value_ As Variant)
    ValueArray(I_) = Value_
End Property

Public Property Set ValueIndex(I_ As Long, Value_ As Variant)
    Set ValueArray(I_) = Value_
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
    Dim i As Long
    Dim IsObjectValue As Boolean
    
    IsObjectValue = IsObject(Value_)
    
    For i = 0 To localImax
        If IsObjectValue Then
            If IsObject(ValueArray(i)) Then
                If ValueArray(i) Is Value_ Then
                    ValueExists = True
                    Exit Function
                End If
            End If
        Else
            If Not IsObject(ValueArray(i)) Then
                If ValueArray(i) = Value_ Then
                    ValueExists = True
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Function FindIndicesByKey(ByVal Key_ As String) As cPyDicList
    Dim i As Long
    Dim CompareMethod As Long
    Dim TempList As New cPyDicList
    
    If CaseSensitive Then
        CompareMethod = vbBinaryCompare
    Else
        CompareMethod = vbTextCompare
    End If
    
    For i = 0 To localImax
        If StrComp(KeyArray(i), Key_, CompareMethod) = 0 Then
            TempList.Append Key_, i
        End If
    Next i
    
    Set FindIndicesByKey = TempList
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
    Dim i As Long, j As Long
    i = GetKeyIndex(Key_)
    If i = -1 Then Exit Sub
    
    For j = i To localImax - 1
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
    Dim i As Long
    For i = 0 To localImax
        Debug.Print KeyArray(i)
    Next
End Sub
