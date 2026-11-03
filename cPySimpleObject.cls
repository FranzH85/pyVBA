VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPySimpleObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' ==============================================================================
' CLASS: cPySimpleObject
' DEVELOPER: Franz Hardinger
' HOMEPAGE / YOUTUBE: "Python dic list stile in vba"
' GITHUB: https://github.com
'
' REQUIREMENTS (WINDOWS ONLY):
' Core foundation of the pyVBA framework. Actively used by 'cPyDicList' and
' 'cPyXmltodict' to handle dynamic node mutations seamlessly.
' ==============================================================================
' LICENSE: MIT License
'
' Copyright (c) 2026 Franz Hardinger
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
' ==============================================================================
'
' DESCRIPTION:
' A universal data vessel implementing a genuine dynamic variant type in VBA.
' It seamlessly encapsulates both primitive data and complex object references
' via unified Let/Set property interfaces. Features a transparent default member
' ID attribute alongside a dedicated metadata slot (.MyTyp).
' ==============================================================================

Option Explicit

' ==========================
' Basis-Arrays & Einstellungen
' ==========================
Dim localValue As Variant
Dim localTyp As Variant
 

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

' ==========================
' Properties (Eigenschaften)
' ==========================
Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
    If IsObject(localValue) Then
        Set Value = localValue
    Else
        Value = localValue
    End If
End Property
 
Public Property Let Value(Value_ As Variant)
Attribute Value.VB_UserMemId = 0
    If IsObject(Value_) Then
        Set localValue = Value_
    Else
        localValue = Value_
    End If
End Property

Public Property Set Value(Value_ As Variant)
Attribute Value.VB_UserMemId = 0
    If IsObject(Value_) Then
        Set localValue = Value_
    Else
        localValue = Value_
    End If
End Property
