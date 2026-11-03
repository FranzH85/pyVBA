VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPyBookmark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' ==============================================================================
' CLASS: cPyBookmark
' DEVELOPER: Franz Hardinger
' HOMEPAGE / YOUTUBE: "Python dic list stile in vba"
' GITHUB: https://github.com
'
' REQUIREMENTS (WINDOWS ONLY):
' This module wraps Microsoft Word Bookmark manipulation into a clean, key-value
' dictionary syntax. It requires a standard reference to the Microsoft Word Object Library.
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
' An elegant wrapper for Word Bookmarks providing a Python dictionary style
' interface, automated bookmark preservation during text updates, runtime formatting
' via inline markup tags (\bold, \itali, \size:), and a bookmark constant generator.
' ==============================================================================

Option Explicit


' ==========================
' Properties (Eigenschaften)
' ==========================
Public Property Get Value(Key_ As String) As String
Attribute Value.VB_UserMemId = 0
    Value = ThisDocument.Bookmarks.Item(Key_).Range.Text
End Property
 
Public Property Let Value(Key_ As String, Text_ As String)
Attribute Value.VB_UserMemId = 0
    Dim currentRange As Range
    Dim Text As String
    Dim i As Long
    Dim Imax As Long
    Dim posSize As Long
    Dim sizeStr As String
    Dim isBold As Boolean
    Dim isItalic As Boolean
    Dim fontSize As Single
    
    Text = Text_

    Set currentRange = ThisDocument.Bookmarks.Item(Key_).Range
    
    If InStr(Text, "\bold") > 0 Then
        isBold = True
        Text = Replace(Text, "\bold", "")
    Else
        isBold = False
    End If
    
    If InStr(Text, "\itali") > 0 Then
        isItalic = True
        Text = Replace(Text, "\itali", "")
    Else
        isItalic = False
    End If
    
    posSize = InStr(Text, "\size:")
    If posSize > 0 Then
        sizeStr = ""
        For i = posSize + 6 To Len(Text)
            Dim zeichen As String
            zeichen = Mid$(Text, i, 1)
            If IsNumeric(zeichen) Or zeichen = "," Or zeichen = "." Then
                sizeStr = sizeStr & zeichen
            Else
                Exit For
            End If
        Next i
        
        If IsNumeric(sizeStr) Then
            fontSize = CSng(Val(Replace(sizeStr, ",", ".")))
        End If
        Text = Replace(Text, "\size:" & sizeStr, "")
    End If
    
    currentRange.Text = Text
    currentRange.Font.bold = isBold
    currentRange.Font.Italic = isItalic
    If fontSize > 0 Then
        currentRange.Font.size = fontSize
    End If
    
    ThisDocument.Bookmarks.Add Name:=Key_, Range:=currentRange
End Property

Sub DebugPrintBookmarkConst()
    Dim i As Long
    
    Debug.Print "Const Bo_bold As String = ""\bold"""
    Debug.Print "Const Bo_itali As String = ""\itali"""
    Debug.Print "Const Bo_size As String = ""\size:"""
    
    For i = 1 To ThisDocument.Bookmarks.Count
        Debug.Print "Const Bo_" & ThisDocument.Bookmarks.Item(i).Name & " As String = """ & ThisDocument.Bookmarks.Item(i).Name & """"
    Next i
End Sub
