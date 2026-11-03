VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPyWordTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' ==============================================================================
' CLASS: cPyWordTable
' DEVELOPER: Franz Hardinger
' HOMEPAGE / YOUTUBE: "Python dic list stile in vba"
' GITHUB: https://github.com
'
' REQUIREMENTS (WINDOWS ONLY):
' This module wraps Microsoft Word Table manipulation into a clean, matrix-style
' syntax. It requires a standard reference to the Microsoft Word Object Library.
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
' An elegant wrapper for Word Tables providing a 2D matrix-style index interface,
' runtime text formatting via string markup payload tags (\bold, \itali, \size:),
' advanced border styles, and cell manipulation tools.
' ==============================================================================

Option Explicit

Private TargetTable As Table

Sub Init(WdTable As Table)
    Set TargetTable = WdTable
End Sub

Public Property Get Value(Row_ As Long, Col_ As Long) As String
Attribute Value.VB_UserMemId = 0
    Dim cellText As String
    cellText = TargetTable.Cell(Row_, Col_).Range.Text
    Value = Left(cellText, Len(cellText) - 2)
End Property

Public Property Let Value(Row_ As Long, Col_ As Long, Text_ As String)
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
    
    Set currentRange = TargetTable.Cell(Row_, Col_).Range
    
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
End Property

Sub AddText(Row_ As Long, Col_ As Long, Text_ As String)
    Dim currentRange As Range
    Dim Text As String
    Dim i As Long
    Dim posSize As Long
    Dim sizeStr As String
    Dim isBold As Boolean
    Dim isItalic As Boolean
    Dim fontSize As Single
    
    Text = Text_
    
    Set currentRange = TargetTable.Cell(Row_, Col_).Range
    currentRange.Collapse Direction:=wdCollapseEnd
    currentRange.MoveEnd Unit:=wdCharacter, Count:=-1
    
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
    
    ' Text anhängen und exakt diesen neuen Bereich formatieren
    currentRange.Text = Text
    currentRange.Font.bold = isBold
    currentRange.Font.Italic = isItalic
    If fontSize > 0 Then
        currentRange.Font.size = fontSize
    End If
End Sub


 Sub DeleteRows(Optional StartRow As Long = 1)
    Dim i As Long
    
    For i = TargetTable.Rows.Count To StartRow Step -1
        TargetTable.Rows(i).Delete
    Next i
End Sub

Sub AddRows(Optional Count_ As Long = 1)
    Dim i As Long
    
    For i = 1 To Count_
        TargetTable.Rows.Add
    Next i
End Sub

Sub LineTop(Row_ As Long, Optional Col_ As Long = 0)
    Dim i As Long
    
    If Col_ > 0 Then
        With TargetTable.Cell(Row_, Col_).Borders(wdBorderTop)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth100pt
        End With
    Else
        For i = 1 To TargetTable.Columns.Count
            With TargetTable.Cell(Row_, i).Borders(wdBorderTop)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth100pt
            End With
        Next i
    End If
End Sub

Sub LineBottom(Row_ As Long, Optional Col_ As Long = 0)
    Dim i As Long
    
    If Col_ > 0 Then
        With TargetTable.Cell(Row_, Col_).Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth100pt
        End With
    Else
        For i = 1 To TargetTable.Columns.Count
            With TargetTable.Cell(Row_, i).Borders(wdBorderBottom)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth100pt
            End With
        Next i
    End If
End Sub

Sub DoubleLineBottom(Row_ As Long, Optional Col_ As Long = 0)
    Dim i As Long
    
    If Col_ > 0 Then
        With TargetTable.Cell(Row_, Col_).Borders(wdBorderBottom)
            .LineStyle = wdLineStyleDouble
            .LineWidth = wdLineWidth050pt
        End With
    Else
        For i = 1 To TargetTable.Columns.Count
            With TargetTable.Cell(Row_, i).Borders(wdBorderBottom)
                .LineStyle = wdLineStyleDouble
                .LineWidth = wdLineWidth050pt
            End With
        Next i
    End If
End Sub

Public Sub ClearBorders(Row_ As Long, Optional Col_ As Long = 0)
    Dim i As Long
    
    If Col_ > 0 Then
        TargetTable.Cell(Row_, Col_).Borders.OutsideLineStyle = wdLineStyleNone
    Else
        For i = 1 To TargetTable.Columns.Count
            TargetTable.Cell(Row_, i).Borders.OutsideLineStyle = wdLineStyleNone
        Next i
    End If
End Sub

Sub MergeCells(Row_ As Long, StartCol_ As Long, EndCol_ As Long)
    TargetTable.Cell(Row_, StartCol_).Merge TargetTable.Cell(Row_, EndCol_)
End Sub

Public Property Get Imax() As Long
    Imax = TargetTable.Rows.Count
End Property


