VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPyXmltodict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' ==============================================================================
' CLASS: cPyXmltodict
' DEVELOPER: Franz Hardinger
' HOMEPAGE / YOUTUBE: https://www.youtube.com/watch?v=DABjl62x0bY
' GITHUB: https://github.com/FranzH85/pyVBA/tree/main/VBA_XmlParsing
'
' REQUIREMENTS (WINDOWS ONLY):
' Uses Late Binding for MSXML2.DOMDocument.6.0 and ADODB.Stream.
' No manual VBA References (Tools -> References) required!
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
' A highly optimized XML parser and exporter for Windows VBA that operates
' natively on top of the 'cPyDicList' structure to provide Python-like
' dictionary syntax with strict element sequence preservation.
' ==============================================================================

Option Explicit

Private Const VBA_ElementDelete As String = "VBA_ElementDelete"
Private Const VBA_TextMarker   As String = "#text"

' Native MSXML NodeType constants for Late Binding
Private Const NODE_TEXT As Long = 3
Private Const NODE_COMMENT As Long = 8

Public XmlObjekt As cPyDicList

Public Property Get setElementDeleteString() As String
    setElementDeleteString = VBA_ElementDelete
End Property

Public Property Get setTextMarkerString() As String
    setTextMarkerString = VBA_TextMarker
End Property

Function ParseFile(FilePath_ As String) As Boolean
    Dim XmlDoc As Object
    Set XmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    
    Set XmlObjekt = New cPyDicList
    
    XmlDoc.async = False
    XmlDoc.validateOnParse = False
    
    If XmlDoc.Load(FilePath_) Then
        ParseFile = True
        LocalParse XmlObjekt, XmlDoc.ChildNodes
    Else
        MsgBox "Fehler beim Laden der XML!", vbCritical
        ParseFile = False
    End If
End Function

Sub Parse(XmlText_ As String)
    Dim XmlDoc As Object
    Set XmlDoc = CreateObject("MSXML2.DOMDocument.6.0")

    Set XmlObjekt = New cPyDicList
    
    XmlDoc.async = False
    XmlDoc.validateOnParse = False
    XmlDoc.LoadXML XmlText_
    
    XmlObjekt.NewPyXMLSyndax = True
    LocalParse XmlObjekt, XmlDoc.ChildNodes
End Sub

Private Sub LocalParse(CurrentContainer_ As cPyDicList, lcNodes_ As Object)
    Dim I As Long
    Dim Imax As Long
    Dim I2 As Long
    Dim I2max As Long
    Dim IsChildrenList As Boolean
    Dim MaxICurrentContainer As Long
    Dim I_CurrentContainer As Long

    Imax = lcNodes_.Length - 1
    
    CurrentContainer_.NewPyXMLSyndax = True

    If Imax > -1 Then
        MaxICurrentContainer = -1
        For I = 0 To Imax
            If lcNodes_.Item(I).NodeType <> NODE_COMMENT Then
                MaxICurrentContainer = MaxICurrentContainer + 1
            End If
        Next I
        
        CurrentContainer_.SetNewMaxIndex MaxICurrentContainer
        
        I_CurrentContainer = -1
        
        For I = 0 To Imax
            I_CurrentContainer = I_CurrentContainer + 1
            
            If lcNodes_.Item(I).NodeType = NODE_COMMENT Then
                I_CurrentContainer = I_CurrentContainer - 1
                GoTo NextI
            End If
            
            If lcNodes_.Item(I).NodeType = NODE_TEXT Then
                CurrentContainer_.UpdateWithIndex I_CurrentContainer, lcNodes_.Item(I).nodeName, lcNodes_.Item(I).NodeValue
            Else
            
                If lcNodes_.Item(I).ChildNodes.Length > 0 Then
                    IsChildrenList = True
                    
                    If lcNodes_.Item(I).ChildNodes.Length = 1 Then
                        If lcNodes_.Item(I).ChildNodes.Item(0).NodeType = NODE_TEXT Then
                            IsChildrenList = False
                            CurrentContainer_.UpdateWithIndex I_CurrentContainer, lcNodes_.Item(I).nodeName, New cPySimpleObject
                            CurrentContainer_.ValueIndex(I_CurrentContainer) = lcNodes_.Item(I).ChildNodes.Item(0).nodeTypedValue
                        End If
                    End If
                    
                    If IsChildrenList Then
                        CurrentContainer_.UpdateWithIndex I_CurrentContainer, lcNodes_.Item(I).nodeName, New cPyDicList
                        LocalParse CurrentContainer_.ValueIndex(I_CurrentContainer), lcNodes_.Item(I).ChildNodes
                    End If
                Else
                    CurrentContainer_.UpdateWithIndex I_CurrentContainer, lcNodes_.Item(I).nodeName, New cPySimpleObject
                End If
                
                I2max = lcNodes_.Item(I).Attributes.Length - 1
                
                If I2max > -1 Then
                    CurrentContainer_.ValueIndex(I_CurrentContainer).MyTyp = New cPyDicList
                    
                    CurrentContainer_.ValueIndex(I_CurrentContainer).MyTyp.SetNewMaxIndex I2max
                    
                    For I2 = 0 To I2max
                        CurrentContainer_.ValueIndex(I_CurrentContainer).MyTyp.UpdateWithIndex I2, lcNodes_.Item(I).Attributes.Item(I2).nodeName, lcNodes_.Item(I).Attributes.Item(I2).NodeValue
                    Next I2
                End If
            End If

NextI:
        Next I
    End If
End Sub

Function CreateXmlNode(XmlString As String) As Variant
    Dim XmlStruct As Variant
    
    If InStr(XmlString, "<") <> 0 Then
        Set XmlStruct = New cPyXmltodict
        XmlStruct.Parse XmlString
        Set CreateXmlNode = XmlStruct.XmlObjekt
    Else
        Set XmlStruct = New cPySimpleObject
        XmlStruct.Value = XmlString
        Set CreateXmlNode = XmlStruct
    End If
End Function

Function toXml() As String
    Dim XmlDoc As String
    
    LocaltoXml XmlDoc
    toXml = XmlDoc
End Function

Sub toXmlFile(FilePath_ As String)
    Dim XmlDoc As String
    Dim File As Object
    Dim oStreamBinary As Object

    LocaltoXml XmlDoc

    Set File = CreateObject("ADODB.Stream")
    File.Type = 2
    File.Charset = "UTF-8"
    File.Open
    File.WriteText XmlDoc
    
    Set oStreamBinary = CreateObject("ADODB.Stream")
    oStreamBinary.Type = 1
    oStreamBinary.Open
    File.Position = 3 ' BOM abschneiden
    File.CopyTo oStreamBinary
    File.Flush
    File.Close
    
    oStreamBinary.SaveToFile FilePath_, 2 ' 2 = adSaveCreateOverWrite
    oStreamBinary.Flush
    oStreamBinary.Close
    Set oStreamBinary = Nothing
End Sub

Private Sub LocaltoXml(xmlDoc_ As String)
    If XmlObjekt.Key(0) = "XML" Or XmlObjekt.Key(0) = "xml" Then
        xmlDoc_ = "<?xml" & localWriteAttributes(XmlObjekt.ValueIndex(0).MyTyp) & "?>" & Chr(13)
        XmlObjekt.Key(0) = VBA_ElementDelete
    End If
    
    LocaltoXmlNextElement XmlObjekt, xmlDoc_, 0
End Sub

Private Sub LocaltoXmlNextElement(CurrentContainer_ As cPyDicList, Element_ As String, vbTabCount_ As Long)
    Dim I  As Long
    Dim vbTabString As String
    
    For I = 1 To vbTabCount_
        vbTabString = vbTabString & vbTab
    Next I
    
    For I = 0 To CurrentContainer_.Imax
        If CurrentContainer_.Key(I) <> VBA_ElementDelete Then
            If CurrentContainer_.Key(I) = VBA_TextMarker Then
                Element_ = Element_ & CurrentContainer_.ValueIndex(I)
                GoTo NextI
            End If
            
            If TypeOf CurrentContainer_.ValueIndex(I) Is cPyDicList Then
                If CurrentContainer_.ValueIndex(I).Imax = -1 Then
                    Element_ = Element_ & vbTabString & "<" & CurrentContainer_.Key(I) & localWriteAttributes(CurrentContainer_.ValueIndex(I).MyTyp) & "/>" & Chr(13)
                Else
                    Element_ = Element_ & vbTabString & "<" & CurrentContainer_.Key(I) & localWriteAttributes(CurrentContainer_.ValueIndex(I).MyTyp) & ">" & Chr(13)
                    LocaltoXmlNextElement CurrentContainer_.ValueIndex(I), Element_, vbTabCount_ + 1
                    Element_ = Element_ & vbTabString & "</" & CurrentContainer_.Key(I) & ">" & Chr(13)
                End If
            Else
                If IsEmpty(CurrentContainer_.ValueIndex(I)) Then
                    Element_ = Element_ & vbTabString & "<" & CurrentContainer_.Key(I) & localWriteAttributes(CurrentContainer_.ValueIndex(I).MyTyp) & "/>" & Chr(13)
                Else
                    Element_ = Element_ & vbTabString & "<" & CurrentContainer_.Key(I) & localWriteAttributes(CurrentContainer_.ValueIndex(I).MyTyp) & ">" & CurrentContainer_.ValueIndex(I) & "</" & CurrentContainer_.Key(I) & ">" & Chr(13)
                End If
            End If
        End If
NextI:
    Next I
End Sub

Private Function localWriteAttributes(Attributes_ As Variant) As String
    Dim localStrng As String
    Dim I As Long
    
    If Not IsEmpty(Attributes_) Then
        If IsObject(Attributes_) Then
            For I = 0 To Attributes_.Imax
                localStrng = localStrng & " " & Attributes_.Key(I) & "=""" & Attributes_.ValueIndex(I) & """"
            Next I
        End If
    End If
    
    localWriteAttributes = localStrng
End Function

Sub DebugPrintKeyConst()
    Dim KeyNameList As New cPyDicList
    Dim I As Long
    
    Debug.Print "Const VBA_ElementDelete As String = ""VBA_ElementDelete"""
    Debug.Print "Const VBA_TextMarker   As String = ""#text"""
    
    NextDebugPrintKeyConst XmlObjekt, KeyNameList
    
    For I = 0 To KeyNameList.Imax
        Debug.Print "Const xml_" & Replace(KeyNameList.Key(I), ":", "_") & " As String = """ & KeyNameList.Key(I) & """"
    Next I
End Sub

Private Sub NextDebugPrintKeyConst(CurrentContainer_ As cPyDicList, KeyNameList_ As cPyDicList)
    Dim I As Long
    
    For I = 0 To CurrentContainer_.Imax
        KeyNameList_.Update CurrentContainer_.Key(I), Empty
        If TypeOf CurrentContainer_.ValueIndex(I) Is cPyDicList Then
            NextDebugPrintKeyConst CurrentContainer_.ValueIndex(I), KeyNameList_
        End If
    Next I
End Sub
