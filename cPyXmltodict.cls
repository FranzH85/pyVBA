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
' HOMEPAGE / YOUTUBE: "Python dic list stile in vba"
' GITHUB: https://github.com
'
' REQUIREMENTS (WINDOWS ONLY):
' This module requires the native Windows XML library for maximum performance.
' In the VBA Editor, go to Tools -> References and enable:
' "Microsoft XML, v6.0" (MSXML2)
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

Public XmlObjekt As cPyDicList

Public Property Get setElementDeleteString() As String
    setElementDeleteString = VBA_ElementDelete
End Property

Public Property Get setTextMarkerString() As String
    setTextMarkerString = VBA_TextMarker
End Property


Function ParseFile(FilePath_ As String) As Boolean
    Dim xmlDoc As New MSXML2.DOMDocument60
    
    Set XmlObjekt = New cPyDicList
    
    xmlDoc.async = False
    xmlDoc.validateOnParse = False
    
    If xmlDoc.Load(FilePath_) Then
        ParseFile = True
        LocalParse XmlObjekt, xmlDoc.ChildNodes
    Else
        MsgBox "Fehler beim Laden der XML!", vbCritical
        ParseFile = False
    End If
End Function

Sub Parse(XmlText_ As String)
    Dim xmlDoc As New MSXML2.DOMDocument60

    Set XmlObjekt = New cPyDicList
    
    xmlDoc.async = False
    xmlDoc.validateOnParse = False
    xmlDoc.LoadXML XmlText_
    
    XmlObjekt.NewPyXMLSyndax = True
    LocalParse XmlObjekt, xmlDoc.ChildNodes
   
End Sub

Private Sub LocalParse(CurrentContainer_ As cPyDicList, lcNodes_ As MSXML2.IXMLDOMNodeList)
    Dim i As Long
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
        For i = 0 To Imax
            If lcNodes_.Item(i).NodeType <> NODE_COMMENT Then
                MaxICurrentContainer = MaxICurrentContainer + 1
            End If
        Next i
        
        CurrentContainer_.SetNewMaxIndex MaxICurrentContainer
        
        I_CurrentContainer = -1
        
        For i = 0 To Imax
            I_CurrentContainer = I_CurrentContainer + 1
            
            If lcNodes_.Item(i).NodeType = NODE_COMMENT Then
                I_CurrentContainer = I_CurrentContainer - 1
                GoTo NextI
            End If
            
            If lcNodes_.Item(i).NodeType = NODE_TEXT Then
                CurrentContainer_.UpdateWithIndex I_CurrentContainer, lcNodes_.Item(i).nodeName, lcNodes_.Item(i).NodeValue
            Else
            
                If lcNodes_.Item(i).ChildNodes.Length > 0 Then
                    IsChildrenList = True
                    
                    If lcNodes_.Item(i).ChildNodes.Length = 1 Then
                        If lcNodes_.Item(i).ChildNodes.Item(0).NodeType = NODE_TEXT Then
                            IsChildrenList = False
                            CurrentContainer_.UpdateWithIndex I_CurrentContainer, lcNodes_.Item(i).nodeName, New cPySimpleObject
                            CurrentContainer_.ValueIndex(I_CurrentContainer) = lcNodes_.Item(i).ChildNodes.Item(0).nodeTypedValue
                        End If
                    End If
                    
                    If IsChildrenList Then
                        CurrentContainer_.UpdateWithIndex I_CurrentContainer, lcNodes_.Item(i).nodeName, New cPyDicList
                        LocalParse CurrentContainer_.ValueIndex(I_CurrentContainer), lcNodes_.Item(i).ChildNodes
                    End If
                Else
                    CurrentContainer_.UpdateWithIndex I_CurrentContainer, lcNodes_.Item(i).nodeName, New cPySimpleObject
                End If
                
                I2max = lcNodes_.Item(i).Attributes.Length - 1
                
                If I2max > -1 Then
                    CurrentContainer_.ValueIndex(I_CurrentContainer).MyTyp = New cPyDicList
                    
                    CurrentContainer_.ValueIndex(I_CurrentContainer).MyTyp.SetNewMaxIndex I2max
                    
                    For I2 = 0 To I2max
                        CurrentContainer_.ValueIndex(I_CurrentContainer).MyTyp.UpdateWithIndex I2, lcNodes_.Item(i).Attributes.Item(I2).nodeName, lcNodes_.Item(i).Attributes.Item(I2).NodeValue
                    Next I2
                End If
            End If

NextI:
        Next i
    End If

End Sub

Function toXml() As String
    Dim xmlDoc As String
    
    LocaltoXml xmlDoc
    toXml = xmlDoc
End Function

Sub toXmlFile(FilePath_ As String)
    Dim xmlDoc As String
    Dim File As Object
    Dim oStreamBinary As Object

    LocaltoXml xmlDoc

    Set File = CreateObject("ADODB.Stream")
    File.Type = 2
    File.Charset = "UTF-8"
    File.Open
    File.WriteText xmlDoc
    
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
    Dim i  As Long
    Dim vbTabString As String
    
    For i = 1 To vbTabCount_
        vbTabString = vbTabString & vbTab
    Next i
    
    For i = 0 To CurrentContainer_.Imax
        If CurrentContainer_.Key(i) <> VBA_ElementDelete Then
            If CurrentContainer_.Key(i) = VBA_TextMarker Then
                Element_ = Element_ & CurrentContainer_.ValueIndex(i)
                GoTo NextI
            End If
            
            If TypeOf CurrentContainer_.ValueIndex(i) Is cPyDicList Then
                If CurrentContainer_.ValueIndex(i).Imax = -1 Then
                    Element_ = Element_ & vbTabString & "<" & CurrentContainer_.Key(i) & localWriteAttributes(CurrentContainer_.ValueIndex(i).MyTyp) & "/>" & Chr(13)
                Else
                    Element_ = Element_ & vbTabString & "<" & CurrentContainer_.Key(i) & localWriteAttributes(CurrentContainer_.ValueIndex(i).MyTyp) & ">" & Chr(13)
                    LocaltoXmlNextElement CurrentContainer_.ValueIndex(i), Element_, vbTabCount_ + 1
                    Element_ = Element_ & vbTabString & "</" & CurrentContainer_.Key(i) & ">" & Chr(13)
                End If
            Else
                If IsEmpty(CurrentContainer_.ValueIndex(i)) Then
                    Element_ = Element_ & vbTabString & "<" & CurrentContainer_.Key(i) & localWriteAttributes(CurrentContainer_.ValueIndex(i).MyTyp) & "/>" & Chr(13)
                Else
                    Element_ = Element_ & vbTabString & "<" & CurrentContainer_.Key(i) & localWriteAttributes(CurrentContainer_.ValueIndex(i).MyTyp) & ">" & CurrentContainer_.ValueIndex(i) & "</" & CurrentContainer_.Key(i) & ">" & Chr(13)
                End If
            End If
        End If
NextI:
    Next i

End Sub

Private Function localWriteAttributes(Attributes_ As Variant) As String
    Dim localStrng As String
    Dim i As Long
    
    If Not IsEmpty(Attributes_) Then
        If IsObject(Attributes_) Then
            For i = 0 To Attributes_.Imax
                localStrng = localStrng & " " & Attributes_.Key(i) & "=""" & Attributes_.ValueIndex(i) & """"
            Next i
        End If
    End If
    
    localWriteAttributes = localStrng
End Function

Sub DebugPrintKeyConst()
    Dim KeyNameList As New cPyDicList
    Dim i As Long
    
    Debug.Print "Const VBA_ElementDelete As String = ""VBA_ElementDelete"""
    Debug.Print "Const VBA_TextMarker   As String = ""#text"""
    
    NextDebugPrintKeyConst XmlObjekt, KeyNameList
    
    For i = 0 To KeyNameList.Imax
        Debug.Print "Const xml_" & Replace(KeyNameList.Key(i), ":", "_") & " As String = """ & KeyNameList.Key(i) & """"
    Next i
End Sub

Private Sub NextDebugPrintKeyConst(CurrentContainer_ As cPyDicList, KeyNameList_ As cPyDicList)
    Dim i As Long
    
    For i = 0 To CurrentContainer_.Imax
        KeyNameList_.Update CurrentContainer_.Key(i), Empty
        If TypeOf CurrentContainer_.ValueIndex(i) Is cPyDicList Then
            NextDebugPrintKeyConst CurrentContainer_.ValueIndex(i), KeyNameList_
        End If
    Next i
End Sub


