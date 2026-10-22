VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsJson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public JsonObjekt As cPyDicList

Public Property Get setObjectString() As String
    setObjectString = "Object"
End Property

Public Property Get setArrayString() As String
    setArrayString = "Array"
End Property

Sub Parse(jsonText_ As String)
    Set JsonObjekt = New cPyDicList
    JsonObjekt.MyTyp = "Object"
    LokalParse JsonObjekt, jsonText_, 0, "Object"
End Sub

Sub LokalParse(CurrentContainer As cPyDicList, jsonText_ As String, EndPos_ As Long, ObjektType_ As String)
    Dim EndPosContainer As Long
    Dim EndPosValue As Long
    
    Dim MyType As String
    Dim currentJsonText As String
    Dim KeyName As String
    Dim ArreyIndex As Long
    
    
    ArreyIndex = -1
    CurrentContainer.MyTyp = ObjektType_
    
    ' 1) Inneren Text holen + Container-Ende merken (] oder })
    currentJsonText = GetInnerObjectText(jsonText_, ObjektType_, EndPosContainer)
    
    Do While Len(currentJsonText) > 0
        
        ' 2) Vorne Komma/Whitespace weg
        Do While Len(currentJsonText) > 0
            Select Case Left$(currentJsonText, 1)
                Case " ", vbTab, vbCr, vbLf, ","
                    currentJsonText = Mid$(currentJsonText, 2)
                Case Else
                    Exit Do
            End Select
        Loop
        If Len(currentJsonText) = 0 Then Exit Do
        
        ' 3) Key + Typ bestimmen und auf Value-Start positionieren
        If ObjektType_ = "Array" Then
            MyType = GetTyp(currentJsonText)
            ArreyIndex = ArreyIndex + 1
            KeyName = CStr(ArreyIndex)
        Else
            KeyName = GetKey(currentJsonText, EndPosValue, MyType)
            currentJsonText = Mid$(currentJsonText, EndPosValue)
        End If
        
        ' 4) Value lesen und direkt per .Append in die cPyDicList schreiben
        Select Case MyType
        
            Case "String"
                CurrentContainer.Append KeyName, GetInnerStringText(currentJsonText, EndPosValue)
                
            Case "Number"
                CurrentContainer.Append KeyName, GetInnerNumber(currentJsonText, EndPosValue)
                
            Case "Array"
                CurrentContainer.Append KeyName, New cPyDicList
                CurrentContainer.LastValue.MyTyp = "Array"
                LokalParse CurrentContainer.LastValue, currentJsonText, EndPosValue, "Array"
            
            Case "Object"
                CurrentContainer.Append KeyName, New cPyDicList
                CurrentContainer.LastValue.MyTyp = "Object"
                LokalParse CurrentContainer.LastValue, currentJsonText, EndPosValue, "Object"
            
            Case "Boolean"
                CurrentContainer.Append KeyName, GetInnerBoolean(currentJsonText, EndPosValue)
            
            Case "Null"
                CurrentContainer.Append KeyName, GetInnerNull(currentJsonText, EndPosValue)
                
        End Select
        
        ' 5) currentJsonText hinter dem Value weiter setzen
        If Len(currentJsonText) <= EndPosValue Then
            currentJsonText = ""
        Else
            If Mid$(currentJsonText, EndPosValue + 1, 1) = "," Then
                currentJsonText = Mid$(currentJsonText, EndPosValue + 2)
            Else
                currentJsonText = Mid$(currentJsonText, EndPosValue + 1)
            End If
        End If
        
    Loop
    
    ' 6) Output setzen
    EndPos_ = EndPosContainer
End Sub


' ==========================================================
' Gibt den Inhalt zwischen dem äußeren { ... } zurück
' - berücksichtigt Verschachtelung: { { } }
' - ignoriert Klammern innerhalb von Strings: " ... } ... "
' ==========================================================
Function GetInnerObjectText(ByVal jsonText_ As String, Type_ As String, EndPos_ As Long) As String
    Dim i As Long
    Dim j As Long
    Dim maxI As Long
    Dim StartPos As Long
    Dim CountOfBracket As Long
    Dim StarBracket As String
    Dim EndBracket As String
    Dim CurrentSymbol As String
    Dim QuotationMarksOpen As Boolean
    
    Dim BackslashCount As Long
    Dim QuoteIsEscaped As Boolean
    
    jsonText_ = Trim$(jsonText_)
    maxI = Len(jsonText_)
    
    If Type_ = "Object" Then
        StarBracket = "{"
        EndBracket = "}"
    Else
        StarBracket = "["
        EndBracket = "]"
    End If
    
    For i = 1 To maxI
        CurrentSymbol = Mid$(jsonText_, i, 1)
        
        ' -------------------------
        ' Quotes behandeln: " nur toggeln, wenn NICHT escaped (\")
        ' -------------------------
        If CurrentSymbol = """" Then
            
            BackslashCount = 0
            j = i - 1
            
            Do While j >= 1 And Mid$(jsonText_, j, 1) = "\"
                BackslashCount = BackslashCount + 1
                j = j - 1
            Loop
            
            QuoteIsEscaped = (BackslashCount Mod 2 = 1)
            
            If Not QuoteIsEscaped Then
                QuotationMarksOpen = Not QuotationMarksOpen
            End If
            
            GoTo Go2NextI
        End If
        
        ' -------------------------
        ' Nur ausserhalb von Strings Klammern zählen
        ' -------------------------
        If Not QuotationMarksOpen Then
            
            If CurrentSymbol = StarBracket Then
                If CountOfBracket = 0 Then
                    StartPos = i
                    CountOfBracket = 1
                Else
                    CountOfBracket = CountOfBracket + 1
                End If
                GoTo Go2NextI
            End If
            
            If CurrentSymbol = EndBracket Then
                CountOfBracket = CountOfBracket - 1
                If CountOfBracket = 0 Then
                    EndPos_ = i
                    Exit For
                End If
            End If
            
        End If
        
Go2NextI:
    Next i
    
    ' ==========================
    ' SICHERHEIT: soll krachen!
    ' ==========================
    If StartPos = 0 Then
        Err.Raise vbObjectError + 610, "clsJson.GetInnerObjectText", _
                  "Kein StartBracket '" & StarBracket & "' gefunden. Type_=" & Type_
    End If
    
    If EndPos_ = 0 Then
        Err.Raise vbObjectError + 611, "clsJson.GetInnerObjectText", _
                  "Kein passendes EndBracket '" & EndBracket & "' gefunden (JSON unvollständig/kaputt). Type_=" & Type_
    End If
    
    If EndPos_ <= StartPos Then
        Err.Raise vbObjectError + 612, "clsJson.GetInnerObjectText", _
                  "EndPos_ <= StartPos (kaputter JSON). StartPos=" & StartPos & ", EndPos=" & EndPos_
    End If
    
    If QuotationMarksOpen Then
        Err.Raise vbObjectError + 613, "clsJson.GetInnerObjectText", _
                  "String nicht geschlossen (QuotationMarksOpen=True). JSON kaputt."
    End If
    
    If CountOfBracket <> 0 Then
        Err.Raise vbObjectError + 614, "clsJson.GetInnerObjectText", _
                  "Bracket-Zähler am Ende nicht 0 (CountOfBracket=" & CountOfBracket & "). JSON kaputt."
    End If
    
    ' Innen zurückgeben (ohne äussere Klammern)
    GetInnerObjectText = Mid$(jsonText_, StartPos + 1, EndPos_ - StartPos - 1)
End Function

Function GetInnerStringText(ByVal jsonText_ As String, EndPos_ As Long) As String
    Dim i As Long
    Dim j As Long
    Dim maxI As Long
    
    Dim StartPos As Long
    Dim CurrentSymbol As String
    
    Dim BackslashCount As Long
    Dim QuoteIsEscaped As Boolean
    
    jsonText_ = Trim$(jsonText_)
    maxI = Len(jsonText_)
    
    EndPos_ = 0
    GetInnerStringText = ""
    
    If maxI = 0 Then Exit Function
    
    For i = 1 To maxI
        CurrentSymbol = Mid$(jsonText_, i, 1)
        
        ' -------------------------------------------------------
        ' 1) Start-Quote suchen (erstes nicht-escaped " ist Start)
        ' -------------------------------------------------------
        If StartPos = 0 Then
            
            ' Whitespaces ignorieren (optional)
            If CurrentSymbol = " " Or CurrentSymbol = vbTab Or CurrentSymbol = vbCr Or CurrentSymbol = vbLf Then
                GoTo Go2NextI
            End If
            
            If CurrentSymbol = """" Then
                StartPos = i
            Else
                ' wenn du NUR Whitespace vor dem String erlaubst:
                ' Exit Function
                ' oder einfach weiter suchen:
                GoTo Go2NextI
            End If
            
            GoTo Go2NextI
        End If
        
        ' -------------------------------------------------------
        ' 2) End-Quote suchen (\" beachten)
        ' -------------------------------------------------------
        If CurrentSymbol = """" Then
            
            BackslashCount = 0
            j = i - 1
            
            Do While j >= 1 And Mid$(jsonText_, j, 1) = "\"
                BackslashCount = BackslashCount + 1
                j = j - 1
            Loop
            
            QuoteIsEscaped = (BackslashCount Mod 2 = 1)
            
            If Not QuoteIsEscaped Then
                EndPos_ = i
                Exit For
            End If
        End If
        
Go2NextI:
    Next i
    
    ' -------------------------------------------------------
    ' 3) Wenn kein Ende gefunden -> Fehler (wie gewünscht)
    ' -------------------------------------------------------
    If StartPos = 0 Then Exit Function
    
    If EndPos_ = 0 Then
        Err.Raise vbObjectError + 640, "clsJson.GetInnerStringText", _
                  "String-Ende (schließendes "") nicht gefunden."
    End If
    
    ' -------------------------------------------------------
    ' 4) Inhalt ohne Quotes zurückgeben
    ' -------------------------------------------------------
    GetInnerStringText = Mid$(jsonText_, StartPos + 1, EndPos_ - StartPos - 1)
End Function

Function GetInnerNumber(ByVal jsonText_ As String, EndPos_ As Long) As Double
    Dim Token As String
    Dim pComma As Long

    GetInnerNumber = 0
    EndPos_ = 0

    pComma = InStr(1, jsonText_, ",")
    If pComma = 0 Then
        ' keine Komma -> Zahl geht bis Ende des aktuellen Substrings
        Token = Trim$(jsonText_)
        EndPos_ = Len(jsonText_) ' letztes Zeichen
    Else
        Token = Mid$(jsonText_, 1, pComma - 1)
        Token = Trim$(Token)
        EndPos_ = pComma - 1 ' letztes Zeichen vor Komma
    End If

    GetInnerNumber = Val(Token)
End Function

Function GetInnerBoolean(ByVal jsonText_ As String, EndPos_ As Long) As Boolean
    Dim Token As String
    Dim pComma As Long
    
    EndPos_ = 0
    GetInnerBoolean = False
    
    pComma = InStr(1, jsonText_, ",")
    
    If pComma = 0 Then
        ' kein Komma -> Token geht bis Ende des aktuellen Substrings
        Token = Trim$(jsonText_)
        EndPos_ = Len(jsonText_)
    Else
        Token = Mid$(jsonText_, 1, pComma - 1)
        Token = Trim$(Token)
        EndPos_ = pComma - 1
    End If
    
    Select Case LCase$(Token)
        Case "true"
            GetInnerBoolean = True
        Case "false"
            GetInnerBoolean = False
        Case Else
            Err.Raise vbObjectError + 670, "clsJson.GetInnerBoolean", "Boolean-Token ungültig: " & Token
    End Select
End Function

Function GetInnerNull(ByVal jsonText_ As String, EndPos_ As Long) As Variant
    Dim Token As String
    Dim pComma As Long
    
    EndPos_ = 0
    GetInnerNull = Null
    
    pComma = InStr(1, jsonText_, ",")
    
    If pComma = 0 Then
        Token = Trim$(jsonText_)
        EndPos_ = Len(jsonText_)
    Else
        Token = Mid$(jsonText_, 1, pComma - 1)
        Token = Trim$(Token)
        EndPos_ = pComma - 1
    End If
    
    If LCase$(Token) <> "null" Then
        Err.Raise vbObjectError + 680, "clsJson.GetInnerNull", "Null-Token ungültig: " & Token
    End If
End Function

Function GetKey(ByVal jsonText_ As String, EndPos_ As Long, ValueType_ As String) As String
    Dim i As Long
    Dim j As Long
    Dim maxI As Long
    
    Dim KeyStartPos As Long
    Dim KeyEndPos As Long
    
    Dim CurrentSymbol As String
    Dim BackslashCount As Long
    Dim QuoteIsEscaped As Boolean
    
    jsonText_ = Trim$(jsonText_)
    maxI = Len(jsonText_)
    
    GetKey = ""
    EndPos_ = 0
    ValueType_ = ""
    
    If maxI = 0 Then Exit Function
    
    ' -------------------------------------------------------
    ' 1) Key-Startquote " suchen
    ' -------------------------------------------------------
    For i = 1 To maxI
        CurrentSymbol = Mid$(jsonText_, i, 1)
        
        If CurrentSymbol = " " Or CurrentSymbol = vbTab Or CurrentSymbol = vbCr Or CurrentSymbol = vbLf Then
            GoTo Go2NextI_1
        End If
        
        If CurrentSymbol = """" Then
            KeyStartPos = i
            Exit For
        End If
        
Go2NextI_1:
    Next i
    
    If KeyStartPos = 0 Then Exit Function
    
    ' -------------------------------------------------------
    ' 2) Key-Endquote suchen (\" beachten)
    ' -------------------------------------------------------
    For i = KeyStartPos + 1 To maxI
        CurrentSymbol = Mid$(jsonText_, i, 1)
        
        If CurrentSymbol = """" Then
            BackslashCount = 0
            j = i - 1
            
            Do While j >= 1 And Mid$(jsonText_, j, 1) = "\"
                BackslashCount = BackslashCount + 1
                j = j - 1
            Loop
            
            QuoteIsEscaped = (BackslashCount Mod 2 = 1)
            
            If Not QuoteIsEscaped Then
                KeyEndPos = i
                Exit For
            End If
        End If
    Next i
    
    If KeyEndPos = 0 Then
        Err.Raise vbObjectError + 650, "clsJson.GetKey", "Key-Ende (schließendes "") nicht gefunden."
    End If
    
    GetKey = Mid$(jsonText_, KeyStartPos + 1, KeyEndPos - KeyStartPos - 1)
    
    ' -------------------------------------------------------
    ' 3) ":" nach dem Key suchen
    ' -------------------------------------------------------
    For i = KeyEndPos + 1 To maxI
        CurrentSymbol = Mid$(jsonText_, i, 1)
        
        If CurrentSymbol = " " Or CurrentSymbol = vbTab Or CurrentSymbol = vbCr Or CurrentSymbol = vbLf Then
            GoTo Go2NextI_3
        End If
        
        If CurrentSymbol = ":" Then
            Exit For
        End If
        
        ' sobald was anderes als whitespace kommt und nicht ":" -> kaputt
        Err.Raise vbObjectError + 651, "clsJson.GetKey", "Kein ':' nach Key gefunden."
        
Go2NextI_3:
    Next i
    
    If i > maxI Or Mid$(jsonText_, i, 1) <> ":" Then
        Err.Raise vbObjectError + 651, "clsJson.GetKey", "Kein ':' nach Key gefunden."
    End If
    
    ' -------------------------------------------------------
    ' 4) Value-Startposition finden (Whitespace nach ":" skippen)
    ' -------------------------------------------------------
    For i = i + 1 To maxI
        CurrentSymbol = Mid$(jsonText_, i, 1)
        
        If CurrentSymbol = " " Or CurrentSymbol = vbTab Or CurrentSymbol = vbCr Or CurrentSymbol = vbLf Then
            GoTo Go2NextI_4
        End If
        
        EndPos_ = i
        Exit For
        
Go2NextI_4:
    Next i
    
    If EndPos_ = 0 Then
        Err.Raise vbObjectError + 652, "clsJson.GetKey", "Kein Value nach ':' gefunden."
    End If
    
    ' -------------------------------------------------------
    ' 5) ValueType_ bestimmen (nur über 1. Zeichen)
    ' -------------------------------------------------------
    CurrentSymbol = Mid$(jsonText_, EndPos_, 1)
    
    If CurrentSymbol = "{" Then
        ValueType_ = "Object"
    ElseIf CurrentSymbol = "[" Then
        ValueType_ = "Array"
    ElseIf CurrentSymbol = """" Then
        ValueType_ = "String"
    ElseIf (CurrentSymbol >= "0" And CurrentSymbol <= "9") Or CurrentSymbol = "-" Then
        ValueType_ = "Number"
    ElseIf CurrentSymbol = "t" Or CurrentSymbol = "f" Then
        ValueType_ = "Boolean"
    ElseIf CurrentSymbol = "n" Then
        ValueType_ = "Null"
    Else
        ValueType_ = "Unknown"
    End If
End Function

Function GetTyp(ByVal jsonText_ As String) As String
    Dim i As Long
    Dim maxI As Long
    Dim CurrentSymbol As String
    
    GetTyp = "Unknown"
    
    maxI = Len(jsonText_)
    If maxI = 0 Then Exit Function
    
    ' erstes relevantes Zeichen suchen (Whitespace skip)
    For i = 1 To maxI
        CurrentSymbol = Mid$(jsonText_, i, 1)
        
        If CurrentSymbol = " " Or CurrentSymbol = vbTab Or CurrentSymbol = vbCr Or CurrentSymbol = vbLf Then
            GoTo Go2NextI
        End If
        
        Exit For
        
Go2NextI:
    Next i
    
    If i > maxI Then Exit Function
    
    CurrentSymbol = Mid$(jsonText_, i, 1)
    
    ' nur Typ setzen und fertig
    If CurrentSymbol = """" Then
        GetTyp = "String"
        Exit Function
    End If
    
    If CurrentSymbol = "{" Then
        GetTyp = "Object"
        Exit Function
    End If
    
    If CurrentSymbol = "[" Then
        GetTyp = "Array"
        Exit Function
    End If
    
    If (CurrentSymbol >= "0" And CurrentSymbol <= "9") Or CurrentSymbol = "-" Then
        GetTyp = "Number"
        Exit Function
    End If
    
    If CurrentSymbol = "t" Or CurrentSymbol = "f" Then
        GetTyp = "Boolean"
        Exit Function
    End If
    
    If CurrentSymbol = "n" Then
        GetTyp = "Null"
        Exit Function
    End If
End Function

' ==============================================================================
' Startfunktion für den JSON-Export (wird vom User aufgerufen)
' ==============================================================================
Public Function ToJson() As String
    ToJson = BuildJsonRecursive(JsonObjekt)
End Function

' ==============================================================================
' DIE ECHTE REKURSIVE FUNKTION (ruft sich selbst für Unterobjekte auf)
' ==============================================================================
Private Function BuildJsonRecursive(CurrentContainer_ As cPyDicList) As String
    Dim i As Long
    Dim OutStr As String
    Dim Val As Variant
    Dim ElementStr As String
    Dim CleanStr As String
    
    OutStr = ""
    ' Schleife über alle Elemente des aktuell übergebenen Containers
    For i = 0 To CurrentContainer_.Imax
        
        ElementStr = ""
        
        ' REKURSIONSSCHRITT: Wenn der Wert ein weiteres cPyDicList-Objekt ist
        If IsObject(CurrentContainer_.ValueIndex(i)) Then
            ElementStr = BuildJsonRecursive(CurrentContainer_.ValueIndex(i))
        Else
            Val = CurrentContainer_.ValueIndex(i)
            ' Basis-Datentypen konvertieren
            Select Case VarType(Val)
                Case vbString
                    CleanStr = Replace(CStr(Val), "\", "\\")
                    CleanStr = Replace(CleanStr, """", "\""")
                    ElementStr = """" & CleanStr & """"
                    
                Case vbBoolean
                    If Val = True Then
                        ElementStr = "true"
                    Else
                        ElementStr = "false"
                    End If
                    
                Case vbNull, vbEmpty
                    ElementStr = "null"
                    
                Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
                    ElementStr = Str(Val)
                    ElementStr = Trim$(Replace(ElementStr, ",", "."))
                    
                Case Else
                    ElementStr = """" & CStr(Val) & """"
            End Select
        End If
        
        ' Steuerung über die MyTyp-Eigenschaft des aktuellen Rekursions-Containers
        If CurrentContainer_.MyTyp = "Array" Then
            OutStr = OutStr & ElementStr
        Else
            OutStr = OutStr & """" & CurrentContainer_.Key(i) & """:" & ElementStr
        End If
        
        If i < CurrentContainer_.Imax Then
            OutStr = OutStr & ","
        End If
    Next i
    
    ' Klammern für die aktuelle Rekursions-Ebene schließen
    If CurrentContainer_.MyTyp = "Array" Then
        BuildJsonRecursive = "[" & OutStr & "]"
    Else
        BuildJsonRecursive = "{" & OutStr & "}"
    End If
End Function

Sub Save(FilePath_ As String)
    Dim jsonString As String
    jsonString = ToJson
    
    #If Mac Then
        ' === MAC-VERSION (UTF-8 ohne BOM) ===
        Dim fileNum As Long
        Dim textBytes() As Byte
        
        ' Wir nutzen StrConv, um den String in ein Byte-Array umzuwandeln.
        ' Da Mac-VBA intern mit UTF-8/Unicode arbeitet, schreiben wir die Bytes direkt.
        textBytes = StrConv(jsonString, vbFromUnicode)
        
        fileNum = FreeFile
        Open FilePath_ For Binary Access Write As #fileNum
            Put #fileNum, , textBytes
        Close #fileNum
    #Else
        ' === WINDOWS-VERSION (Dein bewährter ADODB-Trick) ===
        Dim File As Object
        Dim oStreamBinary As Object
        
        Set File = CreateObject("ADODB.Stream")
        File.Type = 2
        File.Charset = "UTF-8"
        File.Open
        File.WriteText jsonString
        
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
    #End If
End Sub


