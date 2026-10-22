Attribute VB_Name = "Develop"
Option Explicit

Sub TestPyDicList1()
    Dim PyDicList As New cPyDicList
    
    PyDicList.Append "Captain", "Picard"
    PyDicList.Append "Schiff", New cPyDicList
    PyDicList.LastValue.Append "Name", "Enterprise"
    
    Debug.Print PyDicList("Schiff")("Name")
    PyDicList("Schiff")("Name") = "James T. Kirk"
    Debug.Print PyDicList("Schiff")("Name")
    
End Sub

Sub TestGetKeyIndex()
    Dim PyDicList As New cPyDicList

    PyDicList.Append "Jonathan Archer", "Enterprise (NX-01 / Start: 2151)"
    PyDicList.Append "Robert April", "USS Enterprise (NCC-1701 / Start: 2245)"
    PyDicList.Append "Christopher Pike", "USS Enterprise (NCC-1701 / Start: 2250)"
    PyDicList.Append "James T. Kirk", "USS Enterprise (NCC-1701 / Start: 2265 / Fünfjahresmission)"
    PyDicList.Append "Willard Decker", "USS Enterprise (NCC-1701 / Start: 2270er / nach Refit)"
    PyDicList.Append "Spock", "USS Enterprise (NCC-1701 / Start: 2280er / als Ausbildungsschiff)"
    PyDicList.Append "James T. Kirk", "USS Enterprise (NCC-1701-A / Start: 2286)"
    PyDicList.Append "John Harriman", "USS Enterprise (NCC-1701-B / Start: 2293)"
    PyDicList.Append "Jean-Luc Picard", "USS Stargazer (NCC-2893 / Start: 2333)"
    PyDicList.Append "Rachel Garrett", "USS Enterprise (NCC-1701-C / Start: 2340er)"
    PyDicList.Append "Lt. Commander Richard Castillo", "USS Enterprise (NCC-1701-C / 2366 / temporär nach Garretts Tod)"
    PyDicList.Append "Jean-Luc Picard", "USS Enterprise (NCC-1701-D / Start: 2364)"
    PyDicList.Append "William T. Riker", "USS Enterprise (NCC-1701-D / 2366 & 2367 / temporär während Picards Borg-Assimilation)"
    PyDicList.Append "Edward Jellico", "USS Enterprise (NCC-1701-D / 2369 / temporär während Picards Geheimmission)"
    PyDicList.Append "Jean-Luc Picard", "USS Enterprise (NCC-1701-E / Start: 2372)"
    PyDicList.Append "Jean-Luc Picard", "USS Verity (NCC-97000 / Start: 2381 / Evakuierungs-Mission)"
    PyDicList.Append "Worf", "USS Enterprise (NCC-1701-E / Start: 2380er / nach Picards Beförderung)"
    PyDicList.Append "Elizabeth Shelby", "USS Enterprise (NCC-1701-F / Start: ca. 2386 / im Kanon bis Frontier Day 2401)"
    PyDicList.Append "Seven of Nine", "USS Enterprise (NCC-1701-G / Start: 2402 / ehemals USS Titan-A)"
    PyDicList.Append "Captain Dax", "USS Enterprise (NCC-1701-J / 26. Jahrhundert / Zeitreise-Begegnung)"
    
    Debug.Print PyDicList.GetKeyIndex("Christopher Pike")
    
    Debug.Print PyDicList.GetKeyIndex("Hardinger Franz")
    
    Debug.Print PyDicList.Imax
    
    Debug.Print PyDicList.Count

End Sub


Sub TestUpdate()
    Dim PyDicList As New cPyDicList
    
    PyDicList.Update "Elizabeth Shelby", "USS Enterprise (NCC-1701-F / Start: ca. 2386 / im Kanon bis Frontier Day 2401)"
    PyDicList.Update "Seven of Nine", "USS Enterprise (NCC-1701-G / Start: 2402 / ehemals USS Titan-A)"
    PyDicList.Update "Captain Dax", "USS Enterprise (NCC-1701-J / 26. Jahrhundert / Zeitreise-Begegnung)"
    PyDicList.Update "Jean-Luc Picard", "USS Stargazer (NCC-2893 / Start: 2333)"
    
    Debug.Print PyDicList("Jean-Luc Picard")
    
    PyDicList.Update "Jean-Luc Picard", "USS Enterprise (NCC-1701-D / Start: 2364)"
    Debug.Print PyDicList("Jean-Luc Picard")
    
End Sub

Sub TestStringListAndKey()
    Dim PyDicList As New cPyDicList
    Dim i As Long

    PyDicList.Update "Jonathan Archer", True
    PyDicList.Update "Robert April", True
    PyDicList.Update "Christopher Pike", True
    PyDicList.Update "James T. Kirk", True
    PyDicList.Update "Willard Decker", True
    PyDicList.Update "Spock", True
    PyDicList.Update "James T. Kirk", True
    PyDicList.Update "John Harriman", True
    PyDicList.Update "Jean-Luc Picard", True
    PyDicList.Update "Rachel Garrett", True
    PyDicList.Update "Lt. Commander Richard Castillo", True
    PyDicList.Update "Jean-Luc Picard", True
    PyDicList.Update "William T. Riker", True
    PyDicList.Update "Edward Jellico", True
    PyDicList.Update "Jean-Luc Picard", True
    PyDicList.Update "Jean-Luc Picard", True
    PyDicList.Update "Worf", True
    PyDicList.Update "Elizabeth Shelby", True
    PyDicList.Update "Seven of Nine", True
    PyDicList.Update "Captain Dax", True
    
    For i = 0 To PyDicList.Imax
        Debug.Print PyDicList.Key(i)
    Next i


End Sub

Sub TestExists()
    Dim PyDicList As New cPyDicList
    Dim i As Long

    PyDicList.Append "Jonathan Archer", "Enterprise (NX-01 / Start: 2151)"
    PyDicList.Append "Robert April", "USS Enterprise (NCC-1701 / Start: 2245)"
    PyDicList.Append "Christopher Pike", "USS Enterprise (NCC-1701 / Start: 2250)"
    PyDicList.Append "James T. Kirk", "USS Enterprise (NCC-1701 / Start: 2265 / Fünfjahresmission)"
    PyDicList.Append "Willard Decker", "USS Enterprise (NCC-1701 / Start: 2270er / nach Refit)"
    PyDicList.Append "Spock", "USS Enterprise (NCC-1701 / Start: 2280er / als Ausbildungsschiff)"
    PyDicList.Append "James T. Kirk", "USS Enterprise (NCC-1701-A / Start: 2286)"
    PyDicList.Append "John Harriman", "USS Enterprise (NCC-1701-B / Start: 2293)"
    PyDicList.Append "Jean-Luc Picard", "USS Stargazer (NCC-2893 / Start: 2333)"
    PyDicList.Append "Rachel Garrett", "USS Enterprise (NCC-1701-C / Start: 2340er)"
    PyDicList.Append "Lt. Commander Richard Castillo", "USS Enterprise (NCC-1701-C / 2366 / temporär nach Garretts Tod)"
    PyDicList.Append "Jean-Luc Picard", "USS Enterprise (NCC-1701-D / Start: 2364)"
    PyDicList.Append "William T. Riker", "USS Enterprise (NCC-1701-D / 2366 & 2367 / temporär während Picards Borg-Assimilation)"
    PyDicList.Append "Edward Jellico", "USS Enterprise (NCC-1701-D / 2369 / temporär während Picards Geheimmission)"
    PyDicList.Append "Jean-Luc Picard", "USS Enterprise (NCC-1701-E / Start: 2372)"
    PyDicList.Append "Jean-Luc Picard", "USS Verity (NCC-97000 / Start: 2381 / Evakuierungs-Mission)"
    PyDicList.Append "Worf", "USS Enterprise (NCC-1701-E / Start: 2380er / nach Picards Beförderung)"
    PyDicList.Append "Elizabeth Shelby", "USS Enterprise (NCC-1701-F / Start: ca. 2386 / im Kanon bis Frontier Day 2401)"
    PyDicList.Append "Seven of Nine", "USS Enterprise (NCC-1701-G / Start: 2402 / ehemals USS Titan-A)"
    PyDicList.Append "Captain Dax", "USS Enterprise (NCC-1701-J / 26. Jahrhundert / Zeitreise-Begegnung)"
    
    Debug.Print PyDicList.Exists("Christopher Pike")
    Debug.Print PyDicList.Exists("Franz Hardinger")
    
    Debug.Print PyDicList.ValueExists("USS Enterprise (NCC-1701 / Start: 2250)")
    Debug.Print PyDicList.ValueExists("USS Hardinger (NCC-1701 / Start: 2250)")
    
    Debug.Print PyDicList.GetDefault("Franz Hardinger", "USS Enterprise (NCC-1701 / Start: 2250)")
    
    Debug.Print PyDicList.Pop("Christopher Pike")
    
    For i = 0 To PyDicList.Imax
        Debug.Print PyDicList.Key(i)
    Next i

End Sub


Sub TestFilter()
    Dim PyDicList As New cPyDicList
    Dim listPicard As cPyDicList
    Dim i As Long

    PyDicList.Append "Jonathan Archer", "Enterprise (NX-01 / Start: 2151)"
    PyDicList.Append "Robert April", "USS Enterprise (NCC-1701 / Start: 2245)"
    PyDicList.Append "Christopher Pike", "USS Enterprise (NCC-1701 / Start: 2250)"
    PyDicList.Append "James T. Kirk", "USS Enterprise (NCC-1701 / Start: 2265 / Fünfjahresmission)"
    PyDicList.Append "Willard Decker", "USS Enterprise (NCC-1701 / Start: 2270er / nach Refit)"
    PyDicList.Append "Spock", "USS Enterprise (NCC-1701 / Start: 2280er / als Ausbildungsschiff)"
    PyDicList.Append "James T. Kirk", "USS Enterprise (NCC-1701-A / Start: 2286)"
    PyDicList.Append "John Harriman", "USS Enterprise (NCC-1701-B / Start: 2293)"
    PyDicList.Append "Jean-Luc Picard", "USS Stargazer (NCC-2893 / Start: 2333)"
    PyDicList.Append "Rachel Garrett", "USS Enterprise (NCC-1701-C / Start: 2340er)"
    PyDicList.Append "Lt. Commander Richard Castillo", "USS Enterprise (NCC-1701-C / 2366 / temporär nach Garretts Tod)"
    PyDicList.Append "Jean-Luc Picard", "USS Enterprise (NCC-1701-D / Start: 2364)"
    PyDicList.Append "William T. Riker", "USS Enterprise (NCC-1701-D / 2366 & 2367 / temporär während Picards Borg-Assimilation)"
    PyDicList.Append "Edward Jellico", "USS Enterprise (NCC-1701-D / 2369 / temporär während Picards Geheimmission)"
    PyDicList.Append "Jean-Luc Picard", "USS Enterprise (NCC-1701-E / Start: 2372)"
    PyDicList.Append "Jean-Luc Picard", "USS Verity (NCC-97000 / Start: 2381 / Evakuierungs-Mission)"
    PyDicList.Append "Worf", "USS Enterprise (NCC-1701-E / Start: 2380er / nach Picards Beförderung)"
    PyDicList.Append "Elizabeth Shelby", "USS Enterprise (NCC-1701-F / Start: ca. 2386 / im Kanon bis Frontier Day 2401)"
    PyDicList.Append "Seven of Nine", "USS Enterprise (NCC-1701-G / Start: 2402 / ehemals USS Titan-A)"
    PyDicList.Append "Captain Dax", "USS Enterprise (NCC-1701-J / 26. Jahrhundert / Zeitreise-Begegnung)"
    
    Set listPicard = PyDicList.FindIndicesByKey("Jean-Luc Picard")
    
    For i = 0 To listPicard.Imax
        Debug.Print PyDicList.ValueIndex(listPicard.ValueIndex(i))
    Next i

End Sub

Sub EinfacherSchleifenTest()
    Dim PyDicList As New cPyDicList
    Dim i As Long
    Dim StartZeit As Double
    Dim EndZeit As Double
    
    ' Die integrierte Excel-Zeitfunktion starten
    StartZeit = Timer
    
    ' Die einfache Schleife mit 100.000 Durchläufen
    For i = 0 To 100000
        PyDicList.Append "Test" & i, i
    Next i
    
    ' Endzeit festhalten
    EndZeit = Timer
    
    ' Ausgabe im Direktfenster (Strg + G)
    Debug.Print "--- TEST ERGEBNIS EXCEL 2007 ---"
    Debug.Print "Durchlaeufe: " & i - 1
    Debug.Print "Dauer: " & Format(EndZeit - StartZeit, "0.000") & " Sekunden"
    Debug.Print "--------------------------------"
End Sub

Sub SchleifenTest()
    Dim PyDicList As New cPyDicList
    Dim i As Long
    Dim StartZeit As Double
    Dim EndZeit As Double
    
    ' Die integrierte Excel-Zeitfunktion starten
    StartZeit = Timer
    
    ' Die einfache Schleife mit 100.000 Durchläufen
    PyDicList.SetNewMaxIndex 1000000
    For i = 0 To 1000000
        PyDicList.UpdateWithIndex i, "Test" & i, "Test" & i
    Next i
    
    ' Endzeit festhalten
    EndZeit = Timer
    
    ' Ausgabe im Direktfenster (Strg + G)
    Debug.Print "--- TEST ERGEBNIS EXCEL 2007 ---"
    Debug.Print "Durchlaeufe: " & i - 1
    Debug.Print "Dauer: " & Format(EndZeit - StartZeit, "0.000") & " Sekunden"
    Debug.Print "--------------------------------"
End Sub


Sub Excel2007LeseTest()
    Dim PyDicList As New cPyDicList
    Dim i As Long, Run As Long
    Dim StartZeit As Double, EndZeit As Double
    
    ' Die 10 unterschiedlichen Positionen zum Auslesen
    Dim TestPositionen As Variant
    TestPositionen = Array(0, 100, 1000, 10000, 50000, 100000, 500000, 750000, 999900, 1000000)
    
    Dim PosIdx As Long
    Dim ZielIndex As Long
    Dim DummyAlsText As String
    
    ' 1. SCHREIBEN (Nutzt das schnelle C-Backend)
    PyDicList.SetNewMaxIndex 1000000
    For i = 0 To 1000000
        PyDicList.Key(i) = "Test" & i
        PyDicList.ValueIndex(i) = "Test" & i
    Next i
    
    Debug.Print "--- EXCEL 2007 LESE-TEST (1 Mio. Aufrufe pro Index) ---"
    Debug.Print "100.000 Eintraege geschrieben."
    Debug.Print "--------------------------------------------------------"
    
    ' 2. LESEN (Jeder Index wird 1.000.000 Mal aufgerufen, um Messzeit zu erzeugen)
    For PosIdx = LBound(TestPositionen) To UBound(TestPositionen)
        ZielIndex = TestPositionen(PosIdx)
        
        StartZeit = Timer
        
        ' Die C-Ebene wird jetzt 1 Million Mal an exakt dieser Stelle abgefragt
        DummyAlsText = PyDicList("Test" & ZielIndex)
        
        EndZeit = Timer
        
        ' Ausgabe im Direktfenster (Strg + G)
        Debug.Print "Index: " & Format(ZielIndex, "000,000") & _
                    " | Gesamtdauer (1 Mio. Lesevorgaenge): " & _
                    Format(EndZeit - StartZeit, "0.000") & " Sekunden"
    Next PosIdx
    
    Debug.Print "--------------------------------------------------------"
End Sub

