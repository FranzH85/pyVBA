Attribute VB_Name = "Develop"
Option Explicit


Sub TestNeuesJsonInEinerZeile()
    Dim JsonParser As New clsJson
    Dim JsonText As String
    Dim i As Long
    Dim t As Boolean
    
    ' Der komplette JSON-String in einer einzigen, sauberen Zeile abgelegt
    JsonText = "{""foederation"":""Vereinigte Foeoderation der Planeten"",""sternzeit_aktuell"":47634.4,""alarm_stufe_rot"":true,""diplomatischer_status"":null,""flaggschiff"":{""name"":""USS Enterprise"",""registrierung"":""NCC-1701-D"",""klasse"":""Galaxy-Klasse"",""status"":""Aktiv"",""besatzung"":1014,""antrieb"":{""typ"":""Warp-Antrieb"",""kern"":""M/AM-Reaktor (Materie/Antimaterie)"",""geschwindigkeit_max"":9.6,""taktischer_status"":""Online""},""fuehrung"":{""captain"":""Jean-Luc Picard"",""erster_offizier"":""William T. Riker"",""chefingenieur"":""Geordi La Forge""}},""aktive_flotte"":[{""name"":""USS Defiant"",""registrierung"":""NX-74205"",""antrieb_typ"":""Quanten-Warp"",""tarnvorrichtung"":true},{""name"":""USS Voyager"",""registrierung"":""NCC-74656"",""antrieb_typ"":""Klassischer Warp"",""tarnvorrichtung"":false}]}"
               
    ' Den JSON-String parsen
    JsonParser.Parse JsonText
    
    ' Daten direkt im Python-Stil ausgeben
    Debug.Print "=== BASIS DATEN ==="
    Debug.Print "Föderation:      " & JsonParser.JsonObjekt("foederation")
    Debug.Print "Alarm Stufe Rot: " & JsonParser.JsonObjekt("alarm_stufe_rot")
    Debug.Print "M/AM-Geschw.:    " & JsonParser.JsonObjekt("flaggschiff")("antrieb")("geschwindigkeit_max")
    Debug.Print ""
    
    
    Debug.Print "=== BASIS DATEN ==="
    Debug.Print "klasse:      " & JsonParser.JsonObjekt("flaggschiff")("klasse")
    Debug.Print "besatzung: " & JsonParser.JsonObjekt("flaggschiff")("besatzung")
    Debug.Print "M/AM-Geschw.:    " & JsonParser.JsonObjekt("flaggschiff")("antrieb")("geschwindigkeit_max")
    Debug.Print ""
    
    
    JsonParser.JsonObjekt("aktive_flotte").Append "", New cPyDicList
    JsonParser.JsonObjekt("aktive_flotte").LastValue.MyTyp = JsonParser.setObjectString
    JsonParser.JsonObjekt("aktive_flotte").LastValue.Append "name", "USS Prometheus"
    JsonParser.JsonObjekt("aktive_flotte").LastValue.Append "registrierung", "NX-59650"
    JsonParser.JsonObjekt("aktive_flotte").LastValue.Append "antrieb_typ", "3a Warp-Himmelsführungssystem"
    JsonParser.JsonObjekt("aktive_flotte").LastValue.Append "tarnvorrichtung", False
    
    
    ' Die For-Schleife für die aktive Flotte
    Debug.Print "=== AKTIVE FLOTTE ==="
    For i = 0 To JsonParser.JsonObjekt("aktive_flotte").Imax
        Debug.Print "Schiff " & (i + 1) & ":"
        Debug.Print "  Name:     " & JsonParser.JsonObjekt("aktive_flotte").ValueIndex(i)("name")
        Debug.Print "  Reg.-Nr.: " & JsonParser.JsonObjekt("aktive_flotte").ValueIndex(i)("registrierung")
        Debug.Print "  Tarnung:  " & JsonParser.JsonObjekt("aktive_flotte").ValueIndex(i)("tarnvorrichtung")
        Debug.Print "----------------------------------------"
    Next i
    Debug.Print JsonParser.ToJson
    
    JsonParser.Save "I:\Dokumente\Meine Softzwareprojekte\YouTube\Vorlagen\Excel\Klassen\Test.json"
End Sub
