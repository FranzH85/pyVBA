Attribute VB_Name = "eRechnung"
Option Explicit

' ==============================================================================
' MODULE: eRechnung
' DEVELOPER: Franz Hardinger
'
' DESCRIPTION:
' Utility module providing bidirectional mapping functions between international
' e-invoicing standards (UN/ECE Rec 20 unit codes, EN 16931 tax categories)
' and localized German terms. Essential for parsing and exporting compliant
' XRechnung data within the pyVBA ecosystem.
'
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

Function GetGermanUnit(ByVal InternationalUnit As String) As String
    Select Case UCase$(Trim$(InternationalUnit))
        Case "NAR"
            GetGermanUnit = "Stk."
        Case "HUR"
            GetGermanUnit = "Std."
        Case "DAY"
            GetGermanUnit = "Tage"
        Case "MON"
            GetGermanUnit = "Mon."
        Case "ANN"
            GetGermanUnit = "Jahr"
        Case "PK"
            GetGermanUnit = "Pkt."
        Case "BX"
            GetGermanUnit = "Karton"
        Case "CT"
            GetGermanUnit = "Karton"
        Case "PF"
            GetGermanUnit = "Pal."
        Case "KGM"
            GetGermanUnit = "kg"
        Case "GRM"
            GetGermanUnit = "g"
        Case "MTR"
            GetGermanUnit = "m"
        Case "MTK"
            GetGermanUnit = "m²"
        Case "MTQ"
            GetGermanUnit = "m³"
        Case "LTR"
            GetGermanUnit = "l"
        Case Else
            GetGermanUnit = InternationalUnit
    End Select
End Function

Function GetInternationalUnit(ByVal GermanUnit As String) As String
    Select Case UCase$(Trim$(GermanUnit))
        Case "STK.", "STÜCK", "STK"
            GetInternationalUnit = "NAR"
        Case "STD.", "STUNDEN", "STD"
            GetInternationalUnit = "HUR"
        Case "TAGE", "TAG"
            GetInternationalUnit = "DAY"
        Case "MON.", "MONATE", "MONAT"
            GetInternationalUnit = "MON"
        Case "JAHR", "JAHRE"
            GetInternationalUnit = "ANN"
        Case "PKT.", "PAKET", "PKT"
            GetInternationalUnit = "PK"
        Case "KARTON", "SCHACHTEL"
            GetInternationalUnit = "BX"
        Case "PAL.", "PALETTE", "PAL"
            GetInternationalUnit = "PF"
        Case "KG", "KILOGRAMM"
            GetInternationalUnit = "KGM"
        Case "G", "GRAMM"
            GetInternationalUnit = "GRM"
        Case "M", "METER"
            GetInternationalUnit = "MTR"
        Case "M²", "QUADRATMETER"
            GetInternationalUnit = "MTK"
        Case "M³", "KUBIKMETER"
            GetInternationalUnit = "MTQ"
        Case "L", "LITER"
            GetInternationalUnit = "LTR"
        Case Else
            GetInternationalUnit = GermanUnit
    End Select
End Function

Function GetGermanTaxType(ByVal InternationalTaxCategory As String) As String
    Select Case UCase$(Trim$(InternationalTaxCategory))
        Case "S"
            GetGermanTaxType = "USt."
        Case "Z"
            GetGermanTaxType = "frei"
        Case "E"
            GetGermanTaxType = "befr"
        Case "AE"
            GetGermanTaxType = "RevC"
        Case "K"
            GetGermanTaxType = "Inne"
        Case "G"
            GetGermanTaxType = "Expo"
        Case Else
            GetGermanTaxType = Left$(InternationalTaxCategory, 4)
    End Select
End Function

Function GetInternationalTaxType(ByVal GermanTaxType As String) As String
    Select Case UCase$(Trim$(GermanTaxType))
        Case "UST.", "UMSATZSTEUER", "MWST.", "MEHRWERTSTEUER"
            GetInternationalTaxType = "S"
        Case "STEUERFREI", "0%", "FREI"
            GetInternationalTaxType = "Z"
        Case "STEUERBEFREIT", "BEFREIT", "BEFR"
            GetInternationalTaxType = "E"
        Case "REVERSE CHARGE", "STEUERSCHULDUMKEHR", "§13B", "REVC"
            GetInternationalTaxType = "AE"
        Case "INNERGEMEINSCHAFTLICH", "INNE"
            GetInternationalTaxType = "K"
        Case "EXPORT", "DRITTLAND", "EXPO"
            GetInternationalTaxType = "G"
        Case Else
            GetInternationalTaxType = Left$(GermanTaxType, 4)
    End Select
End Function

