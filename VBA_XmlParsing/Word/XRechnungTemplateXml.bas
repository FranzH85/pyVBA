Attribute VB_Name = "XRechnungTemplateXml"
Option Explicit

' ==============================================================================
' MODULE: XRechnungTemplateXml
' DEVELOPER: Franz Hardinger
'
' DESCRIPTION:
' Provides XML string templates for electronic invoicing compliant with the
' XRechnung 3.0 standard. Designed to integrate seamlessly with the pyVBA
' framework for dynamic placeholder replacing and object generation.
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


Function GetXRechnungTemplate() As String
    Dim xml As String
    
    xml = "<?xml version=""1.0"" encoding=""UTF-8""?><invoice:Invoice xmlns:invoice=""urn:oasis:names:specification:ubl:schema:xsd:Invoice-2"" xmlns:cac=""urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"" xmlns:cbc=""urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"">"
    xml = xml & "<cbc:CustomizationID>urn:cen.eu:en16931:2017#compliant#urn:xoev-de:kosit:standard:xrechnung_3.0</cbc:CustomizationID><cbc:ID>[RECHNUNGSNUMMER]</cbc:ID><cbc:IssueDate>[RECHNUNGSDATUM_YYYY-MM-DD]</cbc:IssueDate><cbc:InvoiceTypeCode>[RECHNUNGSTYP_Z_B_380]</cbc:InvoiceTypeCode><cbc:DocumentCurrencyCode>[WÄHRUNGSCODE_Z_B_EUR]</cbc:DocumentCurrencyCode><cbc:BuyerReference>[LEITWEG_ID_KÄUFERREFERENZ]</cbc:BuyerReference>"
    xml = xml & "<cac:OrderReference><cbc:ID>[BESTELLNUMMER]</cbc:ID><cbc:SalesOrderID>[AUFTRAGSNUMMER]</cbc:SalesOrderID></cac:OrderReference>"
    xml = xml & "<cac:AccountingSupplierParty><cac:Party><cac:PartyName><cbc:Name>[VERKÄUFER_NAME_FIRMA]</cbc:Name></cac:PartyName><cac:PostalAddress><cbc:StreetName>[VERKÄUFER_STRASSE_HAUSNUMMER]</cbc:StreetName><cbc:CityName>[VERKÄUFER_ORT]</cbc:CityName><cbc:PostalZone>[VERKÄUFER_PLZ]</cbc:PostalZone><cac:Country><cbc:IdentificationCode>[VERKÄUFER_LAND_ISO_CODE]</cbc:IdentificationCode></cac:Country></cac:PostalAddress></cac:Party></cac:AccountingSupplierParty>"
    xml = xml & "<cac:AccountingCustomerParty><cac:Party><cac:PartyIdentification><cbc:ID>[KÄUFER_KUNDENNUMMER]</cbc:ID></cac:PartyIdentification><cac:PartyName><cbc:Name>[KÄUFER_NAME_FIRMA]</cbc:Name></cac:PartyName><cac:PostalAddress><cbc:StreetName>[KÄUFER_STRASSE_HAUSNUMMER]</cbc:StreetName><cbc:CityName>[KÄUFER_ORT]</cbc:CityName><cbc:PostalZone>[KÄUFER_PLZ]</cbc:PostalZone><cac:Country><cbc:IdentificationCode>[KÄUFER_LAND_ISO_CODE]</cbc:IdentificationCode></cac:Country></cac:PostalAddress></cac:Party></cac:AccountingCustomerParty>"
    xml = xml & "<cac:Delivery><cbc:ActualDeliveryDate>[LIEFERDATUM_LEISTUNGSDATUM_YYYY-MM-DD]</cbc:ActualDeliveryDate></cac:Delivery>"
    xml = xml & "<cac:TaxTotal><cbc:TaxAmount currencyID=""[WÄHRUNGSCODE]"">[STEUERBETRAG_GESAMT]</cbc:TaxAmount><cac:TaxSubtotal><cbc:TaxableAmount currencyID=""[WÄHRUNGSCODE]"">[STEUERBASISBETRAG_NETTO]</cbc:TaxableAmount><cbc:TaxAmount currencyID=""[WÄHRUNGSCODE]"">[STEUERBETRAG_DIESER_SATZ]</cbc:TaxAmount><cac:TaxCategory><cbc:ID>[STEUERKATEGORIE_Z_B_S]</cbc:ID><cbc:Percent>[STEUERSATZ_IN_PROZENT]</cbc:Percent><cac:TaxScheme><cbc:ID>VAT</cbc:ID></cac:TaxScheme></cac:TaxCategory></cac:TaxSubtotal></cac:TaxTotal>"
    xml = xml & "<cac:LegalMonetaryTotal><cbc:LineExtensionAmount currencyID=""[WÄHRUNGSCODE]"">[NETTO_SUMME_ALLER_POSITIONEN]</cbc:LineExtensionAmount><cbc:TaxExclusiveAmount currencyID=""[WÄHRUNGSCODE]"">[RECHNUNGSNETTOBETRAG]</cbc:TaxExclusiveAmount><cbc:TaxInclusiveAmount currencyID=""[WÄHRUNGSCODE]"">[RECHNUNGSBRUTTOBETRAG]</cbc:TaxInclusiveAmount><cbc:PayableAmount currencyID=""[WÄHRUNGSCODE]"">[FÄLLIGER_ZAHLUNGSBETRAG]</cbc:PayableAmount></cac:LegalMonetaryTotal>"
    xml = xml & "</invoice:Invoice>"
    
    GetXRechnungTemplate = xml
End Function

Function GetInvoiceLineTemplate() As String
    Dim xml As String
    
    xml = "<?xml version=""1.0"" encoding=""UTF-8""?><invoice:Invoice xmlns:invoice=""urn:oasis:names:specification:ubl:schema:xsd:Invoice-2"" xmlns:cac=""urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"" xmlns:cbc=""urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"">"
    xml = xml & "<cac:InvoiceLine><cbc:ID>[POSITION_ID]</cbc:ID><cbc:InvoicedQuantity unitCode=""[MENGENEINHEIT_ISO_CODE]"">[MENGE]</cbc:InvoicedQuantity><cbc:LineExtensionAmount currencyID=""[WÄHRUNGSCODE]"">[ZEILENNETTOBETRAG]</cbc:LineExtensionAmount>"
    xml = xml & "<cac:Item><cac:SellersItemIdentification><cbc:ID>[ARTIKELNUMMER_VERKÄUFER]</cbc:ID></cac:SellersItemIdentification><cbc:Name>[ARTIKELNAME]</cbc:Name><cbc:Description>[PRODUKTBESCHREIBUNG]</cbc:Description>"
    xml = xml & "<cac:ClassifiedTaxCategory><cbc:ID>[STEUERKATEGORIE_POSITION]</cbc:ID><cbc:Percent>[STEUERSATZ_POSITION]</cbc:Percent><cac:TaxScheme><cbc:ID>VAT</cbc:ID></cac:TaxScheme></cac:ClassifiedTaxCategory></cac:Item>"
    xml = xml & "<cac:Price><cbc:PriceAmount currencyID=""[WÄHRUNGSCODE]"">[NETTOPREIS_PRO_EINHEIT]</cbc:PriceAmount></cac:Price></cac:InvoiceLine>"
    xml = xml & "</invoice:Invoice>"
    
    GetInvoiceLineTemplate = xml
End Function

Function GetInvoiceLineXML() As cPyDicList
    Dim InvoiceLineXML As New cPyXmltodict
    
    InvoiceLineXML.Parse GetInvoiceLineTemplate
    
    Set GetInvoiceLineXML = InvoiceLineXML.XmlObjekt("invoice:Invoice")("cac:InvoiceLine")
End Function

