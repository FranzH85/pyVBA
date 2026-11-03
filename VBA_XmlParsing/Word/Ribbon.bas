Attribute VB_Name = "Ribbon"
Option Explicit


' ==============================================================================
' MODULE: Ribbon
' DEVELOPER: Franz Hardinger
'
' DESCRIPTION:
' Implements the Microsoft Office Ribbon UI callbacks for loading and saving
' electronic invoices (XRechnung) natively in Word using the pyVBA framework.
' Acts as an integration bridge between 'cPyWordTable', 'cPyBookmark', and
' 'cPyXmltodict' to seamlessly map dynamic XML hierarchies onto document elements.
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

Const Bo_bold As String = "\bold"
Const Bo_itali As String = "\itali"
Const Bo_size As String = "\size:"
Const Bo_AuftragsNr As String = "AuftragsNr"
Const Bo_Bearbeiter As String = "Bearbeiter"
Const Bo_BearbeiterE_Mail As String = "BearbeiterE_Mail"
Const Bo_BearbeiterTelefon As String = "BearbeiterTelefon"
Const Bo_BestellNr As String = "BestellNr"
Const Bo_Kommission As String = "Kommission"
Const Bo_KundenName As String = "KundenName"
Const Bo_KundenNr As String = "KundenNr"
Const Bo_KundenOrt As String = "KundenOrt"
Const Bo_KundenStraße As String = "KundenStraße"
Const Bo_Leistungsdatum As String = "Leistungsdatum"
Const Bo_Lieferbedingung As String = "Lieferbedingung"
Const Bo_PreisEinh1 As String = "PreisEinh1"
Const Bo_PreisEinh2 As String = "PreisEinh2"
Const Bo_Rechnungsdatum As String = "Rechnungsdatum"
Const Bo_RechnungsNr As String = "RechnungsNr"
Const Bo_Zahlbar As String = "Zahlbar"

Const VBA_ElementDelete As String = "VBA_ElementDelete"
Const VBA_TextMarker   As String = "#text"
Const xml_xml As String = "xml"
Const xml_invoice_Invoice As String = "invoice:Invoice"
Const xml_cbc_CustomizationID As String = "cbc:CustomizationID"
Const xml_cbc_ID As String = "cbc:ID"
Const xml_cbc_IssueDate As String = "cbc:IssueDate"
Const xml_cbc_InvoiceTypeCode As String = "cbc:InvoiceTypeCode"
Const xml_cbc_DocumentCurrencyCode As String = "cbc:DocumentCurrencyCode"
Const xml_cbc_BuyerReference As String = "cbc:BuyerReference"
Const xml_cac_OrderReference As String = "cac:OrderReference"
Const xml_cbc_SalesOrderID As String = "cbc:SalesOrderID"
Const xml_cac_AccountingSupplierParty As String = "cac:AccountingSupplierParty"
Const xml_cac_Party As String = "cac:Party"
Const xml_cac_PartyName As String = "cac:PartyName"
Const xml_cbc_Name As String = "cbc:Name"
Const xml_cac_PostalAddress As String = "cac:PostalAddress"
Const xml_cbc_StreetName As String = "cbc:StreetName"
Const xml_cbc_CityName As String = "cbc:CityName"
Const xml_cbc_PostalZone As String = "cbc:PostalZone"
Const xml_cac_Country As String = "cac:Country"
Const xml_cbc_IdentificationCode As String = "cbc:IdentificationCode"
Const xml_cac_AccountingCustomerParty As String = "cac:AccountingCustomerParty"
Const xml_cac_PartyIdentification As String = "cac:PartyIdentification"
Const xml_cac_Delivery As String = "cac:Delivery"
Const xml_cbc_ActualDeliveryDate As String = "cbc:ActualDeliveryDate"
Const xml_cac_TaxTotal As String = "cac:TaxTotal"
Const xml_cbc_TaxAmount As String = "cbc:TaxAmount"
Const xml_cac_TaxSubtotal As String = "cac:TaxSubtotal"
Const xml_cbc_TaxableAmount As String = "cbc:TaxableAmount"
Const xml_cac_TaxCategory As String = "cac:TaxCategory"
Const xml_cbc_Percent As String = "cbc:Percent"
Const xml_cac_TaxScheme As String = "cac:TaxScheme"
Const xml_cac_LegalMonetaryTotal As String = "cac:LegalMonetaryTotal"
Const xml_cbc_LineExtensionAmount As String = "cbc:LineExtensionAmount"
Const xml_cbc_TaxExclusiveAmount As String = "cbc:TaxExclusiveAmount"
Const xml_cbc_TaxInclusiveAmount As String = "cbc:TaxInclusiveAmount"
Const xml_cbc_PayableAmount As String = "cbc:PayableAmount"
Const xml_cac_InvoiceLine As String = "cac:InvoiceLine"
Const xml_cbc_InvoicedQuantity As String = "cbc:InvoicedQuantity"
Const xml_cac_Item As String = "cac:Item"
Const xml_cac_SellersItemIdentification As String = "cac:SellersItemIdentification"
Const xml_cbc_Description As String = "cbc:Description"
Const xml_cac_ClassifiedTaxCategory As String = "cac:ClassifiedTaxCategory"
Const xml_cac_Price As String = "cac:Price"
Const xml_cbc_PriceAmount As String = "cbc:PriceAmount"


'Callback for btnERechnungLaden onAction
Sub ERechnungLaden_Click(control As IRibbonControl)
    Dim PyWordTable As New cPyWordTable
    Dim PyBookmark As New cPyBookmark
    Dim eRechnung As New cPyXmltodict
    Dim eRechnungDetails As cPyDicList
    Dim eRechnungPosition As cPyDicList
    Dim fd As fileDialog
    Dim filePath As String
    Dim InvoiceItems As Long
    Dim i As Long
    
    
    Set fd = Application.fileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "eRechnung XML auswählen"
        .Filters.Clear
        .Filters.Add "XML Dateien", "*.xml"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            filePath = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    eRechnung.ParseFile filePath
    Set eRechnungDetails = eRechnung.XmlObjekt(xml_invoice_Invoice)
    eRechnung.DebugPrintKeyConst
    
    PyBookmark(Bo_RechnungsNr) = Bo_bold & eRechnungDetails(xml_cbc_ID)

    PyBookmark(Bo_Rechnungsdatum) = eRechnungDetails(xml_cbc_IssueDate)
    PyBookmark(Bo_Leistungsdatum) = eRechnungDetails(xml_cac_Delivery)(xml_cbc_ActualDeliveryDate)

    PyBookmark(Bo_Kommission) = eRechnungDetails(xml_cbc_BuyerReference)
    PyBookmark(Bo_BestellNr) = eRechnungDetails(xml_cac_OrderReference)(xml_cbc_ID)
    PyBookmark(Bo_AuftragsNr) = eRechnungDetails(xml_cac_OrderReference)(xml_cbc_SalesOrderID)

    PyBookmark(Bo_PreisEinh1) = Bo_bold & eRechnungDetails(xml_cac_LegalMonetaryTotal)(xml_cbc_PayableAmount).MyTyp("currencyID")
    PyBookmark(Bo_PreisEinh2) = Bo_bold & eRechnungDetails(xml_cac_LegalMonetaryTotal)(xml_cbc_PayableAmount).MyTyp("currencyID")

    PyBookmark(Bo_KundenName) = Bo_bold & eRechnungDetails(xml_cac_AccountingCustomerParty)(xml_cac_Party)(xml_cac_PartyName)(xml_cbc_Name)
    PyBookmark(Bo_KundenStraße) = eRechnungDetails(xml_cac_AccountingCustomerParty)(xml_cac_Party)(xml_cac_PostalAddress)(xml_cbc_StreetName)
    PyBookmark(Bo_KundenOrt) = eRechnungDetails(xml_cac_AccountingCustomerParty)(xml_cac_Party)(xml_cac_PostalAddress)(xml_cbc_PostalZone) & " " & eRechnungDetails(xml_cac_AccountingCustomerParty)(xml_cac_Party)(xml_cac_PostalAddress)(xml_cbc_CityName)
    PyBookmark(Bo_KundenNr) = eRechnungDetails(xml_cac_AccountingCustomerParty)(xml_cac_Party)(xml_cac_PartyIdentification)(xml_cbc_ID)

    PyBookmark(Bo_Bearbeiter) = eRechnungDetails(xml_cac_AccountingSupplierParty)(xml_cac_Party)(xml_cac_PartyName)(xml_cbc_Name)

    
    PyWordTable.Init ThisDocument.Tables(1)
    PyWordTable.DeleteRows 2
    
    Set eRechnungPosition = eRechnungDetails.getListByKey(xml_cac_InvoiceLine)

    InvoiceItems = eRechnungPosition.Imax + 1
    PyWordTable.AddRows InvoiceItems + 4

    For i = 3 To (InvoiceItems + 5)
        PyWordTable.ClearBorders i
    Next i

    PyWordTable.MergeCells InvoiceItems + 3, 5, 6
    PyWordTable.LineBottom InvoiceItems + 3, 5
    PyWordTable.LineBottom InvoiceItems + 3, 6

    PyWordTable.MergeCells InvoiceItems + 4, 5, 6
    PyWordTable.DoubleLineBottom InvoiceItems + 4, 5
    PyWordTable.DoubleLineBottom InvoiceItems + 4, 6

    PyWordTable.MergeCells InvoiceItems + 5, 5, 6
    PyWordTable.LineBottom InvoiceItems + 5, 5
    PyWordTable.LineBottom InvoiceItems + 5, 6
    
    For i = 0 To eRechnungPosition.Imax
        PyWordTable(i + 2, 1) = eRechnungPosition.ValueIndex(i)(xml_cbc_ID)
        PyWordTable(i + 2, 2) = eRechnungPosition.ValueIndex(i)(xml_cac_Item)(xml_cac_SellersItemIdentification)(xml_cbc_ID)
        PyWordTable(i + 2, 3) = Bo_bold & eRechnungPosition.ValueIndex(i)(xml_cac_Item)(xml_cbc_Name)
        PyWordTable.AddText i + 2, 3, Chr(13) & eRechnungPosition.ValueIndex(i)(xml_cac_Item)(xml_cbc_Description)
        PyWordTable(i + 2, 4) = Format$(Val(eRechnungPosition.ValueIndex(i)(xml_cbc_InvoicedQuantity)), "0.00")
        PyWordTable(i + 2, 5) = GetGermanUnit(eRechnungPosition.ValueIndex(i)(xml_cbc_InvoicedQuantity).MyTyp("unitCode"))
        PyWordTable(i + 2, 6) = Format$(Val(eRechnungPosition.ValueIndex(i)(xml_cac_Price)(xml_cbc_PriceAmount)), "0.00")
        PyWordTable(i + 2, 7) = Format$(Val(eRechnungPosition.ValueIndex(i)(xml_cbc_LineExtensionAmount)), "0.00")

    Next i
    
    PyWordTable(InvoiceItems + 3, 5) = "Summe Netto"
    PyWordTable(InvoiceItems + 3, 6) = eRechnungDetails(xml_cac_TaxTotal)(xml_cac_TaxSubtotal)(xml_cbc_TaxableAmount).MyTyp("currencyID") & " " & _
                                        Format$(Val(eRechnungDetails(xml_cac_TaxTotal)(xml_cac_TaxSubtotal)(xml_cbc_TaxableAmount)), "0.00")
                                        
    PyWordTable(InvoiceItems + 4, 5) = Bo_size & "7 " & eRechnungDetails(xml_cac_TaxTotal)(xml_cac_TaxSubtotal)(xml_cac_TaxCategory)(xml_cbc_Percent) & "% " & _
                                        GetGermanTaxType(eRechnungDetails(xml_cac_TaxTotal)(xml_cac_TaxSubtotal)(xml_cac_TaxCategory)(xml_cbc_ID)) & " auf " & _
                                        Format$(Val(eRechnungDetails(xml_cac_TaxTotal)(xml_cac_TaxSubtotal)(xml_cbc_TaxableAmount)), "0.00") & " " & _
                                        eRechnungDetails(xml_cac_TaxTotal)(xml_cac_TaxSubtotal)(xml_cbc_TaxableAmount).MyTyp("currencyID")
                                        
    PyWordTable(InvoiceItems + 4, 6) = eRechnungDetails(xml_cac_TaxTotal)(xml_cac_TaxSubtotal)(xml_cbc_TaxAmount).MyTyp("currencyID") & " " & _
                                        Format$(Val(eRechnungDetails(xml_cac_TaxTotal)(xml_cac_TaxSubtotal)(xml_cbc_TaxAmount)), "0.00")
                                        
    PyWordTable(InvoiceItems + 5, 5) = Bo_bold & "Endsumme"
    
    PyWordTable(InvoiceItems + 6, 6) = Bo_bold & eRechnungDetails(xml_cac_LegalMonetaryTotal)(xml_cbc_PayableAmount).MyTyp("currencyID") & " " & _
                                        Format$(Val(eRechnungDetails(xml_cac_LegalMonetaryTotal)(xml_cbc_PayableAmount)), "0.00")
                                        
 
End Sub

'Callback for btnERechnungSpeichern onAction
Sub ERechnungSpeichern_Click(control As IRibbonControl)
   Dim PyWordTable As New cPyWordTable
    Dim PyBookmark As New cPyBookmark
    Dim eRechnung As New cPyXmltodict
    Dim eRechnungDetails As cPyDicList
    Dim KundenOrtSplit() As String
    Dim HelpString As String
    Dim SummeNettoSplit() As String
    Dim i As Long
    Dim InvoiceItems As Long
    Dim PercentSplit() As String
    Dim InternationalTaxType() As String
    Dim TaxableAmountSplit() As String
    Dim TaxAmountSplit() As String
    Dim MonetaryTotal() As String
    Dim ItemNameSplit() As String
    Dim fileDialog As fileDialog
    Dim selectedPath As String

    
    eRechnung.Parse GetXRechnungTemplate
    
    Set eRechnungDetails = eRechnung.XmlObjekt(xml_invoice_Invoice)
    
    eRechnungDetails(xml_cbc_ID) = PyBookmark(Bo_RechnungsNr)
    
    eRechnungDetails(xml_cbc_IssueDate) = PyBookmark(Bo_Rechnungsdatum)
    eRechnungDetails(xml_cac_Delivery)(xml_cbc_ActualDeliveryDate) = PyBookmark(Bo_Leistungsdatum)
    
    eRechnungDetails(xml_cbc_InvoiceTypeCode) = "380"
    eRechnungDetails(xml_cbc_DocumentCurrencyCode) = PyBookmark(Bo_PreisEinh1)
    
    eRechnungDetails(xml_cac_AccountingSupplierParty)(xml_cac_Party)(xml_cac_PartyName)(xml_cbc_Name) = PyBookmark(Bo_Bearbeiter)
    eRechnungDetails(xml_cac_AccountingSupplierParty)(xml_cac_Party)(xml_cac_PostalAddress)(xml_cbc_StreetName) = "Am Orbitallift 17 (Hangar-Sektor Delta)"
    eRechnungDetails(xml_cac_AccountingSupplierParty)(xml_cac_Party)(xml_cac_PostalAddress)(xml_cbc_CityName) = "MARS"
    eRechnungDetails(xml_cac_AccountingSupplierParty)(xml_cac_Party)(xml_cac_PostalAddress)(xml_cbc_PostalZone) = "00001"
    eRechnungDetails(xml_cac_AccountingSupplierParty)(xml_cac_Party)(xml_cac_PostalAddress)(xml_cac_Country)(xml_cbc_IdentificationCode) = "SLO"
    

    eRechnungDetails(xml_cbc_BuyerReference) = PyBookmark(Bo_Kommission)
    eRechnungDetails(xml_cac_OrderReference)(xml_cbc_ID) = PyBookmark(Bo_BestellNr)
    eRechnungDetails(xml_cac_OrderReference)(xml_cbc_SalesOrderID) = PyBookmark(Bo_AuftragsNr)

    eRechnungDetails(xml_cac_LegalMonetaryTotal)(xml_cbc_PayableAmount).MyTyp("currencyID") = PyBookmark(Bo_PreisEinh1)

    eRechnungDetails(xml_cac_AccountingCustomerParty)(xml_cac_Party)(xml_cac_PartyName)(xml_cbc_Name) = PyBookmark(Bo_KundenName)
    eRechnungDetails(xml_cac_AccountingCustomerParty)(xml_cac_Party)(xml_cac_PostalAddress)(xml_cbc_StreetName) = PyBookmark(Bo_KundenStraße)
    eRechnungDetails(xml_cac_AccountingCustomerParty)(xml_cac_Party)(xml_cac_PostalAddress)(xml_cac_Country)(xml_cbc_IdentificationCode) = "SOL"
    
    KundenOrtSplit = Split(PyBookmark(Bo_KundenOrt), " ")
    eRechnungDetails(xml_cac_AccountingCustomerParty)(xml_cac_Party)(xml_cac_PostalAddress)(xml_cbc_PostalZone) = KundenOrtSplit(0)
    eRechnungDetails(xml_cac_AccountingCustomerParty)(xml_cac_Party)(xml_cac_PostalAddress)(xml_cbc_CityName) = Trim$(Replace(PyBookmark(Bo_KundenOrt), KundenOrtSplit(0), ""))
    
    eRechnungDetails(xml_cac_AccountingCustomerParty)(xml_cac_Party)(xml_cac_PartyIdentification)(xml_cbc_ID) = PyBookmark(Bo_KundenNr)

    PyWordTable.Init ThisDocument.Tables(1)
    
    InvoiceItems = PyWordTable.Imax - 5

    SummeNettoSplit() = Split(PyWordTable(InvoiceItems + 3, 6), " ")

    eRechnungDetails(xml_cac_TaxTotal)(xml_cac_TaxSubtotal)(xml_cbc_TaxableAmount).MyTyp("currencyID") = SummeNettoSplit(0)
    eRechnungDetails(xml_cac_LegalMonetaryTotal)(xml_cbc_LineExtensionAmount).MyTyp("currencyID") = SummeNettoSplit(0)
    eRechnungDetails(xml_cac_LegalMonetaryTotal)(xml_cbc_TaxExclusiveAmount).MyTyp("currencyID") = SummeNettoSplit(0)
    
    eRechnungDetails(xml_cac_TaxTotal)(xml_cac_TaxSubtotal)(xml_cbc_TaxableAmount) = Val(SummeNettoSplit(1))
    eRechnungDetails(xml_cac_LegalMonetaryTotal)(xml_cbc_LineExtensionAmount) = Val(SummeNettoSplit(1))
    eRechnungDetails(xml_cac_LegalMonetaryTotal)(xml_cbc_TaxExclusiveAmount) = Val(SummeNettoSplit(1))
    
    
    PercentSplit() = Split(PyWordTable(InvoiceItems + 4, 5), "% ")
    eRechnungDetails(xml_cac_TaxTotal)(xml_cac_TaxSubtotal)(xml_cac_TaxCategory)(xml_cbc_Percent) = Val(PercentSplit(0))
    
    InternationalTaxType() = Split(PercentSplit(1), " auf ")
    eRechnungDetails(xml_cac_TaxTotal)(xml_cac_TaxSubtotal)(xml_cac_TaxCategory)(xml_cbc_ID) = GetInternationalTaxType(InternationalTaxType(0))
    eRechnungDetails(xml_cac_TaxTotal)(xml_cac_TaxSubtotal)(xml_cac_TaxCategory)(xml_cac_TaxScheme)(xml_cbc_ID) = "VAT"
    
    TaxableAmountSplit() = Split(InternationalTaxType(1), " ")
    eRechnungDetails(xml_cac_TaxTotal)(xml_cac_TaxSubtotal)(xml_cbc_TaxableAmount) = Val(TaxableAmountSplit(0))
    eRechnungDetails(xml_cac_TaxTotal)(xml_cac_TaxSubtotal)(xml_cbc_TaxableAmount).MyTyp("currencyID") = TaxableAmountSplit(1)
    
    
    TaxAmountSplit() = Split(PyWordTable(InvoiceItems + 4, 6), " ")
    eRechnungDetails(xml_cac_TaxTotal)(xml_cac_TaxSubtotal)(xml_cbc_TaxAmount).MyTyp("currencyID") = TaxAmountSplit(0)
    eRechnungDetails(xml_cac_TaxTotal)(xml_cbc_TaxAmount).MyTyp("currencyID") = TaxAmountSplit(0)
    eRechnungDetails(xml_cac_TaxTotal)(xml_cbc_TaxAmount) = Val(TaxAmountSplit(1))
    eRechnungDetails(xml_cac_TaxTotal)(xml_cac_TaxSubtotal)(xml_cbc_TaxAmount) = Val(TaxAmountSplit(1))
    

    MonetaryTotal() = Split(PyWordTable(InvoiceItems + 6, 6), " ")
    eRechnungDetails(xml_cac_LegalMonetaryTotal)(xml_cbc_PayableAmount).MyTyp("currencyID") = MonetaryTotal(0)
    eRechnungDetails(xml_cac_LegalMonetaryTotal)(xml_cbc_PayableAmount) = Val(MonetaryTotal(1))
    
    eRechnungDetails(xml_cac_LegalMonetaryTotal)(xml_cbc_TaxInclusiveAmount).MyTyp("currencyID") = MonetaryTotal(0)
    eRechnungDetails(xml_cac_LegalMonetaryTotal)(xml_cbc_TaxInclusiveAmount) = Val(MonetaryTotal(1))

    For i = 1 To InvoiceItems
        eRechnungDetails.Append xml_cac_InvoiceLine, GetInvoiceLineXML
        
        eRechnungDetails.LastValue(xml_cbc_ID) = PyWordTable(i + 1, 1)
        eRechnungDetails.LastValue(xml_cac_Item)(xml_cac_SellersItemIdentification)(xml_cbc_ID) = PyWordTable(i + 1, 2)
        
        ItemNameSplit() = Split(PyWordTable(i + 1, 3), Chr(13))
        eRechnungDetails.LastValue(xml_cac_Item)(xml_cbc_Name) = ItemNameSplit(0)
        eRechnungDetails.LastValue(xml_cac_Item)(xml_cbc_Description) = ItemNameSplit(1)
        eRechnungDetails.LastValue(xml_cac_Item)(xml_cac_ClassifiedTaxCategory)(xml_cbc_ID) = GetInternationalTaxType(InternationalTaxType(0))
        eRechnungDetails.LastValue(xml_cac_Item)(xml_cac_ClassifiedTaxCategory)(xml_cbc_Percent) = Val(PercentSplit(0))
        eRechnungDetails.LastValue(xml_cbc_InvoicedQuantity) = Val(PyWordTable(i + 1, 4))
        eRechnungDetails.LastValue(xml_cbc_InvoicedQuantity).MyTyp("unitCode") = GetInternationalUnit(PyWordTable(i + 1, 5))
    
        eRechnungDetails.LastValue(xml_cac_Price)(xml_cbc_PriceAmount).MyTyp("currencyID") = PyBookmark(Bo_PreisEinh1)
        eRechnungDetails.LastValue(xml_cac_Price)(xml_cbc_PriceAmount) = Val(PyWordTable(i + 1, 6))
        
        eRechnungDetails.LastValue(xml_cbc_LineExtensionAmount).MyTyp("currencyID") = PyBookmark(Bo_PreisEinh2)
        eRechnungDetails.LastValue(xml_cbc_LineExtensionAmount) = Val(PyWordTable(i + 1, 7))
        
    Next i
    

                                        
    Set fileDialog = Application.fileDialog(msoFileDialogSaveAs)
    
    With fileDialog
        .Title = "Save XRechnung As..."
        .InitialFileName = PyBookmark(Bo_RechnungsNr) & ".xml"
        
        For i = 1 To .Filters.Count
            If InStr(1, .Filters.Item(i).Extensions, "xml", vbTextCompare) > 0 Then
                .FilterIndex = i
                Exit For
            End If
        Next i
        
        If .Show = -1 Then
            selectedPath = .SelectedItems(1)
            
            If Right$(LCase$(selectedPath), 4) <> ".xml" Then
                selectedPath = selectedPath & ".xml"
            End If
            
            eRechnung.toXmlFile selectedPath
        Else
            MsgBox "Save operation cancelled.", vbInformation, "Cancelled"
        End If
    End With
             
End Sub
