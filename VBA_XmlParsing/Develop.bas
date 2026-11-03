Attribute VB_Name = "Develop"
Option Explicit

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

Sub TestArtikel()
    Dim pyXml As New cPyXmltodict
    Dim E1MARAM As cPyDicList
    Dim i As Long
    
    pyXml.ParseFile "I:\GitHub\pyVBA\VBA_XmlParsing\Artikel.xml"
    
    Debug.Print "TABNAM: " & pyXml.XmlObjekt("MATMAS05")("IDOC")("EDI_DC40")("TABNAM")
    Debug.Print "TABNAM SEGMENT: " & pyXml.XmlObjekt("MATMAS05")("IDOC")("EDI_DC40").MyTyp("SEGMENT")
    
    Set E1MARAM = pyXml.XmlObjekt("MATMAS05")("IDOC").getListByKey("E1MARAM")
    
    For i = 0 To E1MARAM.Imax
        Debug.Print E1MARAM.ValueIndex(i)("MATNR") & ": " & E1MARAM.ValueIndex(i)("E1MAKTM")("MAKTX")
    Next i
    
    pyXml.XmlObjekt("MATMAS05")("IDOC")("EDI_DC40")("TABNAM") = "USS"
    pyXml.XmlObjekt("MATMAS05")("IDOC")("EDI_DC40").MyTyp("SEGMENT") = "1701"
    
    pyXml.toXmlFile "I:\GitHub\pyVBA\VBA_XmlParsing\Output.xml"
    
End Sub

Sub TesteRechnung()
    Dim pyXml As New cPyXmltodict
    Dim fd As FileDialog
    Dim filePath As String
    Dim i As Long
    Dim InvoiceLine As cPyDicList
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "XML auswählen"
        .Filters.Clear
        .Filters.Add "XML Dateien", "*.xml"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            filePath = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    pyXml.ParseFile filePath
    pyXml.DebugPrintKeyConst

    Debug.Print "Rechnungs Nr.:" & pyXml.XmlObjekt("invoice:Invoice")("cbc:ID")
    Debug.Print "Datum: " & pyXml.XmlObjekt("invoice:Invoice")("cbc:IssueDate")

    Set InvoiceLine = pyXml.XmlObjekt(xml_invoice_Invoice).getListByKey(xml_cac_InvoiceLine)

    For i = 0 To InvoiceLine.Imax
        Debug.Print "Betrag Pos " & InvoiceLine.ValueIndex(i)("cbc:ID") & ":" & InvoiceLine.ValueIndex(i)("cbc:LineExtensionAmount") & " " & InvoiceLine.ValueIndex(i)("cbc:LineExtensionAmount").MyTyp("currencyID")
    Next i

End Sub



