Attribute VB_Name = "Develop"
Option Explicit

Sub TestTime()
    Dim ExcelTable As New cPyExcelTable
    Dim StartZeit As Double
    Dim EndZeit As Double
    Dim MyRange As Range
    
    StartZeit = Timer
    ExcelTable.TableRangeToEnd ThisWorkbook.Sheets("TestTime").Range("A1")
    ExcelTable.TableToRange ThisWorkbook.Sheets("ResultTestTime").Range("A1")
    EndZeit = Timer
    Debug.Print "Dauer: " & Format(EndZeit - StartZeit, "0.000") & " Sekunden"
End Sub

Sub ReadXmlSapData()
    Dim FileDlg As FileDialog
    Dim FilePath As String
    Dim XmlDoc As New cPyXmltodict
    Dim ArtikelDetails As cPyDicList
    Dim ExcelTable As New cPyExcelTable
    Dim I As Long
    Dim Lagerorte As cPyDicList
    Dim I2 As Long
    

    Set FileDlg = Application.FileDialog(msoFileDialogFilePicker)
    With FileDlg
        .Title = "SAP MATMAS XML-Datei auswählen"
        .Filters.Clear
        .Filters.Add "XML Dateien", "*.xml"
        .AllowMultiSelect = False
        
        If .Show <> -1 Then Exit Sub
        FilePath = .SelectedItems(1)
    End With
    
    XmlDoc.ParseFile FilePath
    Set ArtikelDetails = XmlDoc.XmlObjekt("MATMAS05")("IDOC").getListByKey("E1MARAM")
    
    ExcelTable.CreateTable 5, ArtikelDetails.Imax
    
    ExcelTable.ColumnName(0) = "Artikel Nr."
    ExcelTable.ColumnName(1) = "Einheit"
    ExcelTable.ColumnName(2) = "Kategorie"
    ExcelTable.ColumnName(3) = "Bezeichnung"
    ExcelTable.ColumnName(4) = "Lagerort 1"
    ExcelTable.ColumnName(5) = "Lagerort 2"
    
    For I = 0 To ArtikelDetails.Imax
        ExcelTable("Artikel Nr.", I) = ArtikelDetails.ValueIndex(I)("MATNR")
        ExcelTable("Einheit", I) = ArtikelDetails.ValueIndex(I)("MEINS")
        ExcelTable("Kategorie", I) = ArtikelDetails.ValueIndex(I)("MATKL")
        ExcelTable("Bezeichnung", I) = ArtikelDetails.ValueIndex(I)("E1MAKTM")("MAKTX")
        
        Set Lagerorte = ArtikelDetails.ValueIndex(I)("E1MARCM").getListByKey("E1MARDM")
        
        For I2 = 0 To Lagerorte.Imax
            Select Case I2
                Case 0
                    ExcelTable("Lagerort 1", I) = "'" & Lagerorte.ValueIndex(I2)("LGORT")
                Case 1
                    ExcelTable("Lagerort 2", I) = "'" & Lagerorte.ValueIndex(I2)("LGORT")
            End Select
        Next I2
    Next I
    
    ExcelTable.TableToRange ThisWorkbook.Sheets("TestWirteArtikel").Range("A1")
    
    With ExcelTable.MyWorksheet.ListObjects.Add(xlSrcRange, ExcelTable.CurrentTableRange, , xlYes)
        .Name = "TestWirteArtikel"
        .TableStyle = "TableStyleMedium11"
    End With
    
End Sub

Sub YouTube_Demo_Tabelle_Lesen()
    Dim Tbl As New cPyExcelTable
    Dim I As Long
    Dim ZeilenAnzahl As Long
    
    Tbl.TableRangeToEnd ThisWorkbook.Sheets("TestWirteArtikel").Range("A1")
    
    ZeilenAnzahl = Tbl.ImaxRow + 1
    Debug.Print "Erfolgreich eingelesen: " & ZeilenAnzahl & " Sternenflotten-Artikel."
    
    For I = 0 To Tbl.ImaxRow
        If Tbl("Kategorie", I) = "TAKTIK" Then
            Tbl("Artikel Nr.", I) = Tbl("Artikel Nr.", I) & "-REMEDY"
            Debug.Print "Gefunden: " & Tbl("Bezeichnung", I) & " (ID: " & Tbl("Artikel Nr.", I) & ")"
        End If
        
    Next I
    
    Tbl.TableToRange ThisWorkbook.Sheets("ResultTestTime").Range("A1")
    
    MsgBox "Tabelle erfolgreich im Speicher verarbeitet und exportiert!", vbInformation
End Sub

Sub YouTube_Demo_CreateXmlNode()
    Dim XmlString As String
    Dim xml As New cPyXmltodict
    Dim Node As cPyDicList
    
    ' Ein dynamischer XML-String (z.B. frisch aus einer Web-API geladen)
    XmlString = "<WarpAntrieb Typ='Klasse 9'><Status>Bereit</Status><Energie>100</Energie></WarpAntrieb>"
    
    ' DEINE NEUE FUNKTION IM EINSATZ: Parsing und Zuweisung in einem Rutsch!
    xml.Parse XmlString
    xml.XmlObjekt("WarpAntrieb").Append "Data", xml.CreateXmlNode("<maxWarp>9.999</maxWarp>")
    
    xml.XmlObjekt("WarpAntrieb").Append "CurrentWarp", xml.CreateXmlNode("7")
    
    
    ' Zugriff mit deiner eleganten Python-Syntax:
    Debug.Print "Antriebs-Status: " & xml.XmlObjekt("WarpAntrieb")("Status")  ' Ausgabe: Bereit
    Debug.Print "Energie-Level:  " & xml.XmlObjekt("WarpAntrieb")("Energie") ' Ausgabe: 100
    Debug.Print "maxWarp:  " & xml.XmlObjekt("WarpAntrieb")("Data")("maxWarp") ' Ausgabe: 9.999
    Debug.Print "CurrentWarp:  " & xml.XmlObjekt("WarpAntrieb")("CurrentWarp") ' Ausgabe: 7
    
    ' Und sogar Attribute holst du dir direkt ab:
    Debug.Print "Spezifikation:  " & xml.XmlObjekt("WarpAntrieb").MyTyp("Typ") ' Ausgabe: Klasse 9
End Sub

