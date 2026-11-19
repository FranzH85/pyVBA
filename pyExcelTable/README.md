# pyVBA – Python-Style Excel Tables & Data Structures in Native VBA

A lightweight, blazing-fast, and zero-dependency VBA framework that brings the elegant table syntax of **Python Pandas** directly into Excel VBA.

📺 **Watch the full tutorial on YouTube:** [Excel VBA im PYTHON PANDAS-Style! 🐼 (10x schneller als normale Schleifen)](https://www.youtube.com/watch?v=tUhp5EoQ-DY)

**Inspired by Python's popular `pandas` library**, this class allows you to load entire Excel tables into memory (RAM) all at once, manipulate them using readable column names, and write them back in a single batch. Say goodbye to slow cell-by-cell loops – **cPyExcelTable** processes tens of thousands of cells in about one second and runs natively without external libraries.

---

## 🚀 Key Features & Highlights Demonstrated in the Video

* **High-Speed In-Memory Processing:** Reads entire table ranges via Variant arrays into RAM in one go (`TableRange` / `TableRangeToEnd`) and writes them back just as fast (`TableToRange`) – e.g., 50,000 cells in approx. 1.18 seconds.
* **Intuitive Pandas Syntax:** Access cells directly via column name and row index (0-based) without having to append `.Value`: `Tbl("Kategorie", I)` or `Tbl("Artikel Nr.", I) = "12345"`.
* **Dynamic In-Memory Table Manipulation:** Insert new columns at any specific position (`AddColumn`), delete unneeded columns (`DeleteColumn`), or mark rows for removal (`DeleteRow`) before outputting the table.
* **Direct Excel Range Linking:** Immediately retrieve the entire table range after writing (`CurrentTableRange`) for automatic formatting (`ListObjects`), or jump directly to specific cells (`CellRange`, `NameCellRange`).
* **Seamless XML Integration (`cPyXmltodict`):** Transform complex SAP IDoc XML data (`MATMAS`) into structured, formatted Excel tables in a flash – including the new helper function `CreateXmlNode` for dynamic XML nodes.

---

## 📦 Core Component Classes

1. **`cPyExcelTable`**: The main class for high-speed table processing in memory. It replicates the behavior of a Python DataFrame in VBA (`TableRangeToEnd`, `CreateTable`, `TableToRange`, `AddColumn`, `DeleteRow`).
2. **`cPyDicList`**: The underlying data container that combines Python's dictionaries and lists, enabling fast column and row management.

---

## 🛠️ Critical Installation Instructions

### ⚠️ Importing Class Files Correctly!
To use true Python syntax without constantly appending `.Value` (e.g., `Tbl("Column", 0)` instead of `Tbl.Value("Column", 0)`), the properties in `cPyExcelTable` and `cPyDicList` implement hidden VBA attributes (`Attribute Value.VB_UserMemId = 0`). **Because the MS Office VBA Editor strips these attributes when copying text, you must import the `.cls` files directly:**

1. Download the `.cls` files from this repository.
2. Open the VBA Editor in Excel (`Alt + F11`) and go to **File** -> **Import File...** (`Ctrl + M`).
3. Import `cPyExcelTable.cls`, `cPyDicList.cls`, and `cPySimpleObject.cls`. *(For the XML examples, also import `cPyXmltodict.cls`)*.

### Required References
* For standalone use of **`cPyExcelTable`**, **no** additional references are needed – the class runs completely natively.
* If you also use the XML combination with **`cPyXmltodict`**, go to **Tools** -> **References** and check **Microsoft XML, v6.0**.

---

## 💻 Real-World Examples from the Video

### 1. Load an Existing Excel Table into RAM, Filter & Write Back
```vba
Sub YouTube_Demo_Tabelle_Lesen()
    Dim Tbl As New cPyExcelTable
    Dim I As Long
    Dim ZeilenAnzahl As Long
    
    ' Load the entire table starting at A1 into memory all at once
    Tbl.TableRangeToEnd ThisWorkbook.Sheets("TestWirteArtikel").Range("A1")
    
    ZeilenAnzahl = Tbl.ImaxRow + 1
    Debug.Print "Erfolgreich eingelesen: " & ZeilenAnzahl & " Artikel."
    
    ' Filter and modify values in RAM at lightning speed (Python Pandas-Style)
    For I = 0 To Tbl.ImaxRow
        If Tbl("Kategorie", I) = "TAKTIK" Then
            Tbl("Artikel Nr.", I) = Tbl("Artikel Nr.", I) & "-REMEDY"
        End If
    Next I
    
    ' Write the processed table to the target sheet in a single batch
    Tbl.TableToRange ThisWorkbook.Sheets("ResultTestTime").Range("A1")
End Sub
```

### 2. Create a New Table from Scratch & Format It
```vba
Sub Neue_Tabelle_Erstellen()
    Dim ExcelTable As New cPyExcelTable
    
    ' Initialize table with 3 columns (index 0-2) and 100 rows (index 0-99)
    ExcelTable.CreateTable 2, 99
    
    ' Define column names
    ExcelTable.ColumnName(0) = "Artikel Nr."
    ExcelTable.ColumnName(1) = "Kategorie"
    ExcelTable.ColumnName(2) = "Bezeichnung"
    
    ' Assign values directly in RAM
    ExcelTable("Artikel Nr.", 0) = "NCC-1701"
    ExcelTable("Kategorie", 0) = "ANTRIEB"
    ExcelTable("Bezeichnung", 0) = "Dilithium-Kristall"
    
    ' Write table to the sheet and format it directly as a smart Excel table
    ExcelTable.TableToRange ThisWorkbook.Sheets("Tabelle1").Range("A1")
    ExcelTable.MyWorksheet.ListObjects.Add(xlSrcRange, ExcelTable.CurrentTableRange, , xlYes).TableStyle = "TableStyleMedium11"
End Sub
```

---

## 🖖 Star Trek Rationale
> *"Software is structure, not syntax – so that Picard knows what Kirk did, even though Kirk has been dead for a long time."*

---

## 📄 License

This framework is licensed under the **MIT License**. Feel free to use, fork, alter, and embed these classes into private or commercial macro projects.
