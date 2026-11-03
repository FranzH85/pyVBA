# pyVBA – Word Extension (cPyWordTable & cPyBookmark)

This subdirectory contains the specialized Microsoft Word components of the **pyVBA** framework. While the core classes (`cPyXmltodict` and `cPyDicList`) handle the sequential and structurally sound parsing of XML data, these two components bridge that data directly into your active Microsoft Word documents.

They allow you to interact with Document Bookmarks and Tables via an intuitive pythonic matrix/dictionary syntax, providing native support for dynamic runtime text formatting.

---

## 🚀 Key Features

* **Matrix-Style Table Access (`cPyWordTable`):** Eliminates cumbersome Word table cell selection structures. Address any table cell instantly using a clean 2D matrix interface: `PyWordTable(Row, Column) = "Value"`.
* **Dictionary-Style Bookmarks (`cPyBookmark`):** Manage your document bookmarks via intuitive key-value pairs just like a Python dictionary: `PyBookmark("InvoiceNo") = "INV-1701"`.
* **Inline Markup Text Formatting:** Pass layout properties straight inside text strings. The classes automatically parse and apply the corresponding font properties at runtime:
  * `\bold` -> Toggles text weight to **Bold**
  * `\itali` -> Toggles text style to *Italic*
  * `\size:7` -> Adjusts the font size dynamically (e.g., to 7pt)
* **Automated Bookmark Preservation:** Overwriting native bookmarks via standard VBA typically deletes the bookmark object from the document scope. `cPyBookmark` automatically redefines the target range in the background, keeping your bookmarks fully intact for successive read/write cycles.
* **Automated Constant Generation:** The `DebugPrintBookmarkConst` utility scans all active bookmarks in the document and outputs ready-to-copy VBA code constants into your Immediate Window.

---

## 📦 Component Classes in this Directory

1. **`cPyBookmark`**: The primary wrapper for text bookmark management. Encapsulates the `Document.Bookmarks` collection and exposes the key-value dictionary interface alongside the string formatting engine.
2. **`cPyWordTable`**: The dedicated wrapper for Microsoft Word Tables. Provides convenience operations for bulk row allocation (`AddRows`, `DeleteRows`), cell merging (`MergeCells`), and precise layout borders (`LineTop`, `LineBottom`, `DoubleLineBottom`).

---

## 🛠️ Critical Setup Instructions

### ⚠️ Importing Class Files Correctly
Just like the core framework elements, both `cPyWordTable` and `cPyBookmark` implement hidden default member attributes (`VB_UserMemId = 0`) to achieve true bracket-chaining without requiring an explicit `.Value` suffix. **You must import the native `.cls` files using the editor import menu:**

1. Open your VBA Editor inside Microsoft Word (`Alt + F11`).
2. Go to **File** -> **Import File...** (`Ctrl + M`).
3. Select and import `cPyWordTable.cls` and `cPyBookmark.cls`.

### Environment Reference
Because these classes explicitly interact with the Word Object Model, ensure that the **Microsoft Word Object Library** is checked under **Tools** -> **References** (this is active by default when running macros within Microsoft Word).

---

## 💻 Practical Implementation Example (Ribbon Callback)

The following example shows how gracefully the Word extensions integrate with the core framework components to parse a standard-compliant **XRechnung 3.0** XML payload straight into a formatted document structure:

```vba
' 1. Fill document bookmarks using pythonic syntax and force bold weight
Dim PyBookmark As New cPyBookmark
PyBookmark("CustomerName") = "\(\bold\)" & eRechnungDetails("cac:AccountingCustomerParty")("cac:Party")("cac:PartyName")("cbc:Name")

' 2. Initialize the matrix table wrapper
Dim PyWordTable As New cPyWordTable
PyWordTable.Init ThisDocument.Tables(1)

' 3. Flush old rows except the header
PyWordTable.DeleteRows 2

' 4. Isolate repeating XML line items and allocate new rows
Dim eRechnungPosition As cPyDicList
Set eRechnungPosition = eRechnungDetails.getListByKey("cac:InvoiceLine")
PyWordTable.AddRows eRechnungPosition.Imax + 1

' 5. Write tree leaf nodes straight onto the Word Table grid matrix
Dim i As Long
For i = 0 To eRechnungPosition.Imax
    PyWordTable(i + 2, 1) = eRechnungPosition.ValueIndex(i)("cbc:ID")
    PyWordTable(i + 2, 3) = "\(\bold\)" & eRechnungPosition.ValueIndex(i)("cac:Item")("cbc:Name")
    PyWordTable.AddText i + 2, 3, Chr(13) & eRechnungPosition.ValueIndex(i)("cac:Item")("cbc:Description")
    PyWordTable(i + 2, 7) = Format\$(Val(eRechnungPosition.ValueIndex(i)("cbc:LineExtensionAmount")), "0.00")
Next i
```
