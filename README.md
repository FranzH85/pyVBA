# pyVBA – Python-Style XML Parsing, Excel Tables & Data Structures in Native VBA

A lightweight, blazing-fast, and zero-dependency VBA framework that brings the elegant syntax of Python dictionaries, `xmltodict` parsing, `pandas`-style Excel table manipulation, and dynamic data structures directly into Visual Basic for Applications (VBA).

---

## 📂 Repository Structure

The architecture of this repository is designed to separate **universal core VBA modules** from **application-specific extensions** and real-world orchestration examples:

### ⚙️ Core Framework (Root Directory)
The component files located directly in the root directory are **universal** and can be imported into any VBA host application (Excel, Word, Access, PowerPoint, etc.):
* **`cPyDicList.cls`**: The main data container mimicking Python's hybrid dictionary/list structures. Supports advanced indexing, sequence preservation, and safe-lookup comfort methods.
* **`cPyXmltodict.cls`**: The highly optimized XML parsing engine replicating the exact behavior of Python's `xmltodict` library using native OS document object models.
* **`cPySimpleObject.cls`**: A universal variant vessel utilizing a default member ID pointer (`VB_UserMemId = 0`) to handle primitive data and complex references uniformly.
* **`clsJson.cls`**: The universal core JSON processing module.

### 🔌 Application-Specific Extensions
These subfolders contain specialized classes tailored strictly for the object models of their respective Microsoft Office host applications:
* 📂 **`Word/`**: Contains `cPyWordTable.cls` and `cPyBookmark.cls` to map parsed XML/JSON data structures natively onto Word text grids and document bookmarks via inline markup.
* 📂 **`Excel/`**: Contains `cPyExcelTable.cls` as well as tools, worksheet handlers, and data structures tailored exclusively for high-speed Excel automation.
* 📂 **`Access/`**: Contains engine extensions and database-handling classes designed explicitly for Microsoft Access.

### 🔬 Practical Examples & Showcases
These directories contain implementation examples, test procedures, and specific showcase files featured in video tutorials:
* 📂 **`pyExcelTable/`**: The Python Pandas-style Excel table showcase (`cPyExcelTable.cls`), demonstrating high-speed in-memory table processing (50,000 cells in ~1.18 seconds) and SAP XML-to-Excel workflows.
* 📂 **`VBA_XmlParsing/`**: The complete parsing showcase (SAP IDoc files, XRechnung 3.0 e-invoices, and dynamic template rendering).
* 📂 **`VBA_JsonParsing/`**: Comprehensive code samples and test cases for working with structured JSON data payloads.
* 📂 **`Python-dic-list-stile-in-vba/`**: The benchmark test field showcasing performance scripts handling over 1,000,000 dictionary elements in under 0.5 seconds.

---

## 🛠️ Critical Installation Instructions

### ⚠️ Importing Class Files Correctly
To achieve true Python bracket-chaining without appending `.Value` explicitly on every call, the properties within the core data classes implement hidden default member attributes (`VB_UserMemId = 0`). **Because the MS Office VBA Editor masks these flags, you must import the native `.cls` files using the editor's import interface:**

1. Clone or download this repository to your local drive.
2. Open your VBA Editor (`Alt + F11`) in any Office host application.
3. Go to **File** -> **Import File...** (`Ctrl + M`) and import the required core files directly.

### Environment Reference (Windows Only)
To enable native high-performance XML parsing, open the VBA Editor, navigate to **Tools** -> **References**, and check:
* **Microsoft XML, v6.0** (MSXML2)

*(Note: The ADODB binary stream engine utilized for clean UTF-8 file exporting uses Late-Binding within the source code and does not require an explicit library reference in the menu!)*

---

## 🖖 Star Trek Rationale
> *"Software is structure, not syntax – so that Picard knows what Kirk did, even though Kirk has been dead for a long time."*

---

## 📄 License

This framework is licensed under the **MIT License**. You are completely free to fork, modify, share, and embed these components into private or commercial macro projects.
