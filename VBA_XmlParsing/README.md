# pyVBA – Python-Style XML Parsing & Data Structures in Native VBA

A lightweight, blazing-fast, and zero-dependency VBA framework that brings the elegant syntax of Python dictionaries and dynamic XML parsing directly into Visual Basic for Applications (VBA).

📺 **Watch the full tutorial on YouTube:** [VBA XML parsen wie in PYTHON! 🐍 (xmltodict für Excel & Word)](https://youtu.be/DABjl62x0bY)

**Inspired by Python's popular `xmltodict` library**, this framework allows you to work with XML files as if they were native hierarchical dictionaries. Unlike traditional, nested MSXML traversal methods or the native `Scripting.Dictionary`, **pyVBA** preserves element sequences, handles complex XML attributes dynamically, and allows intuitive bracket-chaining. It operates natively on both Windows and macOS without external library installations.

---

## 🚀 Key Features Demonstrated in the Video

* **Pythonic `xmltodict` Syntax:** Just like in Python, access deeply nested data trees instantly using standard parentheses: `pyXml.XmlObjekt("MATMAS05")("IDOC")("EDI_DC40")("TABNAM")`.
* **Dynamic XML Attribute Access:** Query or overwrite node metadata and attributes (e.g., `<TaxAmount currencyID="EUR">` or `SEGMENT="1"`) seamlessly via the dedicated `.MyTyp` property slot.
* **Smart Sequence Arrays (`getListByKey`):** Extract and manipulate repeating XML segments (such as SAP IDoc segments or XRechnung invoice line items) into iterable collection lists while maintaining strict element ordering.
* **Instant IntelliSense Autocompletion:** Utilize the `DebugPrintKeyConst` utility to automatically generate fully typed VBA string constants from any parsed XML file—converting XML colons (`invoice:Invoice`) into safe VBA underscores (`xml_invoice_Invoice`) for full autocomplete support.
* **Template-Driven XML Exporting:** Dynamically swap out placeholders within structured XML templates via code and export fully valid UTF-8 strings into files using raw binary streams—eliminating trailing BOM issues.

---

## 📦 Core Component Classes

1. **`cPyDicList`**: The main data container mimicking Python's dictionary/list hybrids. Features advanced index manipulation, sequential safe-guarding, and comfort methods like `GetDefault`, `Pop`, and `getListByKey`.
2. **`cPyXmltodict`**: The XML parsing engine that wraps native OS document object models into dynamic object structures with seamless file input/output workflows (`Parse`, `ParseFile`, `toXmlFile`). **This class replicates the core behavior of Python's `xmltodict.parse()`**.
3. **`cPySimpleObject`**: A universal variant vehicle utilizing a default member ID pointer (`VB_UserMemId = 0`) to handle primitive data types and deep sub-object assignments uniformly within bracket sequences.

---

## 🛠️ Critical Installation Instructions

### ⚠️ Importing Class Files Correctly
To achieve true Python bracket-chaining without appending `.Value` explicitly on every call, properties within `cPyDicList` and `cPySimpleObject` implement hidden attributes. **Because the MS Office VBA Editor masks these flags, you must import the native `.cls` files directly into your project:**

1. Clone or download the `.cls` files from this repository.
2. In the VBA Editor (`Alt + F11`), go to **File** -> **Import File...** (`Ctrl + M`).
3. Import `cPyDicList.cls`, `cPyXmltodict.cls`, and `cPySimpleObject.cls`.

### Environment Reference (Windows Only)
To enable fast native parsing, open the VBA Editor, navigate to **Tools** -> **References**, and check:
* **Microsoft XML, v6.0** (MSXML2)

*(Note: The ADODB stream utilized for clean UTF-8 exporting is implemented via Late-Binding within the source code and does not require an explicit reference in the menu!)*

---

## 💻 Real-World Implementation (Word Ribbon Integration)

As shown in the second half of the channel video, the framework maps standard-compliant electronic invoice fields (**XRechnung 3.0 / ZUGFeRD / UBL Schema**) straight onto Microsoft Word Document structures using custom wrappers like `cPyBookmark` and `cPyWordTable`.

### Safeguarding Object References (The Pointer Rule)
When appending templates in a loop to add rows or line items dynamically, **pyVBA** strictly adheres to Python's instantiation rules. Always append the clean class template to the container object *first*, and then query `.LastValue` to fill fields:

```vba
' ❌ WRONG: Will overwrite all prior elements due to sharing the same pointer reference
For i = 1 To InvoiceItems
    Set MyLineItem = GetInvoiceLineXML
    MyLineItem("cbc:ID") = i
    eRechnungDetails.Append "cac:InvoiceLine", MyLineItem
Next i

'  CORRECT: Safely appends an independent instance every time
For i = 1 To InvoiceItems
    eRechnungDetails.Append "cac:InvoiceLine", GetInvoiceLineXML
    eRechnungDetails.LastValue("cbc:ID") = i
    eRechnungDetails.LastValue("cac:Item")("cbc:Name") = "Warp Drive System Class 9"
Next i
```

---

## 🖖 Star Trek Rationale
> *"Software is structure, not syntax – so that Picard knows what Kirk did, even though Kirk has been dead for a long time."*

---

## 📄 License

This framework is licensed under the **MIT License**. Feel free to use, fork, alter, and embed these classes into private or commercial macro projects.
