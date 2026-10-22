# cPyDicList – Python-Style Dictionary & List for VBA (Excel)

[![License: MIT](https://shields.io)](https://opensource.org)
[![Excel Version](https://shields.io)]()
[![Platform](https://shields.io)]()

An extremely fast, flexible, and lightweight data container for VBA that brings the elegant **Python dictionary syntax directly into Microsoft Excel**. The best part: It works completely out of the box without installing external libraries and runs natively on both **Windows and macOS**.

---

## 🚀 Key Features

* **True Python Syntax:** Allows nested bracket chaining like `PyDicList("Ship")("Name")` without the need for tedious `.Value` calls.
* **Cross-Platform:** Unlike the classic `Scripting.Dictionary`, this class runs natively on Mac as well.
* **Extended Capabilities:** Includes powerful Python-inspired methods such as `.GetDefault()`, `.Pop()`, and `.Update()`.
* **Preserves Insertion Order:** Unlike standard dictionaries, the sequence (index) of added elements is strictly maintained.
* **High-Performance:** Features an optimized memory pre-allocation hack (`.SetNewMaxIndex`) for lightning-fast write operations when dealing with over 1 million records.

---

## 📺 YouTube Video & Tutorial

Watch the full video walkthrough and explanation for this class on YouTube:
👉 **[https://youtu.be/tH-zB4UXN4Y]**

---

## ⚠️ IMPORTANT INSTALLATION NOTE

To make the magic bracket chaining work without requiring `.Value`, this class utilizes a hidden VBA attribute (`VB_UserMemId = 0`) that sets the `Property Get Value` as the default function.

**Do not simply copy and paste the raw code text into the VBA editor!** 
Instead, follow these steps:
1. Download the `cPyDicList.cls` file from this repository.
2. Open your VBA Editor (Alt + F11).
3. Click on **File -> Import File...** (Ctrl + M) in the menu and select the downloaded `.cls` file.

---

## 💻 Code Examples (Star Trek Edition)

### 1. Nested Dictionaries (Bracket Chaining)
```vba
Sub StarTrekExample()
    Dim PyDicList As New cPyDicList
    
    ' Create a nested structure
    PyDicList.Append "Ship", New cPyDicList
    PyDicList.LastValue.Append "Name", "Enterprise"
    PyDicList.LastValue.Append "Class", "Constitution"
    
    ' Read values Python-style
    Debug.Print PyDicList("Ship")("Name")   ' Output: Enterprise
    
    ' Overwrite values Python-style
    PyDicList("Ship")("Name") = "Enterprise-D"
    Debug.Print PyDicList("Ship")("Name")   ' Output: Enterprise-D
End Sub
```

### 2. Eliminating Duplicates with `.Update`
The `.Update` method overwrites existing keys or creates them if they don't exist. Perfect for filtering lists:
```vba
Sub FilterExample()
    Dim PyDicList As New cPyDicList
    
    ' Jean-Luc Picard is added multiple times, overwriting the previous key state
    PyDicList.Update "Jean-Luc Picard", "USS Stargazer"
    PyDicList.Update "Jean-Luc Picard", "USS Enterprise-D"
    
    Debug.Print PyDicList("Jean-Luc Picard") ' Output: USS Enterprise-D (the latest state)
End Sub
```

### 3. Preventing Crashes with `.GetDefault` and `.Pop`
```vba
Sub HelperExample()
    Dim PyDicList As New cPyDicList
    PyDicList.Append "Captain", "Picard"
    
    ' If a key does not exist, it safely returns the fallback value instead of crashing
    Debug.Print PyDicList.GetDefault("First Officer", "Riker") ' Output: Riker
    
    ' Retrieve a value and remove it from the list at the same time
    Debug.Print PyDicList.Pop("Captain") ' Output: Picard
    Debug.Print PyDicList.Exists("Captain") ' Output: False
End Sub
```

---

## 🛠️ Performance Mode for Big Data
When processing more than 100,000 entries inside a loop, you can pre-allocate memory beforehand to bypass the slow native `ReDim Preserve` behavior:

```vba
Sub BigDataTest()
    Dim PyDicList As New cPyDicList
    Dim i As Long
    
    ' Pre-allocate 1 million entries in memory
    PyDicList.SetNewMaxIndex 1000000
    
    For i = 0 To 1000000
        PyDicList.UpdateWithIndex i, "Key" & i, "Value" & i
    Next i
End Sub
```

---

## 📄 License

This project is licensed under the **MIT License** – you are free to use this code in both private and commercial projects. Full details are included within the class file headers.

Developed with ❤️ by **Franz Hardinger**
