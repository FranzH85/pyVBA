
🚀 pyVBA – Native JSON Parser for VBA (Python-Style)

This project brings an elegant, fully native JSON parsing experience in true Python dictionary and list style directly into your VBA projects!
Thanks to the combination of the clsJson parser and the flexible cPyDicList data container, you can read and write deeply nested JSON objects without unreadable code. The solution works completely platform-independent out of the box on Windows and macOS (no external DLL installations or references required).

📺 YouTube Video

In the accompanying video, I show the tool in live action, explaining the recursive functionality and how to read and write nested JSON structures in VBA using a complex Star Trek scenario!
👉 Watch the full video here:
https://www.youtube.com/watch?v=v8NZg7K8Q8Y

🛠️ Features

• True Bracket Chaining: Query nested JSON paths easily via JsonObjekt("Key")("SubKey") – completely without the annoying .Value syntax!
• Automatic Typing: The parser automatically detects objects, arrays, strings, numbers, booleans, and null values.
• Array Handling like in Python: When reading and extending JSON arrays, no key is required – access is handled cleanly via the index ValueIndex(i).
• True UTF-8 Saving: The built-in .Save method automatically strips out the annoying Byte Order Marks (BOM), ensuring a clean UTF-8 file without data clutter lands on your hard drive (includes native Mac support).

⚠️ Important Note on Importing!

To activate the elegant bracket queries without needing .Value, you cannot simply copy and paste the raw code text!
You must import the two class files into your project using the VBA Editor via File ➔ Import File..., otherwise the hidden VBA attributes (VB_UserMemId = 0) will be lost.

💾 Files Included in the Repository

• clsJson.cls – The core class for recursive analysis and exporting JSON structures.
• cPyDicList.cls – The underlying data structure combining dictionaries and lists.
• Sternenflotte.xlsm – The ready-to-use workbook including the full test code for immediate testing.

📄 License

This project is licensed under the MIT License – free to use for both private and commercial purposes.
“Software is structure, not syntax, so that Picard knows what Kirk did, even though Kirk has been dead for a long time.” 🚀