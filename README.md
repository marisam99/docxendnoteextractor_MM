# DocxEndnoteExtractor

A lightweight R script to extract endnote text **and** the immediately preceding sentence (context) from Microsoft Word (`.docx`) documents.

---

## 🔍 Features

- **Unzip once**: Extracts both `document.xml` and `endnotes.xml` from a `.docx` ZIP container.
- **Context capture**: Finds every `<w:endnoteReference>` in the main document, then pulls the last complete sentence before each marker.
- **Endnote text**: Parses the endnote bodies, converts IDs to lowercase Roman numerals.
- **Clean CSV**: Outputs a UTF-8, Excel-friendly CSV with BOM so curly quotes, “§” symbols, etc., render correctly.
- **Re-indexed**: Renumbers endnotes consecutively (1,2,3…) regardless of Word’s internal IDs.
