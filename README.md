# DocxEndnoteExtractor

A lightweight R script to extract endnote text **and** the immediately preceding sentence (context) from Microsoft Word (`.docx`) documents.

---

## 🔍 Features

- **Unzip once**: Extracts both `document.xml` and `endnotes.xml` from a `.docx` ZIP container.
- **Context capture**: Finds every `<w:endnoteReference>` in the main document, then pulls the last complete sentence before each marker.
- **Endnote text**: Parses the endnote bodies, converts IDs to lowercase Roman numerals.
- **Clean CSV**: Outputs a UTF-8, Excel-friendly CSV with BOM so curly quotes, “§” symbols, etc., render correctly.
  
---

## 🚀 Quickstart

1. **Install R dependencies**  
   ```r
   install.packages(c("xml2", "dplyr", "tibble", "purrr", "stringr", "readr"))
   ```

---

2. **Source the script**
```r
source("endnote_with_context_extractor.R")
```

---

3. **Run on your docx**
```r
docx_path <- "~/Downloads/MyReport.docx"
tbl <- extract_endnotes_with_context(docx_path)
write_excel_csv(tbl, "~/Desktop/endnotes_with_context.csv")
```

---

4. **Open in Excel** — your CSV will open with proper UTF-8 encoding.

---

📦 Files
extract_endnotes_with_context.R — main function

README.md — this description
