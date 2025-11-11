# Endnote Citation Extractor

Extract endnotes, hyperlinks, page references, and context sentences from Microsoft Word documents for fact-checking and citation analysis.

## Features

- **Extracts endnote numbers and full citation text** from Word .docx files
- **Captures context sentences** - the sentence immediately before each citation reference
- **Extracts hyperlinks** using 4 different methods to handle various Word formatting:
  - `<w:hyperlink>` tags with relationship resolution
  - Simple field codes (`HYPERLINK` instructions)
  - Complex field codes (`<w:instrText>`)
  - Plain text URLs pasted directly into endnotes
- **Parses page numbers** from citations (supports "p." and "pp." formats)
- **Exports to UTF-8 CSV** for analysis in Excel, R, or other tools

## Installation

### Required R Packages

```r
# Install required packages
install.packages(c(
  "xml2",      # XML parsing
  "dplyr",     # Data manipulation
  "tidyr",     # Data tidying
  "purrr",     # Functional programming
  "stringr",   # String operations
  "readr",     # CSV export
  "here",      # Path management
  "tools"      # File utilities
))
```

### Download the Code

Clone or download this repository:
```bash
git clone https://github.com/marisam99/endnote_extractor_MM.git
```

## Usage

### Running the Extractor

1. Open `extractor_main.R` in RStudio or your R environment
2. Source the entire script (Ctrl/Cmd + Shift + S in RStudio, or run all lines)
3. A file chooser dialog will appear - select your `.docx` file
4. The script will process the document and automatically save the results

### Output Location

The CSV file will be saved in the **same directory as your input file** with the prefix `extracted_`:

**Example:**
- Input: `/Documents/Projects/report.docx`
- Output: `/Documents/Projects/extracted_report.csv`

## Output Format

The exported CSV contains 5 columns:

| Column | Description | Example |
|--------|-------------|---------|
| `number` | Endnote number | `1`, `2`, `3` |
| `context_sentence` | Sentence immediately before the citation | `"This finding was later confirmed."` |
| `source_link` | Extracted URL/hyperlink from the endnote | `https://example.com/article` |
| `page_ref` | Page numbers referenced | `p.12`, `pp.45-50`, `p.12, pp.45-50` |
| `full_endnote` | Complete endnote text | `"Smith, J. (2020). Title. Journal, 10(2), 45-50."` |

**Note:** If an endnote contains multiple citations with different URLs, each will appear on a separate row with the same endnote number.

## How It Works

1. **Unzips the .docx file** - Word documents are ZIP archives containing XML files
2. **Parses XML structure:**
   - `word/document.xml` - Main document content and endnote references
   - `word/endnotes.xml` - Endnote definitions and text
   - `word/_rels/endnotes.xml.rels` - Relationship mappings for hyperlinks
3. **Extracts context** - Finds the sentence immediately before each `<w:endnoteReference>` tag
4. **Extracts hyperlinks** - Uses 4 different extraction methods to capture URLs regardless of how they're formatted
5. **Parses page numbers** - Intelligently splits multi-citation endnotes and extracts page references using regex
6. **Exports to CSV** - Combines all data into a structured table with UTF-8 encoding

## Requirements

- **R version:** 4.0 or higher recommended
- **Input format:** Microsoft Word `.docx` files only (not `.doc`)
- **Page number format:** Must use "p." or "pp." prefix for extraction to work
  - ✅ Works: `p. 12`, `pp. 45-50`, `p. S12`, `pp. iv-x`
  - ❌ Won't extract: `page 12`, `12-15`, random numbers without prefix

## Known Limitations

- **Endnotes only** - Does not process footnotes (different XML structure)
- **Page number detection** - Requires "p." or "pp." prefix; won't extract page numbers without this delimiter
- **Multi-citation endnotes** - Page numbers are matched to URLs within the same endnote; matching may be imperfect for complex citations
- **Context extraction** - Captures the last sentence before the citation marker; may not work perfectly with complex paragraph structures

## Project Structure

```
endnote_extractor_MM/
├── extractor_main.R       # Main script (current version with all features)
├── helpers.R              # Helper functions for hyperlinks and page parsing
├── extractor_hyperlinks.R # Standalone version (older, no modular helpers)
├── extractor_original.R   # Biko's original version (basic functionality)
└── README.md              # This file
```

## Version History

- **Current (`extractor_main.R`)**: Full-featured with hyperlink extraction, page parsing, and modular architecture
- **`extractor_hyperlinks.R`**: Added hyperlink extraction (all functions inline)
- **`extractor_original.R`**: Original version by Biko - basic endnote and context extraction

## Acknowledgments

- Original script by Biko
- Enhanced and refactored by Marisa Mission
