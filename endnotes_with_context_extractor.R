# --------------------------------------------------------------------------
# Combined Endnote & Context Extractor for Word (.docx)
#  – Unzips a .docx once
#  – Reads word/document.xml to pull the sentence immediately before each
#    <w:endnoteReference>
#  – Reads word/endnotes.xml to pull the endnote text and its Roman numeral
#  – Returns a tibble: number | roman | endnote | context_sentence
# --------------------------------------------------------------------------

# 0. CHANGE PATH TO YOUR WORD FILE ----------------------------------------------------------------
# Point this at your .docx:
docx_path <- "/path/to/your/document.docx" # Adjust this path

# 1. Packages ----------------------------------------------------------------
library(xml2)    # parse & query XML
library(dplyr)   # mutate, arrange, native pipe |>
library(tibble)  # tibble()
library(purrr)   # map_chr()
library(stringr) # str_squish(), str_split()
library(utils)   # unzip(), unlink(), as.roman()
library(readr)   # write_excel_csv() for UTF-8 CSV

# 2. Function definition ----------------------------------------------------
extract_endnotes_with_context <- function(docx_path) {
  
  # a) Unzip .docx into a temp folder ---------------------------------------
  temp_dir <- tempfile("docx_unzip_")
  dir.create(temp_dir)
  unzip(zipfile = docx_path, exdir = temp_dir)
  
  # b) Paths to the two XMLs ------------------------------------------------
  doc_xml      <- file.path(temp_dir, "word", "document.xml")
  endnotes_xml <- file.path(temp_dir, "word", "endnotes.xml")
  
  # c) Read and namespace ---------------------------------------------------
  doc_main  <- read_xml(doc_xml)
  ns_main   <- xml_ns(doc_main)
  
  doc_notes <- read_xml(endnotes_xml)
  ns_notes  <- xml_ns(doc_notes)
  
  # d) Extract context sentences --------------------------------------------
  refs <- xml_find_all(doc_main, ".//w:endnoteReference", ns_main)
  
  context_df <- tibble(
    number = xml_attr(refs, "w:id", ns = ns_main) |> as.integer()
  ) |>
    mutate(
      context_sentence = map_chr(refs, function(ref_node) {
        run_node <- xml_parent(ref_node)
        if (xml_name(run_node) != "r") {
          alt <- xml_find_first(ref_node, "preceding::w:r[1]", ns_main)
          if (!is.null(alt)) run_node <- alt
        }
        para <- xml_find_first(ref_node, "ancestor::w:p", ns_main)
        runs_before <- xml_find_all(run_node, "preceding-sibling::w:r//w:t", ns_main)
        if (length(runs_before) == 0 && !is.null(para)) {
          runs_before <- xml_find_all(para, ".//w:t", ns_main)
        }
        text_before <- runs_before |>
          xml_text() |>
          paste(collapse = "") |>
          str_squish()
        if (text_before == "") return("")
        sentences <- str_split(text_before,
                               "(?<=[\\.\\!\\?])\\s+",
                               simplify = FALSE)[[1]]
        if (length(sentences) == 0) return(text_before)
        last_sent <- tail(sentences, 1)
        if (str_squish(last_sent) == "") text_before else last_sent
      })
    )
  
  # e) Extract endnotes ------------------------------------------------------
  notes <- xml_find_all(
    doc_notes,
    ".//w:endnote[not(@w:type) and number(@w:id) > 0]",
    ns_notes
  )
  
  ids <- xml_attr(notes, "w:id", ns = ns_notes) |> as.integer()
  
  get_note_text <- function(node) {
    xml_find_all(node, ".//w:t", ns_notes) |>
      xml_text() |>
      paste(collapse = "") |>
      str_squish()
  }
  
  endnotes_df <- tibble(
    number  = ids,
    endnote = vapply(notes, get_note_text, character(1))
  )
  
  # f) Join, re-index from 1, compute roman, sort ----------------------------
  combined <- full_join(endnotes_df, context_df, by = "number") |>
    arrange(number) |>
    mutate(
      number = row_number(),
      roman  = tolower(as.character(as.roman(number)))
    ) |>
    select(number, roman, endnote, context_sentence)
  
  # g) Cleanup ---------------------------------------------------------------
  unlink(temp_dir, recursive = TRUE)
  
  # h) Return ---------------------------------------------------------------
  combined
}

# 3. Run extractor and Save Output ----------------------------------------------------------
# Run the combined extractor:
endnotes_with_context <- extract_endnotes_with_context(docx_path)

# Write out to CSV
write_excel_csv(
  endnotes_with_context,
  file = "~/Desktop/endnotes_with_context.csv"
)
