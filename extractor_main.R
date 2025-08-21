# --------------------------------------------------------------------------
# Combined Endnote, Source, & Context Extractor for Word (.docx)
#  – Unzips a .docx once
#  – Reads word/document.xml to pull the sentence immediately before each
#    <w:endnoteReference>
#  – Reads word/endnotes.xml to pull the endnote text and its Roman numeral
#  - Reads word/rels.xml to pull source links out of the endnote text
#  - Parses the endnote for page numbers for each source link
#  – Returns a tibble: number | context_sentence | source_link | page_ref | full_endnote
# --------------------------------------------------------------------------

# 0. SETUP: CHANGE THIS AS NEEDED ----------------------------------------------
output_file <- "AIFP_measurement_fact_check.csv"
# docx_path <- "C:/Users/Marisa Mission/OneDrive - BW/2025 AIFP Year-Long AI Funder Collaborative/WIP/Measurement/2025-08-05 Measurement draft.docx"

# 1. Load packages ------------------------------------------------------------------
library(xml2)    # parse & query XML
library(tidyr)   # unnest()
library(dplyr)   # mutate, native pipe |>
library(tibble)  # tibble()
library(purrr)   # map_chr()
library(stringr) # str_squish(), str_split()
library(utils)   # unzip(), unlink()
library(readr)   # write_excel_csv() for UTF-8 CSV
library(here)    # here()

# 2. Import helpers ------------------------------------------------------------
source(here("helpers.R"), local = TRUE, encoding = "UTF-8")

# 3. Define main function
full_extractor <- function(docx_path) {
  # a) Extract context sentences -----------------------------------------------
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
  
  # b) Extract endnotes --------------------------------------------------------
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
  
  # c) Extract hyperlinks ------------------------------------------------------
  rels_map <- read_endnote_rels(temp_dir)
  hypers <- purrr::map(notes, ~ extract_hyperlinks(.x, ns_notes, rels_map))

  # d) Build initial endnotes --------------------------------------------------
  endnotes_df <- tibble(
    number  = ids,
    full_endnote = vapply(notes, get_note_text, character(1)),
    hyperlinks = hypers
  )
  
  # e) Join the two dataframes, add page references, and clean things up -------
  final_df <- endnotes_df |>
    # un-nest hyperlinks if there are multiple in 1 endnote 
    unnest_longer(hyperlinks, values_to = "source_link", keep_empty = TRUE) |>
    # find page numbers for each source
    mutate(page_ref = map2_chr(full_endnote, source_link, get_pages_by_source)) |>
    # join the two dataframes
    left_join(context_df, by = "number") |>
    # re-order things
    select(number, context_sentence, source_link, page_ref, full_endnote)
  
  # f) Cleanup ---------------------------------------------------------------
  unlink(temp_dir, recursive = TRUE)
  
  # g) Return ---------------------------------------------------------------
  final_df
}

# 3. Run extraction --------------------------------------------------------
  # a) Read in the file:
    docx_path <- file.choose() # Choose path
    
    # (i) Unzip .docx into a temp folder
    temp_dir <- tempfile("docx_unzip_")
    dir.create(temp_dir)
    unzip(zipfile = docx_path, exdir = temp_dir)
    
    # (ii) Paths to the two XMLs 
    doc_xml      <- file.path(temp_dir, "word", "document.xml")
    endnotes_xml <- file.path(temp_dir, "word", "endnotes.xml")
    
    # (iii)) Read and namespace 
    doc_main  <- read_xml(doc_xml)
    ns_main   <- xml_ns(doc_main)
    
    doc_notes <- read_xml(endnotes_xml)
    ns_notes  <- xml_ns(doc_notes)

  # b) Run the extractor:
  extracted_info <- full_extractor(docx_path)

# 4. Save output ---------------------------------------------------------------
save_path <- getwd()

# Write out to CSV
write_excel_csv(
  extracted_info,
  file = file.path(save_path, output_file)
)
message("Results saved to your working directory. Be sure to copy that file to your project folder!")