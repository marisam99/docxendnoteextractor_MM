# --------------------------------------------------------------------------
# Combined Endnote, Source, & Context Extractor for Word (.docx)
#  – Unzips a .docx once
#  – Reads word/document.xml to pull the sentence immediately before each
#    <w:endnoteReference>
#  – Reads word/endnotes.xml to pull the endnote text and its Roman numeral
#  - Reads word/rels.xml to pull source links out of the endnote text
#  – Returns a tibble: number | context_sentence | source_link | full_citation
# --------------------------------------------------------------------------

# 0. SETUP: CHANGE THIS AS NEEDED ----------------------------------------------
output_file <- "AIFP_measurement_fact_check.csv"
# docx_path <- "C:/Users/Marisa Mission/OneDrive - BW/2025 AIFP Year-Long AI Funder Collaborative/WIP/Measurement/2025-08-05 Measurement draft.docx"

# 1. Packages ------------------------------------------------------------------
library(xml2)    # parse & query XML
library(tidyr)   # unnest()
library(dplyr)   # mutate, native pipe |>
library(tibble)  # tibble()
library(purrr)   # map_chr()
library(stringr) # str_squish(), str_split()
library(utils)   # unzip(), unlink()
library(readr)   # write_excel_csv() for UTF-8 CSV

# 2. Function definitions ------------------------------------------------------
read_endnote_rels <- function(temp_dir) {
  rels_path <- file.path(temp_dir, "word", "_rels", "endnotes.xml.rels")
  if (!file.exists(rels_path)) {
    return(tibble(r_id = character(), target = character(), type = character()))
  }
  rels_doc  <- read_xml(rels_path)
  rel_nodes <- xml_find_all(
    rels_doc, "/*[local-name()='Relationships']/*[local-name()='Relationship']"
  )
  tibble(
    r_id   = xml_attr(rel_nodes, "Id"),
    target = xml_attr(rel_nodes, "Target"),
    type   = xml_attr(rel_nodes, "Type")
  ) |>
    filter(grepl("hyperlink$", type))
}

canonicalize_urls <- function(urls) {
  if (length(urls) == 0) return(character(0))
  urls <- str_trim(urls)
  # remove zero-width spaces (Word sometimes injects these)
  urls <- str_replace_all(urls, "[\u200B\u200C\u200D]", "")
  # strip leading opening punctuation/quotes/brackets
  urls <- str_replace(urls, "^[\\(\\[\\{<\"'“‘]+", "")
  # strip trailing closing punctuation/quotes/brackets and common end-of-sentence chars
  urls <- str_replace(urls, "[\\)\\]\\}>\"'”’\\.;:,!?]+$", "")
  # collapse any trailing slashes beyond one (avoids …///)
  urls <- sub("/+$", "/", urls)
  # de-dupe
  unique(urls[!is.na(urls) & urls != ""])
}

extract_note_hyperlinks <- function(note_node, ns_notes, rels_map) {
  urls <- character(0)

  # (a) <w:hyperlink r:id="...">...</w:hyperlink> → resolve via rels_map
  hyper_nodes <- xml_find_all(note_node, ".//w:hyperlink[@r:id]", ns_notes)
  if (length(hyper_nodes)) {
    rids <- xml_attr(hyper_nodes, "r:id", ns = ns_notes)
    urls <- c(urls, rels_map$target[match(rids, rels_map$r_id)])
  }

  # (b) Simple field: <w:fldSimple w:instr='HYPERLINK "https://..."'>
  fld_nodes <- xml_find_all(note_node, ".//w:fldSimple[@w:instr]", ns_notes)
  if (length(fld_nodes)) {
    instr <- xml_attr(fld_nodes, "w:instr", ns = ns_notes)
    mats  <- str_match_all(instr, 'HYPERLINK\\s+"([^"]+)"')
    urls  <- c(urls, unlist(lapply(mats, function(m) if (ncol(m) > 1) m[, 2] else character(0))))
  }

  # (c) Complex fields: <w:instrText>HYPERLINK "https://..."</w:instrText>
  instr_texts <- xml_find_all(note_node, ".//w:instrText", ns_notes) |> xml_text()
  if (length(instr_texts)) {
    mats <- str_match_all(paste(instr_texts, collapse = " "), 'HYPERLINK\\s+"([^"]+)"')
    urls <- c(urls, unlist(lapply(mats, function(m) if (ncol(m) > 1) m[, 2] else character(0))))
  }

  # (d) Fallback ALWAYS ON: naked URLs pasted into the text
  txt <- xml_find_all(note_node, ".//w:t", ns_notes) |> 
    xml_text() |> 
    paste(collapse = "")
  naked <- str_extract_all(txt, "https?://[^\\s\\)\\]\\}\\>\"'”’]+")
  if (length(naked) && length(naked[[1]])) urls <- c(urls, naked[[1]])

  canonicalize_urls(urls)
}

full_extractor <- function(docx_path) {
  # a) Extract context sentences --------------------------------------------
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
  
  # b) Extract endnotes ------------------------------------------------------
  notes <- xml_find_all(
    doc_notes,
    ".//w:endnote[not(@w:type) and number(@w:id) > 0]",
    ns_notes
  )
  
  ids <- xml_attr(notes, "w:id", ns = ns_notes) |> as.integer() # Word's internal IDs
  
  get_note_text <- function(node) {
    xml_find_all(node, ".//w:t", ns_notes) |>
      xml_text() |>
      paste(collapse = "") |>
      str_squish()
  }
  
  # c) Extract hyperlinks ------------------------------------------------------
  rels_map <- read_endnote_rels(temp_dir)
  hypers <- purrr::map(notes, ~ extract_note_hyperlinks(.x, ns_notes, rels_map))

  # d) Build initial endnotes df -------------------------------------------------------
  endnotes_df <- tibble(
    number  = ids,
    full_citation = vapply(notes, get_note_text, character(1)),
    hyperlinks = hypers
  )
  
  #e) Join the two dataframes and clean things up -------------------------------------
  final_df <- endnotes_df |>
    # un-nest hyperlinks if there are multiple in 1 endnote 
    unnest_longer(hyperlinks, values_to = "source_link", keep_empty = TRUE) |>
    # join the two dataframes
    left_join(context_df, by = "number") |>
    # re-order things
    select(number, context_sentence, source_link, full_citation)
  
  # f) Cleanup ---------------------------------------------------------------
  unlink(temp_dir, recursive = TRUE)
  
  # g) Return ---------------------------------------------------------------
  final_df
}

# 3. Run the Extraction --------------------------------------------------------
# Set up the extraction:
  docx_path <- file.choose() # Choose path
  
  # a) Unzip .docx into a temp folder
  temp_dir <- tempfile("docx_unzip_")
  dir.create(temp_dir)
  unzip(zipfile = docx_path, exdir = temp_dir)
  
  # b) Paths to the two XMLs 
  doc_xml      <- file.path(temp_dir, "word", "document.xml")
  endnotes_xml <- file.path(temp_dir, "word", "endnotes.xml")
  
  # c) Read and namespace 
  doc_main  <- read_xml(doc_xml)
  ns_main   <- xml_ns(doc_main)
  
  doc_notes <- read_xml(endnotes_xml)
  ns_notes  <- xml_ns(doc_notes)

# Run the extractor:
extracted_info <- full_extractor(docx_path)

# 4. Save Output ---------------------------------------------------------------

save_path <- getwd()
# Write out to CSV
write_excel_csv(
  extracted_info,
  file = file.path(save_path, output_file)
)
message("Results saved to your working directory. Be sure to copy that file to your project folder!")