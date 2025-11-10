# --------------------------------------------------------------------------
# Helper functions for the Endnote Citation Extractor
# --------------------------------------------------------------------------

# 1. Helper functions for extracting hyperlinks ----------------------------

  # a) reads relationships to create a map of hyperlinks
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

  # b) normalizes hyperlink urls to remove funky things
  normalize_urls <- function(urls) {
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

  # c) uses the first two to actually extract hyperlinks
  extract_hyperlinks <- function(notes, ns_notes, rels_map) {
    urls <- character(0)

    # (i) <w:hyperlink r:id="...">...</w:hyperlink> → resolve via rels_map
    hyper_nodes <- xml_find_all(notes, ".//w:hyperlink[@r:id]", ns_notes)
    if (length(hyper_nodes)) {
      rids <- xml_attr(hyper_nodes, "r:id", ns = ns_notes)
      urls <- c(urls, rels_map$target[match(rids, rels_map$r_id)])
    }

    # (ii) Simple field: <w:fldSimple w:instr='HYPERLINK "https://..."'>
    fld_nodes <- xml_find_all(notes, ".//w:fldSimple[@w:instr]", ns_notes)
    if (length(fld_nodes)) {
      instr <- xml_attr(fld_nodes, "w:instr", ns = ns_notes)
      mats  <- str_match_all(instr, 'HYPERLINK\\s+"([^"]+)"')
      urls  <- c(urls, unlist(lapply(mats, function(m) if (ncol(m) > 1) m[, 2] else character(0))))
    }

    # (iii) Complex fields: <w:instrText>HYPERLINK "https://..."</w:instrText>
    instr_texts <- xml_find_all(notes, ".//w:instrText", ns_notes) |> xml_text()
    if (length(instr_texts)) {
      mats <- str_match_all(paste(instr_texts, collapse = " "), 'HYPERLINK\\s+"([^"]+)"')
      urls <- c(urls, unlist(lapply(mats, function(m) if (ncol(m) > 1) m[, 2] else character(0))))
    }

    # (iv) Fallback ALWAYS ON: naked URLs pasted into the text
    txt <- xml_find_all(notes, ".//w:t", ns_notes) |> 
      xml_text() |> 
      paste(collapse = "")
    naked <- str_extract_all(txt, "https?://[^\\s\\)\\]\\}\\>\"'”’]+")
    if (length(naked) && length(naked[[1]])) urls <- c(urls, naked[[1]])

    normalize_urls(urls)
  }

# 2. Helper functions for extracting page references -----------------------
  
  # a) if there are more than one citation in the endnote, this splits the entire endnote into individual citations
  #     if not, then it returns the whole endnote
  split_endnote <- function(endnote) {
    if (is.na(endnote) || !nzchar(endnote)) return(character(0))
    ch <- strsplit(endnote, "", fixed = TRUE)[[1]]
    out <- character(); buf <- character()
    p <- b <- c <- 0; q <- FALSE
    isq <- function(z) z %in% c('"', "'", "“","”","‘","’")
    for (z in ch) {
      if (!q && z=="(") p <- p+1 else if (!q && z==")") p <- max(0,p-1)
      else if (!q && z=="[") b <- b+1 else if (!q && z=="]") b <- max(0,b-1)
      else if (!q && z=="{") c <- c+1 else if (!q && z=="}") c <- max(0,c-1)
      else if (isq(z)) q <- !q
      if (z==";" && p==0 && b==0 && c==0 && !q) { out <- c(out, str_squish(paste0(buf, collapse=""))); buf <- character(); next }
      buf <- c(buf, z)
    }
    out <- c(out, str_squish(paste0(buf, collapse="")))
    out[nzchar(out)]
  }

  # b) given a segment of the endnote, this extracts page numbers using "p." or "pp." as a delineator
  extract_pages <- function(segment) {  
    # build regex pattern to find pages
    pattern <- paste0(
      "(?i)\\bpp?\\.?\\s+",                                        # p. / pp. (case-insensitive)
      "(?:[A-Za-z]?\\d+[A-Za-z]?|[ivxlcdm]+)",                     # 12, S12, 12a, or roman
      "(?:\\s*(?:[\u2013-]|\\bto\\b)\\s*",                         # – or - or 'to'
      "(?:[A-Za-z]?\\d+[A-Za-z]?|[ivxlcdm]+))?",                   # optional range end
      "(?:\\s*,\\s*(?:[A-Za-z]?\\d+[A-Za-z]?|[ivxlcdm]+)",         # additional list items
      "(?:\\s*(?:[\u2013-]|\\bto\\b)\\s*(?:[A-Za-z]?\\d+[A-Za-z]?|[ivxlcdm]+))?)*"
    )
    
    # find the first matching chunks followed by numbers
    hit <- str_extract_all(segment, pattern)[[1]]

    # if it's not matched, return NA
    if (length(hit) == 0) return(" ")

    # if it is matched, clean it up and return
    pages <- hit |>
      stringr::str_replace_all("\\s*to\\s*", "–") |>
      stringr::str_replace_all("\\s*-\\s*", "–") |>
      stringr::str_replace_all("\\s+", "")

    # return
    pages
  }

  # c) given a full endnote and one url, it will match the url to the correct citation and,
  #     if available, pull page numbers for that source
  get_pages_by_source <- function(endnote, url) {
    segments <- split_endnote(endnote)
    idx <- which(vapply(segments, function(s) grepl(url, s, fixed = TRUE), logical(1)))[1]
    target_seg <- if(length(idx) && !is.na(idx)) segments[idx] else endnote
    # target_seg <- segments[idx]
    pages <- extract_pages(target_seg)
    paste(pages, collapse = ", ")  # Collapse multiple page refs into single string
  }