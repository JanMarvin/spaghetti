#' TBX Function Name Extractor
#'
#' Parses Microsoft Terminology Collection TBX files to extract localised
#' Excel function names. Downloads available at:
#' https://www.microsoft.com/en-us/language/terminology
#'
#' Why TBX instead of web scraping?
#' ---------------------------------
#' The TBX files are Microsoft's authoritative bilingual termbase — the same
#' source used to localise Office UI strings. They are more complete and more
#' accurate than the support.microsoft.com index tables.
#'
#' Approach
#' ---------
#' The files are ~50 MB of essentially one long XML line, so we cannot use
#' readLines() naively or load the whole thing into memory as a DOM.
#' Instead we use a streaming chunk reader that:
#'   1. Reads the file in raw character chunks
#'   2. Reassembles complete <termEntry>...</termEntry> blocks across chunk
#'      boundaries
#'   3. For each block, checks whether the en-US <term> matches a known
#'      Excel function name from the MS-XLSX spec list
#'   4. If it matches, extracts the localised <term>
#'
#' No packages beyond base R are required.

# ---------------------------------------------------------------------------
# 1.  The authoritative function list from [MS-XLSX] ABNF
#     (legacy + future — everything that can appear in a formula)
# ---------------------------------------------------------------------------

EXCEL_FUNCTIONS <- toupper(c(
  # -- function-list (legacy, no prefix) -----------------------------------
  spaghetti:::LEGACY,
  # -- future-function-list (_xlfn.) ---------------------------------------
  spaghetti:::XLFN,
  # -- 365 post-spec -------------------------------------------------------
  spaghetti:::XLWS
))

# Fast O(1) lookup via environment (faster than %in% for large sets)
EXCEL_FUNCTIONS_SET <- new.env(hash = TRUE, parent = emptyenv())
for (.f in EXCEL_FUNCTIONS) assign(.f, TRUE, envir = EXCEL_FUNCTIONS_SET)
rm(.f)

# ---------------------------------------------------------------------------
# 2.  Streaming TBX parser
# ---------------------------------------------------------------------------

#' Extract Excel function name mappings from one TBX file.
#'
#' @param tbx_path   Path to the .tbx file.
#' @param locale_tag BCP-47 language tag as used in the file's xml:lang
#'                   attributes, e.g. "de", "fr", "fr-FR". Matched
#'                   case-insensitively; both short and full forms work.
#' @param chunk_size Characters per read (default 1 MB). Larger values
#'                   are faster but use more memory.
#'
#' @return Named character vector: names = localised function names (upper),
#'         values = English function names (upper). Identical pairs excluded.
# ---------------------------------------------------------------------------
# 2.  Streaming TBX parser (Updated logic)
# ---------------------------------------------------------------------------
# ... [Keep Section 1: EXCEL_FUNCTIONS_SET exactly as in your file] ...

# ---------------------------------------------------------------------------
# 3.  termEntry extraction helpers (Fixed for .ETS and _xlfn)
# ---------------------------------------------------------------------------

.extract_pair <- function(entry, locale_tag) {

  # --- English term -------------------------------------------------------
  en_block <- .extract_langset(entry, "en-US")
  if (is.null(en_block)) return(NULL)

  en_raw <- .extract_term(en_block)
  if (is.null(en_raw)) return(NULL)

  # 1. Strip annotations like 'function', '()', or '(constant)'
  # We use ignore.case=TRUE because TBX varies between 'Function' and 'function'
  en_stripped <- trimws(gsub("(?i)\\s*\\(?function\\)?\\s*$", "", en_raw))
  en_stripped <- gsub("\\(\\)\\s*$", "", en_stripped)

  # 2. Safety check: If stripping resulted in empty string, abort
  if (nchar(en_stripped) == 0) return(NULL)

  en_clean <- toupper(en_stripped)

  # 3. Fallback: If not found, try stripping the Excel internal prefix '_xlfn.'
  # This is often why modern functions like FORECAST.ETS are missed
  is_known <- exists(en_clean, envir = EXCEL_FUNCTIONS_SET, inherits = FALSE)

  if (!is_known && grepl("^_XLFN\\.", en_clean)) {
    en_clean <- gsub("^_XLFN\\.", "", en_clean)
    is_known <- exists(en_clean, envir = EXCEL_FUNCTIONS_SET, inherits = FALSE)
  }

  # Final English validation
  if (!is_known) return(NULL)
  if (grepl("\\s", en_clean)) return(NULL) # Functions cannot have spaces

  # --- Localised term -----------------------------------------------------
  loc_block <- .extract_langset(entry, locale_tag)
  if (is.null(loc_block)) return(NULL)

  loc_raw <- .extract_term(loc_block)
  if (is.null(loc_raw)) return(NULL)

  # Clean German: Strip "-Funktion" and convert to Upper
  loc_clean <- trimws(gsub("(?i)[\\s-]*\\(?funktion\\)?\\s*$", "", loc_raw))
  loc_clean <- toupper(loc_clean)

  # Validation of the localized result
  if (is.na(loc_clean) || nchar(loc_clean) < 2) return(NULL)
  if (loc_clean == en_clean) return(NULL)  # Skip if not translated
  if (grepl("\\s", loc_clean)) return(NULL) # Skip descriptions with spaces

  list(en = en_clean, loc = loc_clean)
}

# ---------------------------------------------------------------------------
# 2.  Streaming TBX parser (Updated with Product Filter)
# ---------------------------------------------------------------------------

parse_tbx <- function(tbx_path, locale_tag, chunk_size = 1e6) {
  stopifnot(file.exists(tbx_path))

  con <- file(tbx_path, open = "rb")
  on.exit(close(con), add = TRUE)

  buffer    <- ""
  results   <- character(0)
  n_scanned <- 0L

  message("  Parsing: ", basename(tbx_path), " [locale: ", locale_tag, "]")

  repeat {
    raw <- readBin(con, what = "raw", n = chunk_size)
    if (!length(raw)) break
    chunk <- rawToChar(raw)
    Encoding(chunk) <- "UTF-8"

    buffer <- paste0(buffer, chunk)
    parts <- strsplit(buffer, "(?=<termEntry[ >])", perl = TRUE)[[1]]

    if (length(parts) > 1) {
      buffer   <- parts[[length(parts)]]
      complete <- parts[-length(parts)]
    } else {
      next
    }

    for (entry in complete) {
      if (!grepl("</termEntry>", entry, fixed = TRUE)) next

      # Ensure we only look at Excel entries to avoid MOICE/SSAS/DAX acronyms
      if (!grepl("<descrip[^>]*>.*?Excel.*?</descrip>", entry, ignore.case = TRUE)) next

      n_scanned <- n_scanned + 1L
      pair <- .extract_pair(entry, locale_tag)
      if (!is.null(pair)) results[[pair$loc]] <- pair$en
    }
  }

  message(sprintf("    %d entries scanned, %d Excel functions matched",
                  n_scanned, length(results)))
  results
}

#' Pull the content of the first <langSet xml:lang="TAG"> block.
#' @keywords internal
.extract_langset <- function(entry, lang_tag) {
  # lang_tag is literal (not a regex) — escape it for use in pattern
  tag_esc <- gsub("([.\\^$*+?{}()|\\[\\]])", "\\\\\\1", lang_tag)
  pat <- paste0('<langSet[^>]+xml:lang="', tag_esc, '"[^>]*>(.*?)</langSet>')
  m <- regexpr(pat, entry, perl = TRUE, ignore.case = TRUE)
  if (m == -1L) return(NULL)
  regmatches(entry, m)
}

#' Pull the text content of the first <term ...>TEXT</term>.
#' @keywords internal
.extract_term <- function(block) {
  m <- regexpr('<term[^>]*>([^<]+)</term>', block, perl = TRUE)
  if (m == -1L) return(NULL)
  starts  <- attr(m, "capture.start")
  lengths <- attr(m, "capture.length")
  if (is.na(starts[1]) || starts[1] < 1L) return(NULL)
  substr(block, starts[1], starts[1] + lengths[1] - 1L)
}

# ---------------------------------------------------------------------------
# 4.  Multi-file driver
# ---------------------------------------------------------------------------

#' Parse a named list of TBX files and write R/locales_generated.R
#'
#' @param tbx_files Named list: names = two-letter locale codes ("de","fr",…),
#'                  values = file paths to the corresponding TBX files.
#' @param outfile   Destination R file (default "R/locales_generated.R").
#'
#' @return Invisibly: list(fwd, rev) of the generated maps.
#'
#' @examples
#' \dontrun{
#' parse_all_tbx(list(
#'   de = "~/Downloads/GERMAN.tbx",
#'   fr = "~/Downloads/FRENCH.tbx",
#'   es = "~/Downloads/SPANISH.tbx",
#'   it = "~/Downloads/ITALIAN.tbx",
#'   nl = "~/Downloads/DUTCH.tbx",
#'   pt = "~/Downloads/PORTUGUESE.tbx",
#'   pl = "~/Downloads/POLISH.tbx",
#'   sv = "~/Downloads/SWEDISH.tbx"
#' ))
#' }
parse_all_tbx <- function(tbx_files,
                          outfile = "R/locales_generated.R") {

  # xml:lang tag used inside each TBX file (may differ from locale code)
  locale_tags <- list(
    de = "de", fr = "fr", es = "es", it = "it",
    nl = "nl", pt = "pt", pl = "pl", sv = "sv"
  )

  locales_fwd <- list()
  locales_rev <- list()

  for (loc in names(tbx_files)) {
    path <- tbx_files[[loc]]
    tag  <- if (loc %in% names(locale_tags)) locale_tags[[loc]] else loc

    message("\n[", loc, "]")
    raw <- parse_tbx(path, locale_tag = tag)

    fwd <- as.list(raw)
    rev <- list()
    for (loc_nm in names(fwd)) {
      en_nm <- fwd[[loc_nm]]
      if (!en_nm %in% names(rev)) rev[[en_nm]] <- loc_nm
    }

    locales_fwd[[loc]] <- fwd
    locales_rev[[loc]] <- rev
    message(sprintf("  -> %d unique mappings", length(fwd)))
  }

  # ---- Write generated R file -------------------------------------------
  message("\nWriting ", outfile, " ...")
  con <- file(outfile, open = "wt", encoding = "UTF-8")

  cat(
    "# Auto-generated by data-raw/parse_locales.R — DO NOT EDIT MANUALLY\n",
    "# Source : Microsoft Terminology Collection TBX files\n",
    "#          https://www.microsoft.com/en-us/language/terminology\n",
    "# Method : Streaming termEntry extraction matched to MS-XLSX spec list\n\n",
    "#' @keywords internal\n",
    file = con, sep = ""
  )

  .write_list <- function(varname, data, con) {
    cat(sprintf("%s <- list(\n", varname), file = con)
    locs <- sort(names(data))
    for (li in seq_along(locs)) {
      loc   <- locs[[li]]
      items <- data[[loc]]
      keys  <- sort(names(items))
      cat(sprintf("  %s = c(\n", loc), file = con)
      for (i in seq_along(keys)) {
        k  <- gsub('"', '\\"', keys[[i]],          fixed = TRUE)
        v  <- gsub('"', '\\"', items[[keys[[i]]]], fixed = TRUE)
        cm <- if (i < length(keys)) "," else ""
        cat(sprintf('    "%s" = "%s"%s\n', k, v, cm), file = con)
      }
      lc <- if (li < length(locs)) "," else ""
      cat(sprintf("  )%s\n", lc), file = con)
    }
    cat(")\n\n", file = con)
  }

  .write_list(".spaghetti_env$LOCALES",     locales_fwd, con)
  .write_list(".spaghetti_env$LOCALES_REV", locales_rev, con)
  close(con)

  total <- sum(lengths(locales_fwd))
  message("Done. ", total, " total locale mappings written to ", outfile)
  invisible(list(fwd = locales_fwd, rev = locales_rev))
}

# ---------------------------------------------------------------------------
# 5.  Run — edit paths to match where you saved the TBX files
# ---------------------------------------------------------------------------

parse_all_tbx(list(
  de = "~/Downloads/MicrosoftTermCollection/GERMAN.tbx",
  fr = "~/Downloads/MicrosoftTermCollection/FRENCH.tbx",
  es = "~/Downloads/MicrosoftTermCollection/SPANISH.tbx",
  it = "~/Downloads/MicrosoftTermCollection/ITALIAN.tbx",
  nl = "~/Downloads/MicrosoftTermCollection/DUTCH.tbx",
  pt = "~/Downloads/MicrosoftTermCollection/PORTUGUESE (PORTUGAL).tbx",
  pl = "~/Downloads/MicrosoftTermCollection/POLISH.tbx",
  sv = "~/Downloads/MicrosoftTermCollection/SWEDISH.tbx"
))
