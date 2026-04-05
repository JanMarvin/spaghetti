#' TBX Function Name Extractor
#'
#' Parses Microsoft Terminology Collection TBX files to produce a single
#' data frame with one row per Excel function and one column per locale.
#'
#' Output shape
#' ------------
#' fn         : English function name  (e.g. "SUM")
#' description: English description    (e.g. "Adds its arguments")
#' de, fr, …  : localised names        (e.g. "SUMME", "SOMME", …)
#'
#' This flat structure makes it trivial to:
#'   - Add a new locale: parse one more TBX, left-join the result column
#'   - Inspect coverage: filter rows where a locale column is NA
#'   - Power a UI widget: description field is ready to use as tooltip text
#'
#' Parallelism
#' -----------
#' Each TBX file is parsed independently, so all files are processed in
#' parallel using parallel::mclapply (fork-based, Linux/macOS) or
#' parallel::parLapply (socket cluster, Windows-safe).
#' Set workers = 1 to disable parallelism for debugging.
#'
#' No packages beyond base R are required.

library(parallel)

# ---------------------------------------------------------------------------
# 1.  Function list — pulled directly from the installed package registries
# ---------------------------------------------------------------------------

EXCEL_FUNCTIONS <- toupper(unique(c(
  spaghetti:::.spaghetti_env$LEGACY,
  spaghetti:::.spaghetti_env$XLFN,
  spaghetti:::.spaghetti_env$XLWS
)))

EXCEL_FUNCTIONS_SET <- new.env(hash = TRUE, parent = emptyenv())
for (.f in EXCEL_FUNCTIONS) assign(.f, TRUE, envir = EXCEL_FUNCTIONS_SET)
rm(.f)

# ---------------------------------------------------------------------------
# 2.  Low-level XML extraction helpers
# ---------------------------------------------------------------------------

#' Extract the raw content of a <langSet xml:lang="TAG"> block.
#' @keywords internal
.extract_langset <- function(entry, lang_tag) {
  tag_esc <- gsub("([.\\^$*+?{}()|\\[\\]])", "\\\\\\1", lang_tag)
  pat <- paste0('<langSet[^>]+xml:lang="', tag_esc, '"[^>]*>(.*?)</langSet>')
  m   <- regexpr(pat, entry, perl = TRUE, ignore.case = TRUE)
  if (m == -1L) return(NULL)
  regmatches(entry, m)
}

#' Extract the text content of the first <term>…</term> in a block.
#' @keywords internal
.extract_term <- function(block) {
  m <- regexpr('<term[^>]*>([^<]+)</term>', block, perl = TRUE)
  if (m == -1L) return(NULL)
  s <- attr(m, "capture.start")[1L]
  l <- attr(m, "capture.length")[1L]
  if (is.na(s) || s < 1L) return(NULL)
  substr(block, s, s + l - 1L)
}

#' Extract the text content of the first <descrip type="definition">…</descrip>.
#' Falls back to any <descrip> if no definition-typed one is found.
#' @keywords internal
.extract_descrip <- function(block) {
  # Prefer type="definition"
  m <- regexpr('<descrip[^>]+type="definition"[^>]*>([^<]+)</descrip>',
               block, perl = TRUE, ignore.case = TRUE)
  if (m == -1L) {
    # Fallback: any <descrip>
    m <- regexpr('<descrip[^>]*>([^<]+)</descrip>', block, perl = TRUE)
  }
  if (m == -1L) return(NA_character_)
  s <- attr(m, "capture.start")[1L]
  l <- attr(m, "capture.length")[1L]
  if (is.na(s) || s < 1L) return(NA_character_)
  trimws(substr(block, s, s + l - 1L))
}

# ---------------------------------------------------------------------------
# 3.  Per-entry parser — returns list(en, desc, loc) or NULL
# ---------------------------------------------------------------------------

#' Parse one termEntry block.
#'
#' @param entry     Raw XML string for one <termEntry>…</termEntry>.
#' @param locale_tag xml:lang tag for the target locale (e.g. "de").
#' @return list(en = "SUM", desc = "Adds its arguments", loc = "SUMME")
#'         or NULL if entry is not a relevant Excel function.
#' @keywords internal
.extract_row <- function(entry, locale_tag) {

  # Only bother with entries that mention Excel somewhere in a description
  if (!grepl("<descrip[^>]*>.*?Excel.*?</descrip>",
             entry, ignore.case = TRUE, perl = TRUE)) return(NULL)

  # ── English term ────────────────────────────────────────────────────────
  en_block <- .extract_langset(entry, "en-US")
  if (is.null(en_block)) return(NULL)

  en_raw <- .extract_term(en_block)
  if (is.null(en_raw)) return(NULL)

  # Strip trailing annotation: "function", "()", "(constant)", etc.
  en_stripped <- trimws(gsub("(?i)\\s*\\(?function\\)?\\s*$|\\(\\)\\s*$",
                             "", en_raw, perl = TRUE))
  if (nchar(en_stripped) == 0) return(NULL)

  en_clean <- toupper(en_stripped)

  # Strip accidental _xlfn. prefix that occasionally appears in TBX data
  if (!exists(en_clean, envir = EXCEL_FUNCTIONS_SET, inherits = FALSE)) {
    en_clean <- sub("^_XLFN\\.", "", en_clean)
  }
  if (!exists(en_clean, envir = EXCEL_FUNCTIONS_SET, inherits = FALSE)) return(NULL)
  if (grepl("\\s", en_clean)) return(NULL)

  # ── English description ─────────────────────────────────────────────────
  desc <- .extract_descrip(en_block)

  # ── Localised term ──────────────────────────────────────────────────────
  loc_block <- .extract_langset(entry, locale_tag)
  if (is.null(loc_block)) return(NULL)

  loc_raw <- .extract_term(loc_block)
  if (is.null(loc_raw)) return(NULL)

  # Strip locale-language annotation suffixes ("-Funktion", "-fonction", etc.)
  loc_clean <- trimws(gsub(
    paste0("(?i)[\\s-]*\\(",
           "funktion|fonction|funzione|functie|",
           "funkcja|funktioner|función|função",
           "\\)?\\s*$"),
    "", loc_raw, perl = TRUE
  ))
  loc_clean <- toupper(trimws(loc_clean))

  if (is.na(loc_clean) || nchar(loc_clean) < 2) return(NULL)
  if (loc_clean == en_clean)   return(NULL)   # untranslated — skip
  if (grepl("\\s", loc_clean)) return(NULL)   # multi-word — not a function name

  list(en = en_clean, desc = desc, loc = loc_clean)
}

# ---------------------------------------------------------------------------
# 4.  Streaming parser for one TBX file
# ---------------------------------------------------------------------------

#' Stream a TBX file and return a two-column data frame: fn, <locale>.
#'
#' @param tbx_path   Path to the .tbx file.
#' @param locale_tag BCP-47 tag matching xml:lang in the file (e.g. "de").
#' @param locale_col Column name for the localised names in the output frame.
#' @param chunk_size Bytes per read chunk (default 1 MB).
#'
#' @return data.frame(fn = character, <locale_col> = character)
#'         Only rows where a translated name was found are included.
parse_tbx <- function(tbx_path, locale_tag,
                      locale_col  = locale_tag,
                      chunk_size  = 1e6) {

  stopifnot(file.exists(tbx_path))
  con <- file(tbx_path, open = "rb")
  on.exit(close(con), add = TRUE)

  buffer    <- ""
  rows      <- list()
  n_scanned <- 0L

  message("  [", locale_col, "] Parsing: ", basename(tbx_path))

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
      n_scanned <- n_scanned + 1L
      row <- .extract_row(entry, locale_tag)
      if (!is.null(row)) rows[[length(rows) + 1L]] <- row
    }
  }

  # Flush last entry in buffer
  if (grepl("</termEntry>", buffer, fixed = TRUE)) {
    row <- .extract_row(buffer, locale_tag)
    if (!is.null(row)) rows[[length(rows) + 1L]] <- row
  }

  message(sprintf("    -> %d entries scanned, %d functions matched",
                  n_scanned, length(rows)))

  if (length(rows) == 0) {
    return(data.frame(fn          = character(0),
                      description = character(0),
                      stringsAsFactors = FALSE,
                      check.names = FALSE))
  }

  df <- data.frame(
    fn          = vapply(rows, `[[`, character(1), "en"),
    description = vapply(rows, `[[`, character(1), "desc"),
    stringsAsFactors = FALSE,
    check.names = FALSE
  )
  df[[locale_col]] <- vapply(rows, `[[`, character(1), "loc")

  # Keep only the first match per English function name
  df[!duplicated(df$fn), ]
}

# ---------------------------------------------------------------------------
# 5.  Auto-detect TBX files in a folder
# ---------------------------------------------------------------------------

#' Detect TBX files in a directory and infer locale codes from filenames.
#'
#' Filename → locale heuristic (case-insensitive):
#'   GERMAN      → de    FRENCH   → fr    SPANISH  → es
#'   ITALIAN     → it    DUTCH    → nl    PORTUGUESE → pt
#'   POLISH      → pl    SWEDISH  → sv    DANISH   → da
#'   FINNISH     → fi    NORWEGIAN → no   CZECH    → cs
#'   HUNGARIAN   → hu    ROMANIAN → ro    TURKISH  → tr
#'
#' Unrecognised filenames are skipped with a warning.
#'
#' @param folder Path to directory containing .tbx files.
#' @return Named character vector: names = locale codes, values = file paths.
detect_tbx_files <- function(folder) {
  name_to_locale <- c(
    german      = "de", deutsch    = "de",
    french      = "fr", francais   = "fr",
    spanish     = "es", espanol    = "es",
    italian     = "it", italiano   = "it",
    dutch       = "nl", nederlands = "nl",
    portuguese  = "pt", portugues  = "pt",
    polish      = "pl", polski     = "pl",
    swedish     = "sv", svenska    = "sv",
    danish      = "da", dansk      = "da",
    finnish     = "fi", suomi      = "fi",
    norwegian   = "no", norsk      = "no",
    czech       = "cs",
    hungarian   = "hu",
    romanian    = "ro",
    turkish     = "tr"
  )

  files  <- list.files(folder, pattern = "\\.tbx$", full.names = TRUE,
                       ignore.case = TRUE)
  result <- character(0)

  for (f in files) {
    base  <- tolower(tools::file_path_sans_ext(basename(f)))
    # Try longest match first (e.g. "portuguese (portugal)")
    match <- NA_character_
    for (nm in names(name_to_locale)) {
      if (grepl(nm, base, fixed = TRUE)) {
        match <- name_to_locale[[nm]]
        break
      }
    }
    if (is.na(match)) {
      warning("Cannot infer locale for: ", basename(f), " — skipping")
    } else {
      result[[match]] <- f
    }
  }

  if (length(result) == 0)
    stop("No recognised TBX files found in: ", folder)

  message("Detected ", length(result), " locale(s): ",
          paste(names(result), collapse = ", "))
  result
}

# ---------------------------------------------------------------------------
# 6.  Main driver — parallel parse + join into one data frame
# ---------------------------------------------------------------------------

#' Parse all TBX files in a folder and write a flat data frame.
#'
#' The output data frame has columns:
#'   fn          : English function name
#'   description : English description (from TBX definition element)
#'   <locale>    : One column per locale, containing the localised name
#'                 (NA where no translation was found)
#'
#' The frame is written to two places:
#'   outfile_rds : Binary RDS (fast to load at package startup)
#'   outfile_r   : Human-readable R source (for version control / inspection)
#'
#' @param folder     Directory containing .tbx files (auto-detected).
#' @param tbx_files  Optional named vector of locale → path (overrides folder).
#' @param outfile_rds Destination for the RDS file.
#' @param outfile_r   Destination for the generated R source file.
#' @param workers    Number of parallel workers. Defaults to detected cores - 1,
#'                   capped at the number of locale files. Set to 1 to disable.
#' @param locale_tags Named list overriding the xml:lang tag per locale code.
#'
#' @return The data frame (invisibly). Side effect: writes outfile_rds and
#'         outfile_r.
#'
#' @examples
#' \dontrun{
#' # Auto-detect all TBX files in a folder:
#' parse_all_tbx("~/Downloads/MicrosoftTermCollection")
#'
#' # Or supply paths explicitly:
#' parse_all_tbx(tbx_files = list(
#'   de = "~/Downloads/MicrosoftTermCollection/GERMAN.tbx",
#'   fr = "~/Downloads/MicrosoftTermCollection/FRENCH.tbx"
#' ))
#' }
parse_all_tbx <- function(
    folder      = NULL,
    tbx_files   = NULL,
    outfile_rds = "inst/extdata/excel_functions.rds",
    outfile_r   = "data-raw/excel_functions_generated.R",
    workers     = max(1L, parallel::detectCores(logical = FALSE) - 1L),
    locale_tags = list(
      de = "de", fr = "fr", es = "es", it = "it",
      nl = "nl", pt = "pt", pl = "pl", sv = "sv",
      da = "da", fi = "fi", no = "no", cs = "cs",
      hu = "hu", ro = "ro", tr = "tr"
    )
) {

  # ── Resolve file list ────────────────────────────────────────────────────
  if (is.null(tbx_files)) {
    if (is.null(folder)) stop("Provide either `folder` or `tbx_files`.")
    tbx_files <- detect_tbx_files(folder)
  }

  locales <- names(tbx_files)
  workers <- min(workers, length(locales))
  message("\n=== Parsing ", length(locales), " TBX file(s) with ",
          workers, " worker(s) ===\n")

  # ── Build argument list for each parse job ───────────────────────────────
  jobs <- lapply(locales, function(loc) {
    list(
      tbx_path   = tbx_files[[loc]],
      locale_tag = if (loc %in% names(locale_tags)) locale_tags[[loc]] else loc,
      locale_col = loc
    )
  })
  names(jobs) <- locales

  # ── Parallel dispatch ────────────────────────────────────────────────────
  .run_job <- function(job) {
    parse_tbx(job$tbx_path, job$locale_tag, job$locale_col)
  }

  on_windows <- .Platform$OS.type == "windows"

  locale_frames <- if (workers > 1L && !on_windows) {
    parallel::mclapply(jobs, .run_job, mc.cores = workers)
  } else if (workers > 1L && on_windows) {
    cl <- parallel::makeCluster(workers)
    on.exit(parallel::stopCluster(cl), add = TRUE)
    # Export everything the workers need
    parallel::clusterExport(cl, c(
      "parse_tbx", ".extract_row", ".extract_langset",
      ".extract_term", ".extract_descrip", "EXCEL_FUNCTIONS_SET"
    ), envir = environment())
    parallel::parLapply(cl, jobs, .run_job)
  } else {
    lapply(jobs, .run_job)
  }

  # ── Seed with the full English function list ─────────────────────────────
  # Start from the canonical function list so every function appears as a row
  # even if no locale could translate it (those columns will be NA).
  master <- data.frame(
    fn          = sort(EXCEL_FUNCTIONS),
    description = NA_character_,
    stringsAsFactors = FALSE
  )

  # ── Left-join each locale result onto the master frame ───────────────────
  for (loc in locales) {
    lf <- locale_frames[[loc]]
    if (nrow(lf) == 0) {
      master[[loc]] <- NA_character_
      next
    }

    # Fill in descriptions from this locale's parse (English descriptions
    # are in every TBX file; take the first non-NA we encounter)
    desc_map <- setNames(lf$description, lf$fn)
    missing_desc <- is.na(master$description) & master$fn %in% names(desc_map)
    master$description[missing_desc] <- desc_map[master$fn[missing_desc]]

    # Add localised name column
    loc_map       <- setNames(lf[[loc]], lf$fn)
    master[[loc]] <- loc_map[master$fn]
  }

  message(sprintf(
    "\nMaster frame: %d functions x %d locale(s)",
    nrow(master), length(locales)
  ))

  coverage <- vapply(locales, function(l) {
    sum(!is.na(master[[l]]))
  }, integer(1))
  message("Coverage per locale:")
  for (loc in locales)
    message(sprintf("  %-4s : %d / %d", loc, coverage[[loc]], nrow(master)))

  # ── Write RDS (fast binary, used by the package at runtime) ──────────────
  if (!is.null(outfile_rds)) {
    dir.create(dirname(outfile_rds), showWarnings = FALSE, recursive = TRUE)
    saveRDS(master, outfile_rds)
    message("\nSaved RDS: ", outfile_rds)
  }

  # ── Write human-readable R source (for version control) ──────────────────
  if (!is.null(outfile_r)) {
    dir.create(dirname(outfile_r), showWarnings = FALSE, recursive = TRUE)
    con <- file(outfile_r, open = "wt", encoding = "UTF-8")

    cat(
      "# Auto-generated by data-raw/parse_locales.R — DO NOT EDIT MANUALLY\n",
      "# Source : Microsoft Terminology Collection TBX files\n",
      "#          https://www.microsoft.com/en-us/language/terminology\n",
      "# Columns: fn, description, ",
      paste(locales, collapse = ", "), "\n\n",
      file = con, sep = ""
    )

    cat(".spaghetti_env$FUNCTIONS <- data.frame(\n", file = con)

    write_col <- function(col_name, values, last = FALSE) {
      is_char <- is.character(values)
      cat(sprintf("  %s = c(\n", col_name), file = con)
      for (i in seq_along(values)) {
        comma <- if (i < length(values)) "," else ""
        if (is_char) {
          v <- if (is.na(values[i])) "NA_character_"
          else paste0('"', gsub('"', '\\"', values[i], fixed = TRUE), '"')
        } else {
          v <- as.character(values[i])
        }
        cat(sprintf("    %s%s\n", v, comma), file = con)
      }
      trailing <- if (last) "\n  )" else "\n  ),"
      cat(trailing, "\n", file = con, sep = "")
    }

    all_cols  <- c("fn", "description", locales)
    for (ci in seq_along(all_cols)) {
      col  <- all_cols[[ci]]
      last <- ci == length(all_cols)
      write_col(col, master[[col]], last = last)
    }

    cat("  stringsAsFactors = FALSE\n)\n", file = con)
    close(con)
    message("Saved R source: ", outfile_r)
  }

  invisible(master)
}

# ---------------------------------------------------------------------------
# 7.  Run
# ---------------------------------------------------------------------------

# Auto-detect all TBX files in the folder:
result <- parse_all_tbx("~/Downloads/MicrosoftTermCollection")

# Or supply paths explicitly, e.g.:
# result <- parse_all_tbx(tbx_files = list(
#   de = "~/Downloads/MicrosoftTermCollection/GERMAN.tbx",
#   fr = "~/Downloads/MicrosoftTermCollection/FRENCH.tbx",
#   es = "~/Downloads/MicrosoftTermCollection/SPANISH.tbx",
#   it = "~/Downloads/MicrosoftTermCollection/ITALIAN.tbx",
#   nl = "~/Downloads/MicrosoftTermCollection/DUTCH.tbx",
#   pt = "~/Downloads/MicrosoftTermCollection/PORTUGUESE (PORTUGAL).tbx",
#   pl = "~/Downloads/MicrosoftTermCollection/POLISH.tbx",
#   sv = "~/Downloads/MicrosoftTermCollection/SWEDISH.tbx"
# ))
