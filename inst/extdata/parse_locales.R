#' TBX Function Name Extractor
#'
#' Parses Microsoft Terminology Collection TBX files to produce a single
#' data frame with one row per Excel function and one column per locale.
#'
#' Output shape
#' ------------
#' fn          : English function name  (e.g. "SUM")
#' description : English description    (e.g. "Adds its arguments")
#' de, fr, …   : localised names        (e.g. "SUMME", "SOMME", …)
#'
#' Approach
#' --------
#' Uses openxlsx2::read_xml / xml_node / xml_value / xml_attr to navigate
#' the TBX structure directly — no streaming, no regex over raw XML.
#' The filter `grepl("Excel function", descriptions)` is applied first so
#' only relevant termEntry nodes are processed.
#'
#' Parallelism
#' -----------
#' Each TBX file is parsed independently. All files are dispatched in parallel
#' via parallel::mclapply (Linux/macOS) or parallel::parLapply (Windows).
#' Set workers = 1 to disable.
#'
#' Requires: openxlsx2, parallel (base R)

library(parallel)

# ---------------------------------------------------------------------------
# 1.  Function list from the installed package registries
# ---------------------------------------------------------------------------

EXCEL_FUNCTIONS <- sort(toupper(unique(c(
  spaghetti:::.spaghetti_env$LEGACY_WORKSHEET,
  spaghetti:::.spaghetti_env$LEGACY_XLM,
  spaghetti:::.spaghetti_env$XLFN,
  spaghetti:::.spaghetti_env$XLWS
))))

EXCEL_FUNCTIONS_SET <- new.env(hash = TRUE, parent = emptyenv())
for (.f in EXCEL_FUNCTIONS) assign(.f, TRUE, envir = EXCEL_FUNCTIONS_SET)
rm(.f)

# ---------------------------------------------------------------------------
# 2.  Name cleaning helpers
# ---------------------------------------------------------------------------

#' Clean a raw term string into an uppercase Excel function name.
#'
#' Strips trailing annotations like "()", "function", "-Funktion",
#' "-fonction", etc. that appear in TBX term entries.
#' Returns NA_character_ if the result is empty or contains spaces.
#' @keywords internal
.clean_term <- function(raw) {
  if (is.null(raw) || is.na(raw) || !nzchar(trimws(raw))) return(NA_character_)

  x <- trimws(raw)

  # Strip trailing "()" — e.g. "SINGLE()" -> "SINGLE"
  x <- sub("\\(\\)\\s*$", "", x)

  # Strip trailing function-word in any supported language, with optional
  # leading hyphen or space: "-Funktion", " function", "(fonction)", etc.
  x <- trimws(gsub(
    paste0("(?i)[\\s-]*\\(?(",
           "function|funktion|funktionen|",
           "fonction|fonctions|",
           "funzione|funzioni|",
           "functie|functies|",
           "función|funciones|",
           "função|funções|",
           "funkcja|funkcje|",
           "funktion|funktioner",  # sv/da/no
           ")\\)?\\s*$"),
    "", x, perl = TRUE
  ))
  x <- trimws(x)

  if (!nzchar(x))        return(NA_character_)
  if (grepl("\\s", x))  return(NA_character_)   # multi-word phrase, not a name

  toupper(x)
}

#' Validate that a cleaned name is a known Excel function.
#' Also strips accidental _xlfn. prefix that occasionally appears in TBX data.
#' @keywords internal
.validate_en <- function(en_clean) {
  if (is.na(en_clean)) return(NA_character_)
  # Occasional _xlfn. prefix in the TBX source data
  candidate <- sub("^_XLFN\\.", "", en_clean)
  if (exists(candidate, envir = EXCEL_FUNCTIONS_SET, inherits = FALSE))
    return(candidate)
  NA_character_
}

# ---------------------------------------------------------------------------
# 3.  Single-file parser using openxlsx2
# ---------------------------------------------------------------------------

#' Parse one TBX file and return a data frame with columns:
#'   fn (English name), description, id, <locale_col> (localised name)
#'
#' Mirrors the interactive approach that is known to work:
#'   1. read_xml + xml_node to get all termEntry nodes
#'   2. extract_row() on each node to get a named vector of all columns
#'   3. rbindlist to assemble the raw frame
#'   4. Filter to known Excel functions via EXCEL_FUNCTIONS_SET
#'
#' @param tbx_path   Path to the .tbx file.
#' @param locale_tag xml:lang value for the target locale (e.g. "de").
#' @param locale_col Column name for the localised names (e.g. "de").
#'
#' @return data.frame(fn, description, id, <locale_col>)
parse_tbx <- function(tbx_path, locale_tag, locale_col = locale_tag) {
  stopifnot(file.exists(tbx_path))
  message("  [", locale_col, "] ", basename(tbx_path))

  # ── Load XML ─────────────────────────────────────────────────────────────
  raw_xml      <- paste0(readLines(tbx_path, warn = FALSE, encoding = "UTF-8"),
                         collapse = "")
  xml          <- openxlsx2::read_xml(raw_xml)
  term_entries <- openxlsx2::xml_node(xml, c("martif", "text", "body", "termEntry"))

  if (length(term_entries) == 0) {
    message("    -> 0 termEntry nodes found")
    return(.empty_frame(locale_col))
  }

  # ── Pre-filter: keep only entries that mention "Excel function" ──────────
  # Use the description vector exactly as in the working interactive example.
  descriptions <- openxlsx2::xml_value(
    term_entries,
    c("termEntry", "langSet", "descripGrp", "descrip")
  )
  sel <- grepl("Excel (worksheet )?function", descriptions, ignore.case = TRUE)
  fun_entries  <- term_entries[sel]

  message("    -> ", length(fun_entries), " 'Excel function' entries found")
  if (length(fun_entries) == 0) return(.empty_frame(locale_col))

  # ── extract_row: mirrors the working interactive example exactly ──────────
  # Produces a named character vector with columns:
  #   id, desc, en-US, <locale_tag>  (plus any other lang tags present)
  extract_row <- function(entry) {
    term_id <- unlist(openxlsx2::xml_attr(entry, "termEntry"))
    langs   <- unname(unlist(openxlsx2::xml_attr(entry, "termEntry", "langSet")))
    desc    <- openxlsx2::xml_value(
      entry, c("termEntry", "langSet", "descripGrp", "descrip")
    )
    terms   <- openxlsx2::xml_value(
      entry, c("termEntry", "langSet", "ntig", "termGrp", "term")
    )
    names(desc)  <- "desc"
    names(terms) <- langs
    c(term_id = term_id, desc, terms)
  }

  raw_list <- lapply(fun_entries, extract_row)
  df_raw   <- openxlsx2:::rbindlist(raw_list)

  if (nrow(df_raw) == 0) return(.empty_frame(locale_col))

  # ── Extract English name and validate against spec function list ──────────
  en_col <- if ("en-US" %in% names(df_raw)) "en-US" else
    names(df_raw)[grepl("^en", names(df_raw))][1]

  if (is.na(en_col) || !en_col %in% names(df_raw))
    return(.empty_frame(locale_col))

  en_clean <- vapply(df_raw[[en_col]], function(x) {
    v <- .clean_term(x)
    if (is.na(v)) NA_character_ else .validate_en(v)
  }, character(1))

  df_raw <- df_raw[!is.na(en_clean), , drop = FALSE]
  en_clean <- en_clean[!is.na(en_clean)]
  if (nrow(df_raw) == 0) return(.empty_frame(locale_col))

  # ── Find the localised name column ───────────────────────────────────────
  # locale_tag is the xml:lang value sniffed from the file (e.g. "de").
  # Try exact match first, then prefix match (e.g. "de" matches "de-DE").
  loc_col <- if (locale_tag %in% names(df_raw)) {
    locale_tag
  } else {
    candidates <- names(df_raw)[startsWith(names(df_raw), substring(locale_tag, 1, 2))]
    candidates <- setdiff(candidates, c("en-US", en_col, "desc", "term_id"))
    if (length(candidates) > 0) candidates[1] else NA_character_
  }

  if (is.na(loc_col)) {
    message("    -> locale column '", locale_tag, "' not found in file")
    return(.empty_frame(locale_col))
  }

  # ── Build output frame ───────────────────────────────────────────────────
  loc_raw   <- df_raw[[loc_col]]
  loc_clean <- vapply(loc_raw, .clean_term, character(1))

  keep <- !is.na(loc_clean) &
    loc_clean != en_clean &
    !grepl("\\s", loc_clean)

  df_out <- data.frame(
    fn          = en_clean[keep],
    description = trimws(df_raw$desc[keep]),
    term_id     = df_raw$term_id[keep],
    stringsAsFactors = FALSE,
    check.names = FALSE
  )
  df_out[[locale_col]] <- loc_clean[keep]

  # Keep first match per English function name
  df_out <- df_out[!duplicated(df_out$fn), ]
  message("    -> ", nrow(df_out), " unique mappings retained")
  df_out
}

#' @keywords internal
.empty_frame <- function(locale_col) {
  df <- data.frame(fn = character(0), description = character(0),
                   term_id = character(0),
                   stringsAsFactors = FALSE, check.names = FALSE)
  df[[locale_col]] <- character(0)
  df
}

# ---------------------------------------------------------------------------
# 4.  Auto-detect TBX files in a folder
# ---------------------------------------------------------------------------

#' Scan a directory for .tbx files and determine their locale from the XML.
#'
#' Rather than guessing from filenames (which are unreliable across the 100+
#' Microsoft TBX files), this reads the xml:lang attribute of the first
#' non-English <langSet> inside each file. That attribute is authoritative
#' and works regardless of how the file is named.
#'
#' The locale code used as the column name is the xml:lang value with any
#' script/region subtag stripped to two characters where unambiguous
#' (e.g. "de-DE" -> "de", "zh-Hans" -> "zh-Hans" kept as-is since "zh" is
#' ambiguous between Simplified and Traditional).
#'
#' @param folder Path to directory containing .tbx files.
#' @return Named character vector: names = locale codes, values = file paths.
detect_tbx_files <- function(folder) {
  files <- list.files(folder, pattern = "\\.tbx$", full.names = TRUE,
                      ignore.case = TRUE)
  if (length(files) == 0)
    stop("No .tbx files found in: ", folder)

  result <- character(0)

  for (f in files) {
    locale <- .sniff_locale(f)
    if (is.na(locale)) {
      warning("Could not detect locale in: ", basename(f), " — skipping",
              call. = FALSE)
      next
    }
    if (locale %in% names(result)) {
      warning("Duplicate locale '", locale, "' for ", basename(f),
              " (already have ", basename(result[[locale]]), ") — skipping",
              call. = FALSE)
      next
    }
    result[[locale]] <- f
  }

  if (length(result) == 0)
    stop("No locales could be detected in TBX files in: ", folder)

  message("Detected ", length(result), " locale(s): ",
          paste(sort(names(result)), collapse = ", "))
  result
}

#' Read the first few KB of a TBX file and extract the non-English xml:lang tag.
#' @keywords internal
.sniff_locale <- function(path) {
  xml <- openxlsx2::read_xml(paste0(readLines(path, warn = FALSE, encoding = "UTF-8"), collapse = ""))
  entries <- unique(unlist(openxlsx2::xml_attr(xml, c("martif", "text", "body", "termEntry", "langSet"))))
  setdiff(entries, "en-US")
}

# ---------------------------------------------------------------------------
# 5.  Main driver — parallel parse + join into one data frame
# ---------------------------------------------------------------------------

#' Parse all TBX files and produce a flat function/locale data frame.
#'
#' @param folder      Directory of .tbx files (auto-detected).
#' @param tbx_files   Named list of locale → path (overrides folder).
#' @param outfile_rds Path for the output RDS (loaded at package startup).
#' @param outfile_r   Path for a human-readable R source copy.
#' @param workers     Parallel workers. Set to 1 to disable parallelism.
#' @param locale_tags Named list mapping locale code → xml:lang tag.
#'
#' @return The master data frame (invisibly).
#'
#' @examples
#' \dontrun{
#' # Auto-detect:
#' parse_all_tbx("~/Downloads/MicrosoftTermCollection")
#'
#' # Explicit:
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
      "zh-CN" = "zh-Hans", "zh-TW" = "zh-Hant", "pt-BR" = "pt-BR",
      "nb" = "nb-NO", "nn" = "nn-NO", "sr-Cyrl" = "sr-Cyrl-RS",
      "bs-Cyrl" = "bs-Cyrl-BA", "pa-Arab" = "pa-Arab-PK",
      af="af", sq="sq", am="am", ar="ar", hy="hy", as="as", az="az-Latn",
      bn="bn", eu="eu", be="be", bs="bs-Latn", bg="bg", my="my", ca="ca",
      ckb="ku-Arab", chr="chr-Cher", hr="hr", cs="cs", da="da", prs="prs",
      nl="nl", et="et", fil="fil", fi="fi", fr="fr", ff="ff-Latn", gl="gl",
      ka="ka", de="de", el="el", gu="gu", ha="ha-Latn", he="he", hi="hi",
      hu="hu", is="is", ig="ig", id="id", iu="iu-Latn", ga="ga", it="it",
      xh="xh", zu="zu", ja="ja", quc="quc", kn="kn", kk="kk", km="km",
      rw="rw", sw="sw", kok="kok", ko="ko", ky="ky", lo="lo", lv="lv",
      lt="lt", lb="lb", mk="mk", ms="ms", ml="ml", mt="mt", mi="mi",
      mr="mr", mn="mn", ne="ne", or="or", ps="ps", fa="fa", pl="pl",
      pt="pt-PT", pa="pa", qu="qu", ro="ro", ru="ru", gd="gd",
      sr="sr-Latn", nso="nso", tn="tn", sd="sd", si="si", sk="sk",
      sl="sl", es="es", sv="sv", tg="tg-Cyrl", ta="ta", tt="tt-Cyrl",
      te="te", th="th", ti="ti", tr="tr", tk="tk-Latn", uk="uk",
      ur="ur", ug="ug", "uz-Cyrl"="uz-Cyrl", uz="uz-Latn",
      "ca-valencia"="ca-ES-valencia", vi="vi", cy="cy", wo="wo", yo="yo"
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

  # ── Build one job per locale ─────────────────────────────────────────────
  jobs <- lapply(locales, function(loc) {
    list(
      tbx_path   = tbx_files[[loc]],
      locale_tag = if (loc %in% names(locale_tags)) locale_tags[[loc]] else loc,
      locale_col = loc
    )
  })
  names(jobs) <- locales

  .run_job <- function(job) {
    parse_tbx(job$tbx_path, job$locale_tag, job$locale_col)
  }

  # ── Parallel dispatch ────────────────────────────────────────────────────
  on_windows <- .Platform$OS.type == "windows"

  locale_frames <- if (workers > 1L && !on_windows) {
    parallel::mclapply(jobs, .run_job, mc.cores = workers)
  } else if (workers > 1L && on_windows) {
    cl <- parallel::makeCluster(workers)
    on.exit(parallel::stopCluster(cl), add = TRUE)
    parallel::clusterExport(cl, c(
      "parse_tbx", ".clean_term", ".validate_en", ".empty_frame",
      "EXCEL_FUNCTIONS_SET"
    ), envir = environment())  # .sniff_locale not needed in workers
    parallel::clusterEvalQ(cl, library(openxlsx2))
    parallel::parLapply(cl, jobs, .run_job)
  } else {
    lapply(jobs, .run_job)
  }

  # ── Seed master frame from canonical function list ───────────────────────
  master <- data.frame(
    fn          = EXCEL_FUNCTIONS,
    description = NA_character_,
    term_id     = NA_character_,
    stringsAsFactors = FALSE
  )

  # ── Left-join each locale result ─────────────────────────────────────────
  for (loc in locales) {
    lf <- locale_frames[[loc]]

    # Backfill descriptions (English descriptions appear in every TBX file)
    if (nrow(lf) > 0 && "description" %in% names(lf)) {
      desc_map     <- setNames(lf$description, lf$fn)
      needs_desc   <- is.na(master$description) & master$fn %in% names(desc_map)
      master$description[needs_desc] <- desc_map[master$fn[needs_desc]]
    }

    # Add locale column
    if (nrow(lf) > 0 && loc %in% names(lf)) {
      loc_map       <- setNames(lf[[loc]], lf$fn)
      master[[loc]] <- loc_map[master$fn]
    } else {
      master[[loc]] <- NA_character_
    }

    # Backfill id (take from first locale that provides it)
    if (nrow(lf) > 0 && "term_id" %in% names(lf) && "term_id" %in% names(master)) {
      id_map <- setNames(lf$term_id, lf$fn)
      needs_id <- is.na(master$term_id) & master$fn %in% names(id_map)
      master$term_id[needs_id] <- id_map[master$fn[needs_id]]
    }
  }

  # ── Coverage report ──────────────────────────────────────────────────────
  n_fn <- nrow(master)
  message(sprintf("\nMaster frame: %d functions x %d locale(s)", n_fn, length(locales)))
  message("Coverage per locale:")
  for (loc in locales) {
    n_translated <- sum(!is.na(master[[loc]]))
    message(sprintf("  %-12s : %d / %d (%.0f%%)",
                    loc, n_translated, n_fn, 100 * n_translated / n_fn))
  }

  # ── Write RDS ────────────────────────────────────────────────────────────
  if (!is.null(outfile_rds)) {
    dir.create(dirname(outfile_rds), showWarnings = FALSE, recursive = TRUE)
    saveRDS(master, outfile_rds)
    message("\nSaved RDS : ", outfile_rds)
  }

  # ── Write human-readable R source ────────────────────────────────────────
  if (!is.null(outfile_r)) {
    dir.create(dirname(outfile_r), showWarnings = FALSE, recursive = TRUE)
    con <- file(outfile_r, open = "wt", encoding = "UTF-8")
    cat(
      "# Auto-generated by data-raw/parse_locales.R — DO NOT EDIT MANUALLY\n",
      "# Source : Microsoft Terminology Collection TBX files\n",
      "#          https://www.microsoft.com/en-us/language/terminology\n",
      "# Columns: fn, description, ", paste(locales, collapse = ", "), "\n\n",
      file = con, sep = ""
    )
    cat(".spaghetti_env$FUNCTIONS <- data.frame(\n", file = con)
    all_cols <- c("fn", "description", "term_id", locales)
    for (ci in seq_along(all_cols)) {
      col  <- all_cols[[ci]]
      last <- ci == length(all_cols)
      vals <- master[[col]]
      cat(sprintf("  %s = c(\n", col), file = con)
      for (i in seq_along(vals)) {
        comma <- if (i < length(vals)) "," else ""
        v <- if (is.na(vals[i])) "NA_character_"
        else paste0('"', gsub('"', '\\"', vals[i], fixed = TRUE), '"')
        cat(sprintf("    %s%s\n", v, comma), file = con)
      }
      cat(if (last) "  )\n" else "  ),\n", file = con)
    }
    cat("  stringsAsFactors = FALSE\n)\n", file = con)
    close(con)
    message("Saved R source: ", outfile_r)
  }

  invisible(master)
}

# ---------------------------------------------------------------------------
# 5b. Download + verify + unzip + parse driver
# ---------------------------------------------------------------------------

# Source URL for the Microsoft Terminology Collection. Documented at
# https://learn.microsoft.com/en-us/globalization/reference/microsoft-terminology
MTC_URL <- "https://download.microsoft.com/download/b/2/d/b2db7a7c-8d33-47f3-b2c1-ee5e6445cf45/MicrosoftTermCollection.zip"

#' Download the Microsoft Terminology Collection zip, verify it (optionally)
#' against an expected SHA-256, unzip it to a tempdir, parse all .tbx files
#' it contains, and write the resulting master data frame to dest_rds.
#'
#' Intended to be invoked via `spaghetti::setup_terminology()` but can be
#' called directly if you `source()` this script yourself.
#'
#' @param dest_rds        Path to write the parsed RDS to.
#' @param expected_sha256 Optional. If supplied, verify the downloaded
#'                        zip against this hex digest. If NULL, the
#'                        observed digest is printed so you can record
#'                        it for future invocations.
#' @param workers         Parallel workers for TBX parsing.
#' @param quiet           If TRUE, suppress progress messages.
#' @return The master data frame (invisibly).
download_and_parse_tbx <- function(dest_rds,
                                   expected_sha256 = NULL,
                                   workers = max(1L, parallel::detectCores() - 1L),
                                   quiet = FALSE) {

  if (!requireNamespace("digest", quietly = TRUE)) {
    stop("Package 'digest' is required for SHA-256 verification. ",
         "Install it with:\n  install.packages(\"digest\")", call. = FALSE)
  }
  if (!requireNamespace("openxlsx2", quietly = TRUE)) {
    stop("Package 'openxlsx2' is required to parse TBX files. ",
         "Install it with:\n  install.packages(\"openxlsx2\")", call. = FALSE)
  }

  tmp <- tempfile("MicrosoftTermCollection_", fileext = ".zip")
  on.exit(unlink(tmp), add = TRUE)

  if (!quiet) message("Downloading: ", MTC_URL, "\n  -> ", tmp)
  status <- tryCatch(
    utils::download.file(MTC_URL, tmp, mode = "wb", quiet = quiet),
    error = function(e) e
  )
  if (inherits(status, "error")) {
    stop("Download failed: ", conditionMessage(status), call. = FALSE)
  }
  if (!file.exists(tmp) || file.size(tmp) == 0L) {
    stop("Download produced no file (or empty file).", call. = FALSE)
  }

  observed <- digest::digest(file = tmp, algo = "sha256")
  if (!quiet) message("Downloaded ", round(file.size(tmp) / 1024^2, 1),
                      " MB; SHA-256 = ", observed)

  if (!is.null(expected_sha256)) {
    exp_norm <- tolower(gsub("[^0-9a-fA-F]", "", expected_sha256))
    if (!identical(exp_norm, tolower(observed))) {
      stop("SHA-256 mismatch.\n",
           "  expected: ", exp_norm, "\n",
           "  observed: ", observed, "\n",
           "Aborting before parsing. If you trust the new file, ",
           "re-invoke with expected_sha256 = NULL or update the expected ",
           "value.", call. = FALSE)
    }
    if (!quiet) message("SHA-256 OK.")
  } else if (!quiet) {
    message("(No expected_sha256 supplied; verification skipped. ",
            "Record the digest above to pin future runs.)")
  }

  extract_dir <- tempfile("tbx_")
  dir.create(extract_dir, recursive = TRUE)
  on.exit(unlink(extract_dir, recursive = TRUE), add = TRUE)

  if (!quiet) message("Unzipping into: ", extract_dir)
  utils::unzip(tmp, exdir = extract_dir)

  if (!quiet) message("Parsing TBX files...")
  master <- parse_all_tbx(folder      = extract_dir,
                          outfile_rds = dest_rds,
                          outfile_r   = NULL,
                          workers     = workers)

  if (!quiet) message("Cached terminology saved to: ", dest_rds)
  invisible(master)
}

# ---------------------------------------------------------------------------
# 6.  Run
# ---------------------------------------------------------------------------

# Auto-detect all TBX files in the folder.
# Guarded so `source()`ing the file just defines the functions; run by either
# calling parse_all_tbx() directly, or executing this script as the top-level
# command (Rscript inst/extdata/parse_locales.R).
if (sys.nframe() == 0L && interactive() == FALSE) {
  result <- parse_all_tbx("~/Downloads/MicrosoftTermCollection")
}

# Or supply paths explicitly:
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
