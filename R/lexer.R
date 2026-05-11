# R/lexer.R
# Lexical analyser for formula strings.
#
# Formula syntax is not regular: string literals can contain
# function-name-like tokens, and the @ and # operators are contextual.
# A simple regex gsub over the whole formula will corrupt string content.
#
# This tokeniser walks the formula character-by-character, identifying:
#   - STRING   : "..." literals (with "" escape sequences)
#   - FUNC     : identifier immediately followed by '('
#   - IDENT    : identifier not followed by '(' or matching a cell-ref pattern
#                (named range, LAMBDA param, defined name)
#   - REF      : cell reference like A1, $B$2, A1:B10, Sheet1!A1,
#                'My Sheet'!A1, Sheet1!A1:B10
#   - ANCHOR   : the # spill operator (follows a cell ref)
#   - IMPLICIT : the @ implicit intersection operator
#   - OTHER    : any other character (operators, delimiters, numbers)

TOKEN_TYPES <- list(
  STRING   = "STRING",
  FUNC     = "FUNC",
  IDENT    = "IDENT",
  REF      = "REF",
  ANCHOR   = "ANCHOR",
  IMPLICIT = "IMPLICIT",
  OTHER    = "OTHER"
)

# Regexes to classify a bare token as a cell-shape reference. Used to
# decide REF vs IDENT after lexing.
.CELL_RX <- "^\\$?[A-Za-z]+\\$?[0-9]+$"   # A1, $A$1
.COL_RX  <- "^\\$?[A-Za-z]+$"             # A, $AB (whole column)
.ROW_RX  <- "^\\$?[0-9]+$"                # 1, $10 (whole row)

# TRUE iff `s` matches any reference shape (cell, column, row).
.is_ref_shape <- function(s) {
  grepl(.CELL_RX, s) || grepl(.COL_RX, s) || grepl(.ROW_RX, s)
}

#' Tokenise a formula string
#'
#' @param formula Character scalar, optionally starting with '='.
#' @param sep The argument separator (e.g., ',' or ';').
#' @return A list of token objects, each with fields `type` and `val`.
#' @keywords internal
.tokenise <- function(formula, sep = ",") {
  if (startsWith(formula, "=")) formula <- substring(formula, 2L)

  chars <- strsplit(formula, "", fixed = TRUE)[[1]]
  n     <- length(chars)
  pos   <- 1L
  tokens   <- vector("list", n)
  n_tokens <- 0L

  emit <- function(type, val) {
    n_tokens <<- n_tokens + 1L
    if (n_tokens > length(tokens)) length(tokens) <<- length(tokens) * 2L
    tokens[[n_tokens]] <<- list(type = type, val = val)
  }

  # Cell-ref characters: A-Z a-z 0-9 $
  is_ref_char  <- function(c) grepl("[A-Za-z0-9$]", c, perl = TRUE)
  is_id_start  <- function(c) grepl("[A-Za-z_]",      c, perl = TRUE)
  is_id_cont   <- function(c) grepl("[A-Za-z0-9_.]",  c, perl = TRUE)
  is_digit_dot <- function(c) grepl("[0-9.]",         c, perl = TRUE)

  consume_while <- function(predicate) {
    start <- pos
    while (pos <= n && predicate(chars[pos])) pos <<- pos + 1L
    if (pos == start) "" else paste(chars[start:(pos - 1L)], collapse = "")
  }

  # Try to consume a cell/column/row reference, optionally followed by
  # ":<ref>" to form a range. Returns the consumed string or "" if nothing
  # matched. Both endpoints of a range must share the same shape (two
  # cells, two columns, or two rows); mixed-shape ranges aren't real
  # references.
  consume_cell_or_range <- function() {
    start <- pos
    first <- consume_while(is_ref_char)
    if (!nzchar(first) || !.is_ref_shape(first)) {
      pos <<- start
      return("")
    }
    if (pos <= n && chars[pos] == ":") {
      pos <<- pos + 1L
      second_start <- pos
      second <- consume_while(is_ref_char)
      if (!nzchar(second) || !.is_ref_shape(second)) {
        pos <<- second_start - 1L
        return(first)
      }
      return(paste0(first, ":", second))
    }
    first
  }

  while (pos <= n) {
    ch <- chars[pos]

    # ---- String literal "..." -----------------------------------------
    if (ch == '"') {
      start <- pos
      pos   <- pos + 1L          # opening quote
      while (pos <= n) {
        if (chars[pos] == '"') {
          if (pos + 1L <= n && chars[pos + 1L] == '"') {
            pos <- pos + 2L      # escaped ""
          } else {
            pos <- pos + 1L      # closing quote
            break
          }
        } else {
          pos <- pos + 1L
        }
      }
      emit(TOKEN_TYPES$STRING, paste(chars[start:(pos - 1L)], collapse = ""))
      next
    }

    # ---- Quoted sheet name 'My Sheet'!ref → REF -----------------------
    if (ch == "'") {
      start <- pos
      pos   <- pos + 1L          # opening quote
      while (pos <= n) {
        if (chars[pos] == "'") {
          if (pos + 1L <= n && chars[pos + 1L] == "'") {
            pos <- pos + 2L      # escaped ''
          } else {
            pos <- pos + 1L      # closing quote
            break
          }
        } else {
          pos <- pos + 1L
        }
      }
      sheet <- paste(chars[start:(pos - 1L)], collapse = "")
      # Expect '!' followed by a cell ref
      if (pos <= n && chars[pos] == "!") {
        pos <- pos + 1L
        ref <- consume_cell_or_range()
        if (nzchar(ref)) {
          emit(TOKEN_TYPES$REF, paste0(sheet, "!", ref))
        } else {
          emit(TOKEN_TYPES$OTHER, paste0(sheet, "!"))
        }
      } else {
        emit(TOKEN_TYPES$OTHER, sheet)
      }
      next
    }

    # ---- @ implicit intersection --------------------------------------
    if (ch == "@") {
      pos <- pos + 1L
      emit(TOKEN_TYPES$IMPLICIT, "@")
      next
    }

    # ---- # spill anchor -----------------------------------------------
    if (ch == "#") {
      pos <- pos + 1L
      emit(TOKEN_TYPES$ANCHOR, "#")
      next
    }

    # ---- Argument separator (locale-dependent) ------------------------
    if (ch == sep) {
      pos <- pos + 1L
      emit(TOKEN_TYPES$OTHER, ch)
      next
    }

    # ---- $ starting an absolute cell ref ($A$1, $A1, A$1) -------------
    if (ch == "$") {
      start <- pos
      ref <- consume_cell_or_range()
      if (nzchar(ref)) {
        emit(TOKEN_TYPES$REF, ref)
      } else {
        pos <- start + 1L
        emit(TOKEN_TYPES$OTHER, ch)
      }
      next
    }

    # ---- Identifier: function name, named range, REF, sheet prefix ----
    if (is_id_start(ch)) {
      start <- pos
      pos   <- pos + 1L
      while (pos <= n && is_id_cont(chars[pos])) pos <- pos + 1L
      ident <- paste(chars[start:(pos - 1L)], collapse = "")

      # Sheet-qualified ref:  Sheet1!A1[:B10]
      if (pos <= n && chars[pos] == "!") {
        pos <- pos + 1L
        ref <- consume_cell_or_range()
        if (nzchar(ref)) {
          emit(TOKEN_TYPES$REF, paste0(ident, "!", ref))
          next
        }
        emit(TOKEN_TYPES$OTHER, paste0(ident, "!"))
        next
      }

      # Function call: identifier immediately followed by '('
      if (pos <= n && chars[pos] == "(") {
        emit(TOKEN_TYPES$FUNC, ident)
        next
      }

      # Cell-shape bare identifier (A1, B10, AB12). Detect a following
      # ":<ref>" as a range. We deliberately do NOT classify bare letter
      # sequences (e.g. `B`, `XYZ`) as REF here — they're ambiguous with
      # LAMBDA parameters and named ranges. The downstream code handles
      # the column-range "A:Z" case as three OTHER+IDENT tokens, which
      # round-trip correctly through .detokenise.
      if (grepl(.CELL_RX, ident)) {
        if (pos <= n && chars[pos] == ":") {
          pos <- pos + 1L
          second_start <- pos
          second <- consume_while(is_ref_char)
          if (nzchar(second) && .is_ref_shape(second)) {
            emit(TOKEN_TYPES$REF, paste0(ident, ":", second))
            next
          }
          pos <- second_start - 1L
        }
        emit(TOKEN_TYPES$REF, ident)
        next
      }

      emit(TOKEN_TYPES$IDENT, ident)
      next
    }

    # ---- Numbers ------------------------------------------------------
    # Digits, optional decimal point, optional e/E exponent with sign.
    # +/- only allowed immediately after e/E; outside that it's an operator.
    if (is_digit_dot(ch)) {
      start <- pos
      prev  <- ""
      while (pos <= n) {
        p <- chars[pos]
        if (grepl("[0-9.]", p, perl = TRUE)) {
          pos <- pos + 1L; prev <- p
        } else if (p == "e" || p == "E") {
          pos <- pos + 1L; prev <- p
        } else if ((p == "+" || p == "-") && (prev == "e" || prev == "E")) {
          pos <- pos + 1L; prev <- p
        } else {
          break
        }
      }
      emit(TOKEN_TYPES$OTHER, paste(chars[start:(pos - 1L)], collapse = ""))
      next
    }

    # ---- Everything else ----------------------------------------------
    pos <- pos + 1L
    emit(TOKEN_TYPES$OTHER, ch)
  }

  if (n_tokens < length(tokens)) length(tokens) <- n_tokens
  tokens
}

#' Reconstruct a formula string from a token list
#'
#' Separator tokens (`,`/`;`) are already present in the token stream and
#' have been swapped by the transform passes, so this is a simple concat.
#'
#' @param tokens List of token objects.
#' @param prefix_eq Logical; prepend '=' if TRUE.
#' @return Character scalar.
#' @keywords internal
.detokenise <- function(tokens, prefix_eq = TRUE) {
  parts  <- vapply(tokens, function(t) t$val, character(1))
  result <- paste(parts, collapse = "")
  if (prefix_eq) paste0("=", result) else result
}
