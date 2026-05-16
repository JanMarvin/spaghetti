# R/lexer.R
# Lexical analyser for formula strings.
#
# Formula syntax is not regular: string literals can contain
# function-name-like tokens, and the @ and # operators are contextual.
# A simple regex gsub over the whole formula will corrupt string content.
#
# Token types:
#   - STRING   : "..." literals (with "" escape sequences)
#   - FUNC     : identifier immediately followed by '('
#   - IDENT    : identifier not followed by '(' or matching a cell-ref pattern
#                (named range, LAMBDA param, defined name)
#   - REF      : any cell reference. Includes:
#                  - cell:           A1, $B$2, A1:B10
#                  - sheet:          Sheet1!A1, 'My Sheet'!A1
#                  - 3D:             Sheet1:Sheet5!A1
#                  - external:       [Book1]Sheet1!A1, [1]Sheet1!A1
#                  - quoted external '[Book1]Sheet1'!A1
#                  - structured:     Table1[Col], Table1[#Headers], Table1[@Col]
#   - ANCHOR   : the # spill operator (follows a cell ref)
#   - IMPLICIT : the @ implicit intersection operator
#   - OTHER    : everything else (operators, separators, numbers, error
#                literals, array literals)

TOKEN_TYPES <- list(
  STRING   = "STRING",
  FUNC     = "FUNC",
  IDENT    = "IDENT",
  REF      = "REF",
  ANCHOR   = "ANCHOR",
  IMPLICIT = "IMPLICIT",
  OTHER    = "OTHER"
)

# Regexes to classify a bare token as a reference shape.
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

  is_ref_char  <- function(c) grepl("[A-Za-z0-9$]",   c, perl = TRUE)
  is_id_start  <- function(c) grepl("[A-Za-z_]",      c, perl = TRUE)
  is_id_cont   <- function(c) grepl("[A-Za-z0-9_.]",  c, perl = TRUE)
  is_digit_dot <- function(c) grepl("[0-9.]",         c, perl = TRUE)
  is_upper     <- function(c) grepl("[A-Z]",          c, perl = TRUE)

  substr_chars <- function(a, b) {
    if (a > b) "" else paste(chars[a:b], collapse = "")
  }

  consume_while <- function(predicate) {
    start <- pos
    while (pos <= n && predicate(chars[pos])) pos <<- pos + 1L
    substr_chars(start, pos - 1L)
  }

  # Consume an identifier (already known to start with id_start char).
  consume_ident <- function() {
    start <- pos
    pos   <<- pos + 1L
    while (pos <= n && is_id_cont(chars[pos])) pos <<- pos + 1L
    substr_chars(start, pos - 1L)
  }

  # Consume a quoted-sheet token starting at the '. Includes the opening
  # and closing quotes. Handles '' as an escape for a literal '.
  consume_quoted <- function() {
    start <- pos
    pos   <<- pos + 1L                    # opening '
    while (pos <= n) {
      if (chars[pos] == "'") {
        if (pos + 1L <= n && chars[pos + 1L] == "'") {
          pos <<- pos + 2L                # escaped ''
        } else {
          pos <<- pos + 1L                # closing '
          break
        }
      } else {
        pos <<- pos + 1L
      }
    }
    substr_chars(start, pos - 1L)
  }

  # Consume a [...] block (workbook tag or structured-ref body). Balances
  # nested [...] (used in structured refs like Table[[col1],[col2]]) and
  # respects quoted strings. The opening [ must be at chars[pos].
  consume_bracketed <- function() {
    start <- pos
    pos   <<- pos + 1L                    # opening [
    depth <- 1L
    while (pos <= n && depth > 0L) {
      c <- chars[pos]
      if (c == '"') {
        # skip string literal
        pos <<- pos + 1L
        while (pos <= n) {
          if (chars[pos] == '"') {
            if (pos + 1L <= n && chars[pos + 1L] == '"') {
              pos <<- pos + 2L
            } else {
              pos <<- pos + 1L
              break
            }
          } else {
            pos <<- pos + 1L
          }
        }
      } else if (c == "'") {
        # skip quoted name
        pos <<- pos + 1L
        while (pos <= n) {
          if (chars[pos] == "'") {
            if (pos + 1L <= n && chars[pos + 1L] == "'") {
              pos <<- pos + 2L
            } else {
              pos <<- pos + 1L
              break
            }
          } else {
            pos <<- pos + 1L
          }
        }
      } else if (c == "[") {
        depth <- depth + 1L
        pos   <<- pos + 1L
      } else if (c == "]") {
        depth <- depth - 1L
        pos   <<- pos + 1L
      } else {
        pos <<- pos + 1L
      }
    }
    substr_chars(start, pos - 1L)
  }

  # Try to consume a cell/column/row reference, optionally followed by
  # ":<ref>" to form a range. Returns the consumed string or "" if nothing
  # matched. Both endpoints must share the same shape.
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

  # Try to consume a sheet/3D prefix starting at pos. Returns the consumed
  # prefix (including trailing '!') or "" if nothing matched. Forms:
  #   Sheet1!         -> bare sheet
  #   Sheet1:Sheet5!  -> 3D sheet range
  #   'My Sheet'!     -> quoted sheet
  #   'Sheet1:Sheet5'! -> quoted 3D
  #   [Book1]Sheet1!  -> external + sheet
  #   '[Book1]Sheet1'! -> quoted external + sheet
  # If the optional [...] workbook tag is present, consume it first.
  consume_sheet_prefix <- function() {
    start <- pos
    workbook <- ""

    if (pos <= n && chars[pos] == "[") {
      workbook <- consume_bracketed()
    }

    if (pos > n) {
      pos <<- start
      return("")
    }

    if (chars[pos] == "'") {
      # quoted sheet name (may itself contain [Book])
      quoted <- consume_quoted()
      if (pos <= n && chars[pos] == "!") {
        pos <<- pos + 1L
        return(paste0(workbook, quoted, "!"))
      }
      pos <<- start
      return("")
    }

    if (is_id_start(chars[pos])) {
      sheet1 <- consume_ident()
      # Optional :Sheet2 for 3D ref
      if (pos + 1L <= n && chars[pos] == ":" && is_id_start(chars[pos + 1L])) {
        save <- pos
        pos  <<- pos + 1L
        sheet2 <- consume_ident()
        if (pos <= n && chars[pos] == "!") {
          pos <<- pos + 1L
          return(paste0(workbook, sheet1, ":", sheet2, "!"))
        }
        pos <<- save
      }
      if (pos <= n && chars[pos] == "!") {
        pos <<- pos + 1L
        return(paste0(workbook, sheet1, "!"))
      }
      # Bare identifier, no '!' — not actually a sheet prefix
    }

    pos <<- start
    ""
  }

  while (pos <= n) {
    ch <- chars[pos]

    # ---- String literal "..." -----------------------------------------
    if (ch == '"') {
      start <- pos
      pos   <- pos + 1L
      while (pos <= n) {
        if (chars[pos] == '"') {
          if (pos + 1L <= n && chars[pos + 1L] == '"') {
            pos <- pos + 2L
          } else {
            pos <- pos + 1L
            break
          }
        } else {
          pos <- pos + 1L
        }
      }
      emit(TOKEN_TYPES$STRING, substr_chars(start, pos - 1L))
      next
    }

    # ---- Array literal {...} → opaque OTHER token ---------------------
    # Inner ',' (column sep) and ';' (row sep) are part of the array,
    # not function-argument separators, so they must be shielded from
    # the locale-separator-normalisation pass.
    if (ch == "{") {
      start <- pos
      pos   <- pos + 1L
      depth <- 1L
      while (pos <= n && depth > 0L) {
        c <- chars[pos]
        if (c == '"') {
          # skip string literal inside array
          pos <- pos + 1L
          while (pos <= n) {
            if (chars[pos] == '"') {
              if (pos + 1L <= n && chars[pos + 1L] == '"') {
                pos <- pos + 2L
              } else {
                pos <- pos + 1L
                break
              }
            } else {
              pos <- pos + 1L
            }
          }
        } else if (c == "{") {
          depth <- depth + 1L
          pos   <- pos + 1L
        } else if (c == "}") {
          depth <- depth - 1L
          pos   <- pos + 1L
        } else {
          pos <- pos + 1L
        }
      }
      emit(TOKEN_TYPES$OTHER, substr_chars(start, pos - 1L))
      next
    }

    # ---- Quoted sheet name 'My Sheet'!ref OR external '[Book]Sheet'!ref
    if (ch == "'") {
      prefix <- consume_sheet_prefix()
      if (nzchar(prefix)) {
        ref <- consume_cell_or_range()
        if (nzchar(ref)) {
          emit(TOKEN_TYPES$REF, paste0(prefix, ref))
        } else {
          # Sheet prefix with no usable ref following — emit the prefix
          # alone as OTHER. Rare; lets odd inputs round-trip.
          emit(TOKEN_TYPES$OTHER, prefix)
        }
      } else {
        # Quoted thing that isn't a sheet prefix — emit as OTHER
        start <- pos
        quoted <- consume_quoted()
        emit(TOKEN_TYPES$OTHER, quoted)
      }
      next
    }

    # ---- External workbook ref [Book1]Sheet1!A1 -----------------------
    if (ch == "[") {
      prefix <- consume_sheet_prefix()
      if (nzchar(prefix)) {
        ref <- consume_cell_or_range()
        if (nzchar(ref)) {
          emit(TOKEN_TYPES$REF, paste0(prefix, ref))
          next
        }
        emit(TOKEN_TYPES$OTHER, prefix)
        next
      }
      # Standalone [ not a workbook ref (shouldn't normally occur outside
      # a structured-ref body, which is handled in the identifier branch).
      pos <- pos + 1L
      emit(TOKEN_TYPES$OTHER, ch)
      next
    }

    # ---- @ implicit intersection --------------------------------------
    if (ch == "@") {
      pos <- pos + 1L
      emit(TOKEN_TYPES$IMPLICIT, "@")
      next
    }

    # ---- # error literal OR spill anchor ------------------------------
    # An error literal is #<UPPER>... (e.g. #REF!, #N/A, #GETTING_DATA,
    # #SPILL!). The spill-anchor operator # is otherwise context-dependent:
    # it follows a ref and never precedes an uppercase letter directly.
    if (ch == "#") {
      if (pos + 1L <= n && is_upper(chars[pos + 1L])) {
        start <- pos
        pos   <- pos + 1L
        # Body: uppercase letters, digits, /, _
        while (pos <= n && grepl("[A-Z0-9/_]", chars[pos], perl = TRUE)) {
          pos <- pos + 1L
        }
        # Optional trailing ! or ?
        if (pos <= n && (chars[pos] == "!" || chars[pos] == "?")) {
          pos <- pos + 1L
        }
        emit(TOKEN_TYPES$OTHER, substr_chars(start, pos - 1L))
        next
      }
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

    # ---- Identifier: function, named range, REF, sheet prefix, table --
    if (is_id_start(ch)) {
      start_pos  <- pos
      sheet_pref <- consume_sheet_prefix()

      if (nzchar(sheet_pref)) {
        ref <- consume_cell_or_range()
        if (nzchar(ref)) {
          emit(TOKEN_TYPES$REF, paste0(sheet_pref, ref))
          next
        }
        # consume_sheet_prefix ate '!' but no ref followed — emit as OTHER
        emit(TOKEN_TYPES$OTHER, sheet_pref)
        next
      }

      # No sheet prefix — consume a plain identifier
      pos   <- start_pos
      ident <- consume_ident()

      # Structured table ref:  Table1[...]
      if (pos <= n && chars[pos] == "[") {
        body <- consume_bracketed()
        emit(TOKEN_TYPES$REF, paste0(ident, body))
        next
      }

      # Function call: identifier immediately followed by '('
      if (pos <= n && chars[pos] == "(") {
        emit(TOKEN_TYPES$FUNC, ident)
        next
      }

      # Cell-shape bare identifier (A1, B10). Detect a following ':<ref>'
      # as a range. Bare letter sequences are kept as IDENT (ambiguous
      # with LAMBDA parameters / named ranges).
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
    if (is_digit_dot(ch)) {
      start <- pos
      prev  <- ""
      while (pos <= n) {
        p <- chars[pos]
        if (grepl("[0-9.]", p, perl = TRUE)) {
          pos <- pos + 1L
          prev <- p
        } else if (p == "e" || p == "E") {
          pos <- pos + 1L
          prev <- p
        } else if ((p == "+" || p == "-") && (prev == "e" || prev == "E")) {
          pos <- pos + 1L
          prev <- p
        } else {
          break
        }
      }
      emit(TOKEN_TYPES$OTHER, substr_chars(start, pos - 1L))
      next
    }

    # ---- Everything else ----------------------------------------------
    pos <- pos + 1L
    emit(TOKEN_TYPES$OTHER, ch)
  }

  if (n_tokens < length(tokens)) length(tokens) <- n_tokens

  .merge_whitespace_refs(tokens)
}

# Post-lex merge for "REF (space)* : (space)* REF" → single REF token.
# OOXML stores `A1:B10` without whitespace, but users sometimes type
# `A1 : B10` in the formula bar. We collapse the whitespace so anchor /
# implicit-intersection / spill-anchor handlers see the range as one ref.
.merge_whitespace_refs <- function(tokens) {
  n <- length(tokens)
  if (n < 3L) return(tokens)

  is_space <- function(t) {
    t$type == TOKEN_TYPES$OTHER && grepl("^\\s+$", t$val, perl = TRUE)
  }

  out   <- vector("list", n)
  out_n <- 0L
  i     <- 1L
  while (i <= n) {
    if (tokens[[i]]$type != TOKEN_TYPES$REF) {
      out_n <- out_n + 1L
      out[[out_n]] <- tokens[[i]]
      i <- i + 1L
      next
    }
    # Look ahead past optional whitespace for ':' + optional whitespace + REF
    j <- i + 1L
    while (j <= n && is_space(tokens[[j]])) j <- j + 1L
    if (j > n || tokens[[j]]$type != TOKEN_TYPES$OTHER ||
        tokens[[j]]$val != ":") {
      out_n <- out_n + 1L
      out[[out_n]] <- tokens[[i]]
      i <- i + 1L
      next
    }
    k <- j + 1L
    while (k <= n && is_space(tokens[[k]])) k <- k + 1L
    if (k > n || tokens[[k]]$type != TOKEN_TYPES$REF) {
      out_n <- out_n + 1L
      out[[out_n]] <- tokens[[i]]
      i <- i + 1L
      next
    }
    # Merge tokens i..k into one REF "<lhs>:<rhs>"
    merged <- paste0(tokens[[i]]$val, ":", tokens[[k]]$val)
    out_n <- out_n + 1L
    out[[out_n]] <- list(type = TOKEN_TYPES$REF, val = merged)
    i <- k + 1L
  }
  if (out_n < length(out)) length(out) <- out_n
  out
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
