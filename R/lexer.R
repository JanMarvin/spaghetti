# R/lexer.R
# Lexical analyser for Excel formula strings.
#
# Excel formula syntax is not regular: string literals can contain
# function-name-like tokens, and the @ and # operators are contextual.
# A simple regex gsub over the whole formula will corrupt string content.
#
# This tokeniser walks the formula character-by-character, identifying:
#   - STRING   : "..." literals (with "" escape sequences)
#   - FUNC     : identifier immediately followed by '('
#   - IDENT    : identifier not followed by '('  (named range, LAMBDA param)
#   - REF      : cell reference like A1, $B$2, A1:B10, Sheet1!A1
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

#' Tokenise an Excel formula string
#'
#' @param formula Character scalar, optionally starting with '='.
#' @param sep The argument separator (e.g., ',' or ';').
#' @return A list of token objects, each with fields `type` and `val`.
#' @keywords internal
.tokenise <- function(formula, sep = ",") {
  # Strip leading '='
  if (startsWith(formula, "=")) formula <- substring(formula, 2)

  chars <- strsplit(formula, "")[[1]]
  n     <- length(chars)
  pos   <- 1L
  tokens <- list()

  peek <- function(offset = 0L) {
    p <- pos + offset
    if (p > n) return("")
    chars[p]
  }

  advance <- function() {
    c <- chars[pos]
    pos <<- pos + 1L
    c
  }

  emit <- function(type, val) {
    tokens[[length(tokens) + 1L]] <<- list(type = type, val = val)
  }

  while (pos <= n) {
    ch <- peek()

    # ---- String literal ------------------------------------------------
    if (ch == '"') {
      advance()  # consume opening quote
      s <- '"'
      repeat {
        if (pos > n) break
        c <- advance()
        if (c == '"') {
          if (peek() == '"') {      # escaped ""
            s <- paste0(s, '""')
            advance()
          } else {                  # closing quote
            s <- paste0(s, '"')
            break
          }
        } else {
          s <- paste0(s, c)
        }
      }
      emit(TOKEN_TYPES$STRING, s)
      next
    }

    # ---- Implicit intersection operator --------------------------------
    if (ch == "@") {
      advance()
      emit(TOKEN_TYPES$IMPLICIT, "@")
      next
    }

    # ---- Spill/anchor operator -----------------------------------------
    if (ch == "#") {
      advance()
      emit(TOKEN_TYPES$ANCHOR, "#")
      next
    }

    # ---- Argument Separator (Dynamic) ----------------------------------
    if (ch == sep) {
      emit(TOKEN_TYPES$OTHER, advance())
      next
    }

    # ---- Identifier (function name, named range, LAMBDA parameter) -----
    if (grepl("[A-Za-z_]", ch) || (ch == "_")) {
      ident <- ""
      while (pos <= n && grepl("[A-Za-z0-9_.!$]", peek())) {
        ident <- paste0(ident, advance())
        if (endsWith(ident, "!")) break
      }

      if (peek() == "(") {
        emit(TOKEN_TYPES$FUNC, ident)
      } else {
        emit(TOKEN_TYPES$IDENT, ident)
      }
      next
    }

    # ---- Numbers -------------------------------------------------------
    if (grepl("[0-9.]", ch)) {
      num <- ""
      while (pos <= n && grepl("[0-9.eE+\\-]", peek())) {
        num <- paste0(num, advance())
      }
      emit(TOKEN_TYPES$OTHER, num)
      next
    }

    # ---- Everything else ------------------------------------------------
    emit(TOKEN_TYPES$OTHER, advance())
  }

  tokens
}

#' Reconstruct a formula string from a token list
#'
#' @param tokens List of token objects.
#' @param prefix_eq Logical; prepend '=' if TRUE.
#' @param sep The separator used to join arguments (not strictly needed if
#'   tokens already contain the separator, but useful for validation).
#' @return Character scalar.
#' @keywords internal
.detokenise <- function(tokens, prefix_eq = TRUE, sep = ",") {
  # If your transformation logic swaps the comma/semicolon tokens
  # inside the list, this simple collapse works perfectly.
  parts <- vapply(tokens, function(t) t$val, character(1))
  result <- paste(parts, collapse = "")

  if (prefix_eq) paste0("=", result) else result
}
