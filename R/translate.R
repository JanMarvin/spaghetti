# R/translate.R
# Core translation functions.

#' Convert a user-facing Excel formula to OOXML storage format
#' @param formula a formula
#' @param locale a locale
#' @param warn_unknown a logical
#' @export
to_xml <- function(formula, locale = NULL, warn_unknown = TRUE) {
  stopifnot(is.character(formula), length(formula) == 1)
  if (is.null(formula) || is.na(formula) || formula == "") return(formula)

  tokens <- .tokenise(formula)
  out    <- .transform_to_xml(tokens, locale = locale, warn_unknown = warn_unknown)
  .detokenise(out, prefix_eq = TRUE)
}

#' Convert a OOXML formula to a user-facing Excel formula
#' @param formula a formula
#' @param locale a locale
#' @export
from_xml <- function(formula, locale = NULL) {
  stopifnot(is.character(formula), length(formula) == 1)
  if (is.null(formula) || is.na(formula) || formula == "") return(formula)

  tokens <- .tokenise(formula)
  out    <- .transform_from_xml(tokens, locale = locale)
  .detokenise(out, prefix_eq = TRUE)
}

# ── Internal transformation: Excel → OOXML ──────────────────────────────────

#' @keywords internal
.transform_to_xml <- function(tokens, locale, warn_unknown) {
  n   <- length(tokens)
  out <- list()
  i   <- 1L

  # lambda_scope: list(list(depth = int, params = char_vec))
  lambda_scope <- list()

  while (i <= n) {
    tok <- tokens[[i]]

    # ── String literals ────────────────────────────────────────────────
    if (tok$type == TOKEN_TYPES$STRING) {
      out <- c(out, list(tok))
      i   <- i + 1L
      next
    }

    # ── Implicit intersection @ → _xlfn.SINGLE( ... ) ──────────────────
    if (tok$type == TOKEN_TYPES$IMPLICIT) {
      ref_tokens <- list()
      j <- i + 1L
      while (j <= n) {
        t <- tokens[[j]]
        if (t$type %in% c(TOKEN_TYPES$IDENT, TOKEN_TYPES$REF)) {
          ref_tokens <- c(ref_tokens, list(t))
          j <- j + 1L
          if (j <= n && tokens[[j]]$val == ":") {
            ref_tokens <- c(ref_tokens, list(tokens[[j]]))
            j <- j + 1L
          } else { break }
        } else { break }
      }
      ref_str <- paste(vapply(ref_tokens, `[[`, character(1), "val"), collapse = "")
      out <- c(out, list(
        list(type = TOKEN_TYPES$OTHER, val = "_xlfn.SINGLE("),
        list(type = TOKEN_TYPES$OTHER, val = ref_str),
        list(type = TOKEN_TYPES$OTHER, val = ")")
      ))
      i <- j
      next
    }

    # ── FUNC token ─────────────────────────────────────────────────────
    if (tok$type == TOKEN_TYPES$FUNC) {
      fn_raw   <- .strip_prefix(tok$val)
      fn_en    <- toupper(.locale_to_english(fn_raw, locale))
      tier     <- .prefix_for(fn_en)

      prefixed <- switch(tier,
                         xlws   = paste0("_xlfn._xlws.", fn_en),
                         xlfn   = paste0("_xlfn.",       fn_en),
                         legacy = fn_en
      )

      out <- c(out, list(list(type = TOKEN_TYPES$FUNC, val = prefixed)))

      if (fn_en == "LAMBDA") {
        # depth 0 will become 1 when the '(' is processed in OTHER
        lambda_scope <- c(lambda_scope, list(list(depth = 0L, params = character(0))))
      }
      i <- i + 1L
      next
    }

    # ── IDENT (Parameters & Variables) ──────────────────────────────────
    if (tok$type == TOKEN_TYPES$IDENT) {
      raw_name <- toupper(.strip_prefix(tok$val))

      # 1. Register parameter if at depth 1 of a LAMBDA
      if (length(lambda_scope) > 0 && lambda_scope[[length(lambda_scope)]]$depth == 1L) {
        lambda_scope[[length(lambda_scope)]]$params <-
          unique(c(lambda_scope[[length(lambda_scope)]]$params, raw_name))
      }

      # 2. Check if this is a usage or definition of a registered param
      is_param <- any(vapply(lambda_scope, function(s) raw_name %in% s$params, logical(1)))

      if (is_param) {
        # SHIELD from localization: use raw_name directly
        val_out <- paste0("_xlpm.", raw_name)
      } else {
        # Regular identifier: apply localization
        val_out <- toupper(.locale_to_english(raw_name, locale))
      }

      out <- c(out, list(list(type = TOKEN_TYPES$IDENT, val = val_out)))
      i <- i + 1L
      next
    }

    # ── Anchor # → _xlfn.ANCHORARRAY(prev_ref) ──────────────────────────
    if (tok$type == TOKEN_TYPES$ANCHOR) {
      if (length(out) > 0) {
        last_tok <- out[[length(out)]]
        last_val <- last_tok$val
        m <- regmatches(last_val, regexpr("[$A-Za-z]+[$0-9]+$", last_val))
        if (length(m) == 1 && nchar(m) > 0) {
          stripped <- substring(last_val, 1, nchar(last_val) - nchar(m))
          if (nchar(stripped) > 0) out[[length(out)]]$val <- stripped else out <- out[-length(out)]
          out <- c(out, list(list(type = TOKEN_TYPES$OTHER, val = paste0("_xlfn.ANCHORARRAY(", m, ")"))))
        } else { out <- c(out, list(tok)) }
      } else { out <- c(out, list(tok)) }
      i <- i + 1L
      next
    }

    # ── OTHER (Parens, Operators, Commas) ──────────────────────────────
    if (tok$type == TOKEN_TYPES$OTHER) {
      val <- tok$val
      if (val == "(") {
        lambda_scope <- .lambda_scope_open(lambda_scope)
      } else if (val == ")") {
        lambda_scope <- .lambda_scope_close(lambda_scope)
      } else if (val == ",") {
        lambda_scope <- .lambda_scope_comma(lambda_scope)
      }
    }

    out <- c(out, list(tok))
    i   <- i + 1L
  }
  out
}

# ── Internal transformation: OOXML → Excel ───────────────────────────────────

#' @keywords internal
.transform_from_xml <- function(tokens, locale) {
  n <- length(tokens); out <- list(); i <- 1L
  while (i <= n) {
    tok <- tokens[[i]]
    if (tok$type == TOKEN_TYPES$STRING) { out <- c(out, list(tok)); i <- i + 1L; next }
    if (tok$type == TOKEN_TYPES$FUNC) {
      fn_raw   <- tok$val
      fn_clean <- .strip_prefix(fn_raw)
      fn_out   <- if (!is.null(locale)) .english_to_locale(fn_clean, locale) else fn_clean
      if (toupper(fn_clean) == "ANCHORARRAY") {
        j <- i + 1L
        if (j <= n && tokens[[j]]$val == "(") {
          j <- j + 1L; ref_parts <- list()
          while (j <= n && tokens[[j]]$val != ")") { ref_parts <- c(ref_parts, list(tokens[[j]])); j <- j + 1L }
          if (j <= n) j <- j + 1L
          ref_str <- paste(vapply(ref_parts, `[[`, character(1), "val"), collapse = "")
          out <- c(out, list(list(type = TOKEN_TYPES$OTHER, val = paste0(ref_str, "#"))))
          i <- j; next
        }
      }
      if (toupper(fn_clean) == "SINGLE") {
        j <- i + 1L
        if (j <= n && tokens[[j]]$val == "(") {
          j <- j + 1L; ref_parts <- list()
          while (j <= n && tokens[[j]]$val != ")") { ref_parts <- c(ref_parts, list(tokens[[j]])); j <- j + 1L }
          if (j <= n) j <- j + 1L
          ref_str <- paste(vapply(ref_parts, `[[`, character(1), "val"), collapse = "")
          out <- c(out, list(list(type = TOKEN_TYPES$OTHER, val = paste0("@", ref_str))))
          i <- j; next
        }
      }
      out <- c(out, list(list(type = TOKEN_TYPES$FUNC, val = fn_out))); i <- i + 1L; next
    }
    if (tok$type == TOKEN_TYPES$IDENT) {
      fn_clean <- .strip_prefix(tok$val)
      fn_out   <- if (!is.null(locale)) .english_to_locale(fn_clean, locale) else fn_clean
      out <- c(out, list(list(type = TOKEN_TYPES$IDENT, val = fn_out))); i <- i + 1L; next
    }
    out <- c(out, list(tok)); i <- i + 1L
  }
  out
}

# ── Private Helpers ──────────────────────────────────────────────────

.lambda_scope_open <- function(scope) {
  if (length(scope) == 0) return(scope)
  idx <- length(scope)
  scope[[idx]]$depth <- scope[[idx]]$depth + 1L
  scope
}

.lambda_scope_close <- function(scope) {
  if (length(scope) == 0) return(scope)
  idx <- length(scope)
  scope[[idx]]$depth <- scope[[idx]]$depth - 1L
  if (scope[[idx]]$depth == 0L) scope <- scope[-idx]
  scope
}

.lambda_scope_comma <- function(scope) {
  # Required to exist; comma tracking is implicit via IDENT registration at depth 1.
  scope
}
