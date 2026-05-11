# R/xlex.R
#
# xlex(): tokenise a formula and display it as an ASCII tree, styled after
# tidyxl::xlex(). Useful for understanding how spaghetti parses a formula
# and for debugging translation issues.
#
# The tree nesting follows the call structure of the source formula:
#   - Each FUNC token becomes a parent node
#   - Tokens between its '(' and matching ')' are its children
#   - Nested function calls produce nested subtrees

# ── Token classification ─────────────────────────────────────────────────────
# The base lexer emits a coarse set of types. xlex() refines OTHER tokens
# into more descriptive labels for display.

.xlex_label <- function(tok, sep) {
  type <- tok$type
  val  <- tok$val

  if (type == TOKEN_TYPES$FUNC)     return("function")
  if (type == TOKEN_TYPES$STRING)   return("text")
  if (type == TOKEN_TYPES$ANCHOR)   return("operator")   # # spill
  if (type == TOKEN_TYPES$IMPLICIT) return("operator")   # @ implicit

  if (type == TOKEN_TYPES$IDENT) {
    # Named range or LAMBDA/LET parameter — treat as ref
    return("ref")
  }

  if (type == TOKEN_TYPES$OTHER) {
    if (val == "(")                       return("fun_open")
    if (val == ")")                       return("fun_close")
    if (val %in% c(",", ";", sep))        return("separator")
    if (grepl("^[0-9]+(\\.[0-9]*)?([eE][+-]?[0-9]+)?$", val))
      return("number")
    if (val == " ")                       return("operator")   # whitespace op
    if (grepl("^[+*/^&=<>:-]+$", val))   return("operator")
    if (val == "{" || val == "}")         return("operator")   # array literal
    # Cell refs come through as IDENT or as runs parsed by OTHER
    return("other")
  }

  tolower(type)
}

# ── Tree builder ─────────────────────────────────────────────────────────────

#' Build a nested tree from a flat token list.
#'
#' Each node is a list(label, val, children).
#' FUNC tokens own everything up to and including their matching ')'.
#' @keywords internal
.build_tree <- function(tokens, sep) {
  # We'll use a recursive descent over the flat token vector.
  # Returns list(nodes = <tree list>, consumed = <int>)
  .parse_level <- function(toks, start, stop_at_close = FALSE) {
    nodes <- list()
    i     <- start

    while (i <= length(toks)) {
      tok   <- toks[[i]]
      label <- .xlex_label(tok, sep)

      if (label == "fun_close" && stop_at_close) {
        # Return without consuming the ')' — caller will add it
        return(list(nodes = nodes, next_i = i))
      }

      if (label == "function") {
        # Consume the function name, then descend into its argument list
        fn_node <- list(label = "function", val = tok$val, children = list())
        i <- i + 1L

        # Expect '(' next
        if (i <= length(toks) && .xlex_label(toks[[i]], sep) == "fun_open") {
          open_node <- list(label = "fun_open", val = toks[[i]]$val,
                            children = list())
          i <- i + 1L

          # Recurse for the argument tokens
          inner <- .parse_level(toks, i, stop_at_close = TRUE)
          i     <- inner$next_i

          # Consume the matching ')'
          close_node <- NULL
          if (i <= length(toks) && .xlex_label(toks[[i]], sep) == "fun_close") {
            close_node <- list(label = "fun_close", val = toks[[i]]$val,
                               children = list())
            i <- i + 1L
          }

          open_node$children <- c(inner$nodes,
                                  if (!is.null(close_node)) list(close_node))
          fn_node$children   <- list(open_node)
        }

        nodes <- c(nodes, list(fn_node))
        next
      }

      # Leaf node
      node  <- list(label = label, val = tok$val, children = list())
      nodes <- c(nodes, list(node))
      i     <- i + 1L
    }

    list(nodes = nodes, next_i = i)
  }

  result <- .parse_level(tokens, 1L, stop_at_close = FALSE)
  result$nodes
}

# ── ASCII renderer ────────────────────────────────────────────────────────────

#' Render a tree node and its children as lines of text.
#'
#' Uses the tidyxl box-drawing style:
#'   ¦-- child (not last)
#'   °-- child (last)
#'   ¦   (continuation indent for non-last)
#'       (blank indent for last)
#'
#' @keywords internal
.render_tree <- function(nodes, prefix = "", is_root = TRUE) {
  lines <- character(0)

  for (k in seq_along(nodes)) {
    node    <- nodes[[k]]
    is_last <- k == length(nodes)

    connector <- if (is_last) "\u00b0--" else "\u00a6--"
    cont      <- if (is_last) "   "      else "\u00a6   "

    # Val display: strip leading/trailing whitespace for display but keep
    # the label right-aligned in a fixed column (tidyxl uses ~18 char pad)
    val_disp  <- node$val

    line <- paste0(prefix, connector, " ", val_disp)
    lines <- c(lines, line)

    if (length(node$children) > 0) {
      child_lines <- .render_tree(node$children,
                                  prefix  = paste0(prefix, cont),
                                  is_root = FALSE)
      lines <- c(lines, child_lines)
    }
  }

  lines
}

# ── Public API ────────────────────────────────────────────────────────────────

#' Tokenise and display a formula as an ASCII tree
#'
#' Parses a formula using spaghetti's lexer and renders the token tree in
#' the style of `tidyxl::xlex()`. Nesting follows the call structure: tokens
#' inside a function's parentheses appear as children of that function node.
#'
#' The formula is displayed as-is (no OOXML translation). Pass the result of
#' `to_xml()` if you want to inspect the storage form.
#'
#' @param formula Character scalar. Formula with or without leading `=`.
#' @param locale  Two-letter locale code or NULL. Used to select the correct
#'                argument separator (`;` for German etc.) and to translate
#'                localised function names in the display label.
#' @param print   Logical. Print the tree to the console (default TRUE).
#'                Set FALSE to get the data frame silently.
#'
#' @return A data frame with columns `depth`, `val`, `label`, invisibly.
#'   `depth` is the nesting level (0 = root, 1 = top-level arguments, …).
#'   Printed as a side-effect when `print = TRUE`.
#'
#' @export
#' @examples
#' xlex("=SUM(A1:A10)")
#' xlex("=IF(A1>0, VLOOKUP(A1, B:C, 2, 0), NA())")
#' xlex("=LAMBDA(x, x * 2)(5)")
#' xlex("=SUMMEWENNS(C2:C10; A2:A10; \"Berlin\")", locale = "de")
#' # Inspect OOXML form
#' xlex(to_xml("=FILTER(A1:A10, B1:B10 > 0)"))
xlex <- function(formula, locale = NULL, print = TRUE) {
  stopifnot(is.character(formula), length(formula) == 1L)
  if (is.na(formula) || !nzchar(trimws(formula))) {
    if (print) message("<empty formula>")
    return(invisible(data.frame(depth = integer(0), val = character(0),
                                label = character(0),
                                stringsAsFactors = FALSE)))
  }

  sep    <- .get_sep(locale)
  tokens <- .tokenise(formula, sep = sep)
  tree   <- .build_tree(tokens, sep = sep)

  # ── Flatten tree to data frame ───────────────────────────────────────────
  rows <- list()
  .flatten <- function(nodes, depth) {
    for (nd in nodes) {
      rows[[length(rows) + 1L]] <<- data.frame(
        depth = depth, val = nd$val, label = nd$label,
        stringsAsFactors = FALSE
      )
      if (length(nd$children) > 0L)
        .flatten(nd$children, depth + 1L)
    }
  }
  .flatten(tree, 0L)
  df <- do.call(rbind, rows)
  if (is.null(df)) df <- data.frame(depth = integer(0), val = character(0),
                                    label = character(0),
                                    stringsAsFactors = FALSE)

  # ── Print ────────────────────────────────────────────────────────────────
  if (print) {
    # Header row (mirrors tidyxl style)
    val_w <- max(nchar(df$val), 3L) + 2L

    # Root line
    f_display <- if (startsWith(formula, "=")) formula else paste0("=", formula)
    cat("root", "\n")

    lines  <- .render_tree(tree)
    labels <- .collect_labels(tree)

    # Pad val column for alignment
    pad <- max(vapply(lines, nchar, integer(1)), 1L) + 2L
    for (k in seq_along(lines)) {
      cat(formatC(lines[k], width = -pad, flag = "-"), labels[k], "\n")
    }
  }

  invisible(df)
}

#' Collect labels in the same DFS order as .render_tree produces lines.
#' @keywords internal
.collect_labels <- function(nodes) {
  labels <- character(0)
  for (nd in nodes) {
    labels <- c(labels, nd$label)
    if (length(nd$children) > 0L)
      labels <- c(labels, .collect_labels(nd$children))
  }
  labels
}
