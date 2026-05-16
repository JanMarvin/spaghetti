# Reconstruct a formula string from a token list

Separator tokens (`,`/`;`) are already present in the token stream and
have been swapped by the transform passes, so this is a simple concat.

## Usage

``` r
.detokenise(tokens, prefix_eq = TRUE)
```

## Arguments

- tokens:

  List of token objects.

- prefix_eq:

  Logical; prepend '=' if TRUE.

## Value

Character scalar.
