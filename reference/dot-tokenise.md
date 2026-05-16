# Tokenise a formula string

Tokenise a formula string

## Usage

``` r
.tokenise(formula, sep = ",")
```

## Arguments

- formula:

  Character scalar, optionally starting with '='.

- sep:

  The argument separator (e.g., ',' or ';').

## Value

A list of token objects, each with fields `type` and `val`.
