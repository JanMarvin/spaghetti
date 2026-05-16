# Tokenise and display a formula as an ASCII tree

Parses a formula using spaghetti's lexer and renders the token tree in
the style of `tidyxl::xlex()`. Nesting follows the call structure:
tokens inside a function's parentheses appear as children of that
function node.

## Usage

``` r
xlex(formula, locale = NULL, print = TRUE)
```

## Arguments

- formula:

  Character scalar. Formula with or without leading `=`.

- locale:

  Two-letter locale code or NULL. Used to select the correct argument
  separator (`;` for German etc.) and to translate localised function
  names in the display label.

- print:

  Logical. Print the tree to the console (default TRUE). Set FALSE to
  get the data frame silently.

## Value

A data frame with columns `depth`, `val`, `label`, invisibly. `depth` is
the nesting level (0 = root, 1 = top-level arguments, …). Printed as a
side-effect when `print = TRUE`.

## Details

The formula is displayed as-is (no OOXML translation). Pass the result
of
[`to_xml()`](https://janmarvin.github.io/spaghetti/reference/to_xml.md)
if you want to inspect the storage form.

## Examples

``` r
xlex("=SUM(A1:A10)")
#> root 
#> °-- SUM            function 
#>    °-- (           fun_open 
#>       ¦-- A1:A10   ref 
#>       °-- )        fun_close 
xlex("=IF(A1>0, VLOOKUP(A1, B:C, 2, 0), NA())")
#> root 
#> °-- IF                function 
#>    °-- (              fun_open 
#>       ¦-- A1          ref 
#>       ¦-- >           operator 
#>       ¦-- 0           number 
#>       ¦-- ,           separator 
#>       ¦--             operator 
#>       ¦-- VLOOKUP     function 
#>       ¦   °-- (       fun_open 
#>       ¦      ¦-- A1   ref 
#>       ¦      ¦-- ,    separator 
#>       ¦      ¦--      operator 
#>       ¦      ¦-- B    ref 
#>       ¦      ¦-- :    operator 
#>       ¦      ¦-- C    ref 
#>       ¦      ¦-- ,    separator 
#>       ¦      ¦--      operator 
#>       ¦      ¦-- 2    number 
#>       ¦      ¦-- ,    separator 
#>       ¦      ¦--      operator 
#>       ¦      ¦-- 0    number 
#>       ¦      °-- )    fun_close 
#>       ¦-- ,           separator 
#>       ¦--             operator 
#>       ¦-- NA          function 
#>       ¦   °-- (       fun_open 
#>       ¦      °-- )    fun_close 
#>       °-- )           fun_close 
xlex("=LAMBDA(x, x * 2)(5)")
#> root 
#> ¦-- LAMBDA     function 
#> ¦   °-- (      fun_open 
#> ¦      ¦-- x   ref 
#> ¦      ¦-- ,   separator 
#> ¦      ¦--     operator 
#> ¦      ¦-- x   ref 
#> ¦      ¦--     operator 
#> ¦      ¦-- *   operator 
#> ¦      ¦--     operator 
#> ¦      ¦-- 2   number 
#> ¦      °-- )   fun_close 
#> ¦-- (          fun_open 
#> ¦-- 5          number 
#> °-- )          fun_close 
xlex("=SUMMEWENNS(C2:C10; A2:A10; \"Berlin\")", locale = "de")
#> root 
#> °-- SUMMEWENNS       function 
#>    °-- (             fun_open 
#>       ¦-- C2:C10     ref 
#>       ¦-- ;          separator 
#>       ¦--            operator 
#>       ¦-- A2:A10     ref 
#>       ¦-- ;          separator 
#>       ¦--            operator 
#>       ¦-- "Berlin"   text 
#>       °-- )          fun_close 
# Inspect OOXML form
xlex(to_xml("=FILTER(A1:A10, B1:B10 > 0)"))
#> root 
#> °-- _xlfn._xlws.FILTER   function 
#>    °-- (                 fun_open 
#>       ¦-- A1:A10         ref 
#>       ¦-- ,              separator 
#>       ¦--                operator 
#>       ¦-- B1:B10         ref 
#>       ¦--                operator 
#>       ¦-- >              operator 
#>       ¦--                operator 
#>       ¦-- 0              number 
#>       °-- )              fun_close 
```
