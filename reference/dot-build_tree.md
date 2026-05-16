# Build a nested tree from a flat token list.

Each node is a list(label, val, children). FUNC tokens own everything up
to and including their matching ')'.

## Usage

``` r
.build_tree(tokens, sep)
```
