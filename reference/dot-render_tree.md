# Render a tree node and its children as lines of text.

Uses the tidyxl box-drawing style: ¦– child (not last) °– child (last) ¦
(continuation indent for non-last) (blank indent for last)

## Usage

``` r
.render_tree(nodes, prefix = "", is_root = TRUE)
```
