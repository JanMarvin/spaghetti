# spaghetti: Bidirectional Spreadsheet-Formula to OOXML Translator

Translates spreadsheet formulas between the user-facing format (as
displayed to the end user) and the OOXML storage format (as found in
`.xlsx` XML source).

### Key transformations

|  |  |  |
|----|----|----|
| User-facing | OOXML storage | Rule |
| `SEQUENCE(10)` | `_xlfn.SEQUENCE(10)` | Future function prefix |
| `FILTER(A1:A10,...)` | `_xlfn._xlws.FILTER(A1:A10,...)` | Web-service namespace |
| `LAMBDA(x, x*2)` | `_xlfn.LAMBDA(_xlpm.x, _xlpm.x*2)` | Lambda param prefix |
| `A1#` | `_xlfn.ANCHORARRAY(A1)` | Spilled range operator |
| `@A1:A10` | `_xlfn.SINGLE(A1:A10)` | Implicit intersection |

### Main functions

- [`to_xml()`](https://janmarvin.github.io/spaghetti/reference/to_xml.md):
  user-facing form → OOXML storage

- [`from_xml()`](https://janmarvin.github.io/spaghetti/reference/from_xml.md):
  OOXML storage → user-facing form

- [`function_prefix()`](https://janmarvin.github.io/spaghetti/reference/function_prefix.md):
  inspect prefix tier for any function name

- [`supported_locales()`](https://janmarvin.github.io/spaghetti/reference/supported_locales.md):
  list available locale codes

- [`round_trip()`](https://janmarvin.github.io/spaghetti/reference/round_trip.md):
  convert and back-convert for testing

### Localisation

Non-English locales ship with translated function names. Pass
`locale = "de"` (German), `"fr"` (French), `"es"` (Spanish), etc. to
[`to_xml()`](https://janmarvin.github.io/spaghetti/reference/to_xml.md)
to translate local names to English before prefixing, or to
[`from_xml()`](https://janmarvin.github.io/spaghetti/reference/from_xml.md)
to output localised names.

Supported locales: af, sq, am-ET, ar, hy, as-IN, az-Latn, bn, eu, be,
bs-Cyrl, bs-Latn, bg, my, ca, ckb, chr-Cher, zh-Hans, zh-Hant, hr, cs,
da, prs-AF, nl, et, fil-PH, fi, fr, ff, gl, ka, de, el, gu, ha-Latn-NG,
he, hi, hu, is, ig-NG, id, iu-Latn, ga-IE, xh-ZA, zu-ZA, it, ja, qut-GT,
kn, kk, km-KH, rw-RW, sw, kok, ko, ky, lo, lv, lt, lb-LU, mk, ms, ml-IN,
mt-MT, mi-NZ, mr, mn, ne-NP, nb-NO, nn-NO, or-IN, ps-AF, fa, pl, pt-BR,
pt-PT, pa-Arab, pa-Guru, quz, ro, ru, gd, sr-Cyrl, sr-Latn, nso-ZA,
tn-ZA, sd, si-LK, sk, sl, es, sv, tg-Cyrl-TJ, ta, tt-Cyrl, te, th, ti,
tr, tk-TM, uk, ur, ug, uz-Cyrl, uz-Latn, ca-ES-valencia, vi, guc, cy-GB,
wo-SN, yo-NG.

## References

- ECMA-376 Part 1 §18.17: Built-in function list

- XlsxWriter:
  <https://xlsxwriter.readthedocs.io/working_with_formulas.html>

- EPPlus wiki:
  <https://github.com/EPPlusSoftware/EPPlus/wiki/Function-prefixes>

- libxlsxwriter LAMBDA:
  <https://libxlsxwriter.github.io/lambda_8c-example.html>

## Author

**Maintainer**: Jan Marvin Garbuszus <jan.garbuszus@ruhr-uni-bochum.de>

Authors:

- Jan Marvin Garbuszus <jan.garbuszus@ruhr-uni-bochum.de>
