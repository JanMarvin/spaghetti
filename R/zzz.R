# R/zzz.R
# Package initialisation.
#
# R sources files in R/ in ASCII order (locale-independent under
# R CMD INSTALL). Top-level assignments in aaa.R and registry.R run at
# install/load time before .onLoad() is invoked.
#
# This package does not ship Microsoft Terminology data. Users invoke
# setup_terminology() to download and cache it; .onLoad() picks up the
# cache if it exists, otherwise leaves locale tables empty.

.onLoad <- function(libname, pkgname) {
  # Always start with empty locale state. setup_terminology() / the cache
  # loader will populate if data is available.
  .reset_terminology_state()

  # Try to load an existing cache (silent — first-time users with no
  # cache will just see empty locale support).
  .load_terminology_cache()

  invisible(NULL)
}
