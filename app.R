# =============================================================================
# Couche données : téléchargement DBnomics, cache, transformations
# =============================================================================

DBN_API   <- "https://api.db.nomics.world/v22/series/"
CACHE_DIR <- file.path(getwd(), "cache")
CACHE_TTL <- 12 * 3600   # secondes avant de retélécharger une série
DEMO_MODE <- identical(Sys.getenv("MACRO_DEMO"), "1")

dir.create(CACHE_DIR, showWarnings = FALSE)
.mem <- new.env()        # cache en mémoire pour la session R

# ---- Codes géographiques ----------------------------------------------------
expand_ids <- function(ids, geo) {
  eurostat_geo <- if (geo == "EA") c("EA21", "EA20", "EA") else geo
  ecb_geo      <- if (geo == "EA") "U2" else geo
  out <- character(0)
  for (id in ids) {
    if (grepl("{ecb}", id, fixed = TRUE)) {
      out <- c(out, gsub("{ecb}", ecb_geo, id, fixed = TRUE))
    } else if (grepl("{geo}", id, fixed = TRUE)) {
      out <- c(out, vapply(eurostat_geo, function(g) gsub("{geo}", g, id, fixed = TRUE), ""))
    } else {
      # Série sans marqueur pays : uniquement valable pour la France
      # (ex. INSEE) ou pour la zone euro (ex. taux BCE).
      if (geo == "FR" || grepl("ECB/FM/", id)) out <- c(out, id)
    }
  }
  unname(out)
}

# ---- Appel DBnomics ---------------------------------------------------------
fetch_dbnomics <- function(id) {
  url <- paste0(DBN_API, id, "?observations=1&format=json&metadata=false")
  res <- tryCatch(httr::GET(url, httr::timeout(25),
                            httr::user_agent("tableau-macro-france (Shiny)")),
                  error = function(e) NULL)
  if (is.null(res) || httr::status_code(res) != 200) return(NULL)
  js <- tryCatch(jsonlite::fromJSON(httr::content(res, "text", encoding = "UTF-8")),
                 error = function(e) NULL)
  docs <- js$series$docs
  if (is.null(docs) || NROW(docs) == 0) return(NULL)
  dates  <- unlist(docs$period_start_day[[1]])
  values <- suppressWarnings(as.numeric(unlist(docs$value[[1]])))
  if (length(dates) == 0 || length(dates) != length(values)) return(NULL)
  df <- data.frame(date = as.Date(dates), value = values)
  df <- df[!is.na(df$value), ]
  if (nrow(df) == 0) return(NULL)
  df
}

# Série de démonstration (données FICTIVES, pour tester l'app hors ligne)
fake_series <- function(meta, geo) {
  set.seed(sum(utf8ToInt(paste0(meta$code, geo))))
  by <- switch(meta$freq, A = "year", Q = "quarter", M = "month")
  dates <- seq(as.Date("2005-01-01"), as.Date("2026-07-01"), by = by)
  n <- length(dates)
  if (meta$kind == "level") {
    v <- 100 * cumprod(1 + rnorm(n, 0.004, 0.01))
  } else {
    v <- 2 + cumsum(rnorm(n, 0, 0.3)); v <- v - mean(v) + runif(1, -3, 8)
  }
  data.frame(date = dates, value = v)
}

disk_path <- function(id) file.path(CACHE_DIR, paste0(gsub("[^A-Za-z0-9_.-]", "_", id), ".rds"))

fetch_cached <- function(id) {
  if (exists(id, envir = .mem, inherits = FALSE)) return(get(id, envir = .mem))
  p <- disk_path(id)
  if (file.exists(p) && difftime(Sys.time(), file.mtime(p), units = "secs") < CACHE_TTL) {
    df <- readRDS(p)
  } else {
    df <- fetch_dbnomics(id)
    if (!is.null(df)) saveRDS(df, p)
    else if (file.exists(p)) df <- readRDS(p)   # hors ligne : ancienne copie
  }
  assign(id, df, envir = .mem)
  df
}

clear_cache <- function() {
  rm(list = ls(.mem), envir = .mem)
  unlink(list.files(CACHE_DIR, full.names = TRUE))
}

to_monthly_last <- function(df) {
  df$m <- as.Date(format(df$date, "%Y-%m-01"))
  df <- df[order(df$date), ]
  out <- df[!duplicated(df$m, fromLast = TRUE), c("m", "value")]
  names(out) <- c("date", "value")
  out
}

# ---- Récupération d'un indicateur -------------------------------------------
# Renvoie data.frame(date, value) avec les attributs "source" (id utilisé)
# et "tried" (ids essayés), ou NULL.
get_indicator <- function(code, geo = "FR") {
  key <- paste0("IND::", code, "::", geo)
  if (exists(key, envir = .mem, inherits = FALSE)) return(get(key, envir = .mem))

  if (code == "spread") {
    fr <- get_indicator("oat10", "FR"); de <- get_indicator("oat10", "DE")
    df <- NULL
    if (!is.null(fr) && !is.null(de)) {
      df <- merge(fr, de, by = "date", suffixes = c("_fr", "_de"))
      df <- data.frame(date = df$date, value = df$value_fr - df$value_de)
      attr(df, "source") <- paste(attr(fr, "source"), "−", attr(de, "source"))
    }
    assign(key, df, envir = .mem)
    return(df)
  }

  meta <- CATALOGUE[[code]]
  ids  <- expand_ids(meta$ids, geo)
  df <- NULL; used <- NA_character_
  if (DEMO_MODE) {
    if (length(ids)) { df <- fake_series(meta, geo); used <- "DÉMO (données fictives)" }
  } else {
    for (id in ids) {
      df <- fetch_cached(id)
      if (!is.null(df)) { used <- id; break }
    }
  }
  if (!is.null(df)) {
    if (identical(meta$agg, "last")) df <- to_monthly_last(df)
    df$value <- df$value * meta$scale
    df <- df[order(df$date), ]
    attr(df, "source") <- used
  }
  attr_tried <- ids
  if (!is.null(df)) attr(df, "tried") <- attr_tried
  assign(key, df, envir = .mem)
  df
}

ind_meta <- function(code) if (code %in% names(DERIVED)) DERIVED[[code]] else CATALOGUE[[code]]

# ---- Transformations --------------------------------------------------------
TRANSFORMS <- c("Niveau" = "level",
                "Variation sur la période" = "pop",
                "Variation sur un an" = "yoy",
                "Base 100 au début de la fenêtre" = "base100")

lag_n <- function(freq, how) {
  if (how == "pop") return(1L)
  switch(freq, A = 1L, Q = 4L, M = 12L, 1L)
}

# Appliquée sur toute l'historique (avant filtrage de dates) sauf base 100
transform_series <- function(df, how, meta, ma = 0) {
  if (is.null(df) || nrow(df) == 0) return(df)
  v <- df$value
  if (how %in% c("pop", "yoy")) {
    k <- lag_n(meta$freq, how)
    prev <- c(rep(NA, k), head(v, -k))
    v <- if (meta$kind == "rate") v - prev else 100 * (v / prev - 1)
  }
  if (ma > 1 && length(v) >= ma) v <- as.numeric(stats::filter(v, rep(1 / ma, ma), sides = 1))
  df$value <- v
  df[!is.na(df$value), ]
}

transform_unit <- function(meta, how) {
  switch(how,
    level   = meta$unit,
    pop     = if (meta$kind == "rate") "variation en points" else "% de variation",
    yoy     = if (meta$kind == "rate") "variation sur un an, en points" else "% sur un an",
    base100 = "base 100")
}

rebase <- function(df) {
  if (is.null(df) || nrow(df) == 0) return(df)
  b <- df$value[1]
  if (is.na(b) || b == 0) return(df)
  df$value <- 100 * df$value / b
  df
}

# ---- Format ----------------------------------------------------------------
fmt_period <- function(d, freq) {
  if (length(d) == 0 || is.na(d)) return("")
  y <- format(d, "%Y")
  switch(freq,
    A = y,
    Q = paste0("T", (as.integer(format(d, "%m")) - 1) %/% 3 + 1, " ", y),
    M = paste(c("janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août",
                "sept.", "oct.", "nov.", "déc.")[as.integer(format(d, "%m"))], y))
}

fmt_num <- function(x, digits = 1) {
  formatC(x, format = "f", digits = digits, big.mark = " ", decimal.mark = ",")
}
