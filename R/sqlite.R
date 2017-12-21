#' Importer une table d'une base SQLite
#'
#' Importer une table d'une base SQLite
#'
#' @param table Nom de la table à importer.
#' @param base_sqlite Chemin de la base SQLite.
#'
#' @return Un tibble.
#'
#' @export
importer_table_sqlite <- function(table, base_sqlite) {

  if (!file.exists(base_sqlite)) {
    stop("La base SQLite \"", base_sqlite,"\" n'existe pas...", call. = FALSE)
  }

  connexion <- DBI::dbConnect(RSQLite::SQLite(), dbname = base_sqlite)

  if (!table %in% DBI::dbListTables(connexion)) {
    stop("La table \"", table,"\" n'existe pas...", call. = FALSE)
  }

  table <- DBI::dbReadTable(connexion, table) %>%
    tibble::as_tibble()

  DBI::dbDisconnect(connexion)

  return(table)
}

#' Exporter une table vers une base SQLite
#'
#' Exporter une table vers une base SQLite
#'
#' @param table Table à exporter
#' @param base_sqlite Chemin de la base SQLite.
#' @param nom_table Nom de la table à créer dans la base SQLite.
#' @param ecraser Ecraser une table du même nom.
#'
#' @export
exporter_table_sqlite <- function(table, base_sqlite, nom_table = NULL, ecraser = TRUE) {

  if (!file.exists(base_sqlite)) {
    stop("La base SQLite \"", base_sqlite,"\" n'existe pas...", call. = FALSE)
  }

  connexion <- DBI::dbConnect(RSQLite::SQLite(), dbname = base_sqlite)

  if (is.null(nom_table)) {
    nom_table <- deparse(substitute(table))
  }

  if (nom_table %in% DBI::dbListTables(connexion)) {
    message("La table \"", nom_table,"\" est écrasée...")
  }

  DBI::dbWriteTable(connexion, name = nom_table, value = table, row.names = FALSE, overwrite = ecraser)

  DBI::dbDisconnect(connexion)
}
