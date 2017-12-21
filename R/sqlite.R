#' Importer une table d'une base SQLite
#'
#' Importer une table d'une base SQLite
#'
#' @param table Nom de la table à importer.
#' @param base_sqlite Chemin vers la bse SQLite.
#'
#' @return Un tibble.
#'
#' @export
importer_table_sqlite <- function(table, base_sqlite, na_if = NULL) {

  connexion <- DBI::dbConnect(RSQLite::SQLite(), dbname = base_sqlite)

  table <- DBI::dbReadTable(connexion, table) %>%
    tibble::as_tibble()

  if (!is.null(na_if)) {
    table <- table %>%
      dplyr::mutate_if(is.character, ~ dplyr::na_if(., na_if))
  }

  DBI::dbDisconnect(connexion)
}

#' Exporter une table vers une base SQLite
#'
#' Exporter une table vers une base SQLite
#'
#' @param table Table à exporter
#' @param base_sqlite Chemin vers la bse SQLite.
#' @param nom_table Nom de la table à créer dans la base SQLite.
#' @param ecraser Ecraser une table du même nom.
#'
#' @export
exporter_table_sqlite <- function(table, base_sqlite, nom_table = NULL, ecraser = TRUE) {

  connexion <- DBI::dbConnect(RSQLite::SQLite(), dbname = base_sqlite)

  if (is.null(nom_table)) {
    nom_table <- deparse(substitute(table))
  }

  DBI::dbWriteTable(connexion, name = nom_table, value = table, row.names = FALSE, overwrite = ecraser)

  DBI::dbDisconnect(connexion)
}
