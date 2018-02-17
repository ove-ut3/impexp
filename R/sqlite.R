#' Creer une base SQLite vide
#'
#' Créer une base SQLite vide
#'
#' @param base_sqlite Chemin de la base SQLite.
#'
#' @export
sqlite_creer <- function(base_sqlite) {

  creation <- DBI::dbConnect(RSQLite::SQLite(), dbname = base_sqlite)

}

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
sqlite_importer <- function(table, base_sqlite) {

  if (!file.exists(base_sqlite)) {
    stop("La base SQLite \"", base_sqlite,"\" n'existe pas...", call. = FALSE)
  }

  connexion <- DBI::dbConnect(RSQLite::SQLite(), dbname = base_sqlite)

  if (!table %in% DBI::dbListTables(connexion)) {
    stop("La table \"", table,"\" n'existe pas...", call. = FALSE)
  }

  table <- DBI::dbReadTable(connexion, table) %>%
    dplyr::as_tibble()

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
#' @param message Message lors d'un écrasemetn de table.
#'
#' @export
sqlite_exporter <- function(table, base_sqlite, nom_table = NULL, ecraser = FALSE, message = TRUE) {

  if (!file.exists(base_sqlite)) {
    stop("La base SQLite \"", base_sqlite,"\" n'existe pas...", call. = FALSE)
  }

  connexion <- DBI::dbConnect(RSQLite::SQLite(), dbname = base_sqlite)

  if (is.null(nom_table)) {
    nom_table <- deparse(substitute(table))
  }

  if (nom_table %in% DBI::dbListTables(connexion) & ecraser == TRUE & message == TRUE) {
    message("La table \"", nom_table,"\" est écrasée...")
  }

  DBI::dbWriteTable(connexion, name = nom_table, value = table, row.names = FALSE, overwrite = ecraser)

  DBI::dbDisconnect(connexion)
}

#' Ajouter des lignes a une table SQLite
#'
#' Ajouter des lignes à une table SQLite
#'
#' @param table_ajout Table à concaténer.
#' @param table Table initiale.
#' @param base_sqlite Chemin de la base SQLite.
#' @param attente_verrou Tentatives répétées jusqu'à ce que le verrou soit libéré.
#'
#' @details
#' L'ordre des champs de la table à ajouter doit suivre celui de la table initiale
#'
#' @export
sqlite_ajouter_lignes <- function(table_ajout, table, base_sqlite, attente_verrou = TRUE) {

  if (ncol(impexp::sqlite_importer(table, base_sqlite)) != ncol(table_ajout)) {
    stop("Le nombre de colonnes de la table initiale et celle à concaténer doit être identique", call. = FALSE)
  }

  connexion <- DBI::dbConnect(RSQLite::SQLite(), dbname = base_sqlite)

  sql <- table_ajout %>%
    dplyr::mutate(id = row_number()) %>%
    tidyr::gather("champ", "valeur", -id) %>%
    dplyr::group_by(id) %>%
    dplyr::summarise(valeur = paste(valeur, collapse = "', '")) %>%
    dplyr::pull(valeur) %>%
    paste(collapse = "'), ('") %>%
    { paste0("INSERT INTO ", table, " VALUES ('", ., "');") }

  impexp::sqlite_executer_sql(sql, base_sqlite, attente_verrou)

  DBI::dbDisconnect(connexion)
}

#' Executer des commandes SQL dans une base SQLite
#'
#' Exécuter des commandes SQL dans une base SQLite
#'
#' @param liste_sql Un vecteur de commandes SQL au format chaîne de caractères.
#' @param base_sqlite Chemin de la base SQLite.
#' @param attente_verrou Tentatives répétées jusqu'à ce que le verrou soit libéré.
#'
#' @export
sqlite_executer_sql <- function(liste_sql, base_sqlite, attente_verrou = TRUE) {

  if (!file.exists(base_sqlite)) {
    stop("La base SQLite \"", base_sqlite,"\" n'existe pas...", call. = FALSE)
  }

  connexion <- DBI::dbConnect(RSQLite::SQLite(), dbname = base_sqlite)

  if (attente_verrou == FALSE) {
    purrr::walk(liste_sql, ~ DBI::dbExecute(connexion, .))

  } else {
    execute_sql_safe <- purrr::safely(DBI::dbExecute)

    execute_sql <- purrr::map(liste_sql, ~ execute_sql_safe(connexion, .))

    temoin_erreur <- purrr::map_lgl(execute_sql, ~ !is.null(.$error))

    while(any(temoin_erreur)) {

      liste_sql <- liste_sql[temoin_erreur]

      execute_sql <- purrr::map(liste_sql, ~ execute_sql_safe(connexion, .))

      temoin_erreur <- purrr::map_lgl(execute_sql, ~ !is.null(.$error))
    }
  }

  DBI::dbDisconnect(connexion)
}
