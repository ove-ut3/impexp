#' Se connecter a une base Access
#'
#' Se connecter à une base Access.
#'
#' @param base_access Chemin de la base Access.
#'
#' @return Une connexion à une base Access.
#'
#' @export
#' @keywords internal
access_connexion <- function(base_access = "Tables_ref.accdb") {

  if(!stringr::str_detect(base_access, "[A-Z]:\\/")) {
    dbq <- paste0(getwd(), "/", base_access)
  } else {
    dbq <- base_access
  }

  connexion <- DBI::dbConnect(odbc::odbc(),
                              driver = "Microsoft Access Driver (*.mdb, *.accdb)",
                              dbq = iconv(dbq, to = "Windows-1252"),
                              encoding = "Windows-1252")
  return(connexion)
}

#' Lister les tables d'une base Access
#'
#' Lister les tables d'une base Access.
#'
#' @param base_access Chemin de la base Access.
#'
#' @return un vecteur de type caractère contenant les noms des tables de la base Access.
#'
#' @examples
#' impexp::access_liste_tables(paste0(racine_packages, "impexp/inst/extdata/impexp.accdb"))
#'
#' @export
access_liste_tables <- function(base_access = "Tables_ref.accdb"){

  if (!file.exists(base_access)) {
    stop(paste0("La base Access \"", base_access, "\" n'existe pas"), call. = FALSE)
  }

  connexion <- impexp::access_connexion(base_access)

  liste_tables <- DBI::dbListTables(connexion) %>%
    stringr::str_subset("^[^(Msys)]")

  DBI::dbDisconnect(connexion)

  return(liste_tables)
}

#' Importer une table Access
#'
#' Importer une table Access.
#'
#' @param table Nom de la table à importer.
#' @param base_access Chemin de la base Access.
#'
#' @return Un data frame correspondant à la table Access.
#'
#' @examples
#' impexp::access_importer("Table_impexp",
#'   base_access = paste0(racine_packages, "impexp/inst/extdata/impexp.accdb"))
#'
#' @export
access_importer <- function(table, base_access = "Tables_ref.accdb"){

  if (!file.exists(base_access)) {
    stop(paste0("La base Access \"", base_access, "\" n'existe pas"), call. = FALSE)
  }

  connexion <- impexp::access_connexion(base_access)

  if (match(table, DBI::dbListTables(connexion)) %>% .[!is.na(.)] %>% length() == 0) {
    stop(paste0("Table \"", table, "\" non trouvee dans la base"), call. = FALSE)
  }

  import <- DBI::dbReadTable(connexion, table) %>%
    dplyr::as_tibble()

  DBI::dbDisconnect(connexion)

  import <- import %>%
    impexp::normaliser_nom_champs() %>%
    impexp::caracteres_vides_na() %>%
    dplyr::mutate_at(.vars = dplyr::vars(which(purrr::map_lgl(., ~ any(class(.) == "POSIXct")))), lubridate::as_date)

  return(import)
}

#' Exporter une table vers une base Access
#'
#' Exporter une table vers une base Access.
#'
#' @param table Data frame (sans guillemets) à exporter vers Access.
#' @param base_access Chemin de la base Access.
#' @param table_access Nom de la table Access créée. Par défaut, utilisation du nom du data frame.
#' @param ecraser \code{TRUE}, exporter même si une table Access du même nom exite déjà; \code{FALSE}, ne pas écraser si une table du même existe déjà.
#'
#' @examples
#' # Export d'un data frame dont le nom dans Access sera "data_frame"
#' impexp::access_exporter(data_frame, base_access = "Chemin/vers/une/base/Access.accdb")
#'
#' # Export d'un data frame dont le nom dans Access sera "export"
#' impexp::access_exporter(data_frame, base_access = "Chemin/vers/une/base/Access.accdb",
#'   table_access = "export")
#'
#' @export
access_exporter <- function(table, base_access = "Tables_Ref.accdb", table_access = NULL, ecraser = TRUE){

  if (is.null(table_access)) {
    table_access <- deparse(substitute(table))
  }

  # https://github.com/tidyverse/dbplyr/pull/36
  #
  # connexion <- impexp::access_connexion(base_access)
  #
  # if (intersect(DBI::dbListTables(connexion), table_access) %>% length() != 0 & ecraser) {
  #   message("Table \"", table_access, "\" ecrasée")
  #   DBI::dbRemoveTable(connexion, table_access)
  # }
  #
  # colnames(table) <- toupper(colnames(table))
  #
  # DBI::dbWriteTable(connexion, name = table_access, value = table)
  #
  # DBI::dbDisconnect(connexion)

  connexion <- RODBC::odbcConnectAccess2007(base_access)
  connexion_dbi <- impexp::access_connexion(base_access)

  if (intersect(DBI::dbListTables(connexion_dbi), table_access) %>% length() != 0 & ecraser) {
    message("Table \"", table_access, "\" ecrasée")
    RODBC::sqlDrop(connexion, table_access)
  }
  DBI::dbDisconnect(connexion_dbi)

  colnames(table) <- toupper(colnames(table))

  RODBC::sqlSave(connexion, dat = table, tablename = table_access, rownames = FALSE)

  RODBC::odbcClose(connexion)
}

#' Executer des commandes SQL dans une base Access
#'
#' Exécuter des commandes SQL dans une base Access
#'
#' @param liste_sql Un vecteur de commandes SQL au format chaîne de caractères.
#' @param base_access Chemin de la base Access.
#'
#' @export
access_executer_sql <- function(liste_sql, base_access = "Tables_Ref.accdb") {

  if (!file.exists(base_access)) {
    stop("La base Access \"", base_access,"\" n'existe pas...", call. = FALSE)
  }

  connexion <- impexp::access_connexion(base_access)

  purrr::walk(liste_sql, ~ DBI::dbExecute(connexion, .))

  DBI::dbDisconnect(connexion)
}
