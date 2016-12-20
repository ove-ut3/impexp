#' Connexion a une librairie ODBC
#'
#' Connexion à une librairie ODBC.
#'
#' @param librairie Librairie ODBC à laquelle se connecter.
#'
#' @return Un objet de type connexion.\cr
#'
#' Les librairies configurées sont les suivantes :\cr
#' - \code{bce} : accès à la BCE\cr
#' - \code{siredo} : accès aux schémas PEDR, IUF, USR\cr
#' - \code{rnsr} : accès au RNSR\cr
#' - \code{iuf_2016} : accès à IUF à partir de 2016 (PostgreSQL)\cr
#'
#' Il est possible de lister les schémas d'une connexion avec la fonction \code{liste_schemas_odbc()}
#' et les tables d'une connexion avec la fonction \code{liste_tables_odbc()}.
#'
#' @examples
#' connexion_bce <- importr::connexion_odbc("bce")
#'
#' @export
connexion_odbc = function(librairie){

  if(librairie == "bce") {
    #schema BCE
    connexion_odbc <- RODBC::odbcConnect("bce32", uid = "bceconsulter", pwd = "bce", rows_at_time = 500)
  } else if(librairie == "siredo") {
    #schemas : PEDR, IUF, USR
    connexion_odbc <- RODBC::odbcConnect("siredo", uid = "lectureall", pwd = "csilectbo", rows_at_time = 500)
  } else if(librairie == "rnsr") {
    #schema RNSR
    connexion_odbc <- RODBC::odbcConnect("repere", uid = "lectureall", pwd = "csilectbo", rows_at_time = 500)
  } else if(librairie == "iuf_2016") {
    #schema IUF à partir de 2016 (postgreSQL)
    connexion_odbc <- RODBC::odbcConnect("iuf_prod_pg_stat", uid = "statiuf", pwd = "statiuf", rows_at_time = 500)
  }

  return(connexion_odbc)
}

#' Lister les schemas d'une base ODBC
#'
#' Lister les schémas d'une base ODBC.
#'
#' @param connexion_odbc Connexion ODBC établie avec la fonction \code{connexion_odbc()}.
#'
#' @return Un vecteur des schémas ODBC de la connexion.
#'
#' @examples
#' # La liste des schémas de la BCE
#' importr::connexion_odbc("bce") %>%
#'   importr::liste_schemas_odbc()
#'
#' @export
liste_schemas_odbc <- function(connexion_odbc){

  liste_schemas_odbc <- RODBC::sqlTables(connexion_odbc) %>%
    .[["TABLE_SCHEM"]] %>%
    unique()

  return(liste_schemas_odbc)
}

#' Lister les tables d'une base ODBC (filtre possible sur un schema)
#'
#' Lister les tables d'une base ODBC (filtre possible sur un schéma).
#'
#' @param connexion_odbc Connexion ODBC établie avec la fonction \code{connexion_odbc()}.
#' @param schema Nom du schéma ODBC avec lequel filtrer la liste des tables. Par défaut, toutes les tables sont retournées.
#'
#' @return Un vecteur des tables ODBC de la connexion, filtré par le schéma s'il est rempli.
#'
#' @examples
#' # La liste de toutes les tables de la BCE
#' importr::connexion_odbc("bce") %>%
#'   importr::liste_tables_odbc()
#'
#' # La liste des tables de la BCE du schéma "BCE"
#' importr::connexion_odbc("bce") %>%
#'   importr::liste_tables_odbc(schema = "BCE")
#'
#' @export
liste_tables_odbc <- function(connexion_odbc, schema = NULL){

  liste_tables <- RODBC::sqlTables(connexion_odbc) %>%
    dplyr::as_data_frame() %>%
    normaliser_nom_champs() %>%
    dplyr::filter(table_type == "TABLE")

  if (!is.null(schema)) {
    liste_tables <- dplyr::filter_(liste_tables, paste0("table_schem == '", schema, "'"))
  }

  liste_tables_odbc <- liste_tables[["table_name"]]

  return(liste_tables_odbc)
}

#' Importer une table d'une base ODBC
#'
#' Importer une table d'une base ODBC.
#'
#' @param table Nom de table ODBC à importer.
#' @param connexion_odbc Connexion ODBC établie avec la fonction \code{connexion_odbc()}.
#' @param schema Nom du schéma ODBC correspondant.
#' @param message_table Affichage d'un message dans la console lors de l'import de la table.
#'
#' @return Un data frame correspondant à la table ODBC.\cr
#'
#' L'import est plus rapide si le paramètre \code{schema} est rempli.
#'
#' @examples
#' # Import de la table "N_NATURE_UAI" de la BCE
#' connexion_bce <- importr::connexion_odbc("bce")
#' n_nature_uai <- importr::importer_table_odbc("N_NATURE_UAI", connexion_bce, schema = "BCE")
#'
#' @export
importer_table_odbc <- function(table, connexion_odbc, schema = NULL, message_table = FALSE){

  if (message_table == TRUE) message("Import de la table \"", table, "\"")

  if(!is.null(schema)) {
    table <- paste(schema, table, sep = ".")
  }

  df <- RODBC::sqlFetch(connexion_odbc, table)

  if (!is.data.frame(df)) {
    df <- as.data.frame(df)
  }

  table_oracle <- dplyr::as_data_frame(df) %>%
    normaliser_nom_champs()

  return(table_oracle)
}

#' Importer une base ODBC
#'
#' Importer une base ODBC.
#'
#' @param librairie Librairie ODBC à laquelle se connecter.
#' @param schema Nom du schéma ODBC avec lequel filtrer la liste des tables importées.
#' @param message_table Affichage d'un message dans la console lors de l'import de la table.
#' @param fichier_sauvegarde Chemin du fichier RData de sauvegarde.
#' @param copie_rdata_date \code{TRUE}, copie datée de la base pour archive; \code{FALSE}, aucune copie n'est réalisée.
#'
#' @return Une liste nommée de data frame correspondant aux tables de la base ODBC.
#'
#' @examples
#' # Import de la base IUF (2015 et avant)
#' iuf <- importr::importer_base_odbc(librairie = "siredo", schema = "IUF")
#' iuf
#'
#' @export
importer_base_odbc <- function(librairie, schema = NULL, message_table = TRUE, fichier_sauvegarde = NULL, copie_rdata_date = FALSE){

  connexion_odbc <- connexion_odbc(librairie)

  tables <- liste_tables_odbc(connexion_odbc, schema)

  base <- purrr::map(tables, importer_table_odbc, connexion_odbc, schema, message_table = message_table)
  names(base) <- tolower(tables)

  assign(librairie, base)

  if (is.null(fichier_sauvegarde)) return(base)
  else {
    save(list = librairie, file = fichier_sauvegarde)

    if (copie_rdata_date == TRUE) {
      date <- Sys.Date() %>% as.character() %>% stringr::str_replace_all("-", "_")
      fichier_sauvegarde <- stringr::str_replace(fichier_sauvegarde, "\\.RData$", paste0("_", date, ".RData"))
      save(list = librairie, file = fichier_sauvegarde)
    }
  }
}
