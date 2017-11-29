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
connexion_access <- function(base_access = "Tables_ref.accdb") {

  if (Sys.info()['sysname'] == "Windows"){
    if(!stringr::str_detect(base_access, "[A-Z]:\\/")) {
      dbq <- paste0(getwd(), "/", base_access)
    } else dbq <- base_access
    encoding <- "Windows-1252"
  } else if (Sys.info()['sysname'] == "Linux") {
    encoding <- "UTF-8"
  }

  connexion <- DBI::dbConnect(odbc::odbc(),
                              driver = "Microsoft Access Driver (*.mdb, *.accdb)",
                              dbq = iconv(dbq, from = "UTF-8"),
                              encoding = encoding)
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
#' importr::liste_tables_access(paste0(racine_packages, "importr/inst/extdata/importr.accdb"))
#'
#' @export
liste_tables_access <- function(base_access = "Tables_ref.accdb"){

  if (!file.exists(base_access)) {
    stop(paste0("La base Access \"", base_access, "\" n'existe pas"), call. = FALSE)
  }

  if (Sys.info()['sysname'] == "Windows"){

    connexion <- importr::connexion_access(base_access)

    liste_tables <- DBI::dbListTables(connexion) %>%
      stringr::str_subset("^[^(Msys)]")

    DBI::dbDisconnect(connexion)

  } else if (Sys.info()['sysname'] == "Linux") {

    base_access <- stringr::str_replace_all(base_access, fixed(" "), "\\ ")
    liste_tables <- system2("mdb-tables", base_access, stdout = TRUE) %>%
      strsplit(" ") %>%
      dplyr::pull(1)
  }

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
#' importr::importer_table_access("Table_importr",
#'   base_access = paste0(racine_packages, "importr/inst/extdata/importr.accdb"))
#'
#' @export
importer_table_access <- function(table, base_access = "Tables_ref.accdb"){

  if (!file.exists(base_access)) {
    stop(paste0("La base Access \"", base_access, "\" n'existe pas"), call. = FALSE)
  }

  if (Sys.info()['sysname'] == "Windows"){
    #Session R 32bits necessaire

    connexion <- importr::connexion_access(base_access)

    if (match(table, DBI::dbListTables(connexion)) %>% .[!is.na(.)] %>% length() == 0) {
      stop(paste0("Table \"", table, "\" non trouvee dans la base"), call. = FALSE)
    }

    import <- DBI::dbReadTable(connexion, table) %>%
      dplyr::as_tibble()

    DBI::dbDisconnect(connexion)

  } else if (Sys.info()['sysname'] == "Linux") {
    base_access <- stringr::str_replace_all(base_access, stringr::fixed(" "), "\\ ")

    liste_tables <- system2("mdb-tables", base_access, stdout = TRUE) %>%
      strsplit(" ") %>%
      dplyr::pull(1)

    if (match(table, liste_tables) %>% .[!is.na(.)] %>% length() == 0) {
      stop(paste0("Table \"", table, "\" non trouvée dans la base"), call. = FALSE)
    }

    system2("mdb-export", paste(base_access, table,"> import_access.csv"))
    import <- read.csv("import_access.csv")
    file.remove("import_access.csv")

    format_col_type <- c("Text" = "character", "Long Integer" = "integer", "DateTime" = "Date")

    col_type <- system2("mdb-schema", c(base_access, "-T", table), stdout = TRUE) %>%
      stringr::str_match("\t\\[([[:alnum:]_]+)\\]\t\t\t([A-z ]+)(\\(\\d+\\))?,?") %>% .[, 3] %>%
      .[which(!is.na(.))] %>%
      trimws() %>%
      format_col_type[.]

    diff_type <- dplyr::data_frame(champ = names(import), col_type, import = lapply(import, class)) %>%
      dplyr::filter(col_type != import)

    if (nrow(diff_type) >= 1) {

      for(champ in diff_type$champ) {
        if (diff_type$col_type[which(diff_type$champ == champ)] == "character") {
          import[[champ]] <- as.character(import[[champ]])
        } else if (diff_type$col_type[which(diff_type$champ == champ)] == "integer") {
          import[[champ]] <- as.integer(import[[champ]])
        } else if (diff_type$col_type[which(diff_type$champ == champ)] == "Date") {
          import[[champ]] <- as.Date(import[[champ]], "%m/%d/%y %H:%M:%S")
        }
      }

    }
  }

  import <- import %>%
    importr::normaliser_nom_champs() %>%
    importr::caracteres_vides_na() %>%
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
#' importr::exporter_table_access(data_frame, base_access = "Chemin/vers/une/base/Access.accdb")
#'
#' # Export d'un data frame dont le nom dans Access sera "export"
#' importr::exporter_table_access(data_frame, base_access = "Chemin/vers/une/base/Access.accdb",
#'   table_access = "export")
#'
#' @export
exporter_table_access <- function(table, base_access = "Tables_Ref.accdb", table_access = NULL, ecraser = TRUE){
  #Session R 32bits necessaire

  if (is.null(table_access)) {
    table_access <- deparse(substitute(table))
  }

  # https://github.com/tidyverse/dbplyr/pull/36
  #
  # connexion <- importr::connexion_access(base_access)
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
  connexion_dbi <- importr::connexion_access(base_access)

  if (intersect(DBI::dbListTables(connexion_dbi), table_access) %>% length() != 0 & ecraser) {
    message("Table \"", table_access, "\" ecrasée")
    RODBC::sqlDrop(connexion, table_access)
  }
  DBI::dbDisconnect(connexion_dbi)

  colnames(table) <- toupper(colnames(table))

  RODBC::sqlSave(connexion, dat = table, tablename = table_access, rownames = FALSE)

  RODBC::odbcClose(connexion)
}
