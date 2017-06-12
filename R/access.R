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
    #Session R 32bits necessaire

    connexion <- RODBC::odbcConnectAccess2007(base_access)
    liste_tables <- liste_tables_odbc(connexion)

  } else if (Sys.info()['sysname'] == "Linux") {

    base_access <- stringr::str_replace_all(base_access, fixed(" "), "\\ ")
    liste_tables <- system2("mdb-tables", base_access, stdout = T) %>%
      strsplit(" ") %>% .[[1]]
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

    connexion <- RODBC::odbcConnectAccess2007(base_access)
    liste_tables <- importr::liste_tables_odbc(connexion)

    if (match(table, liste_tables) %>% .[!is.na(.)] %>% length() == 0) {
      stop(paste0("Table \"", table, "\" non trouvee dans la base"), call. = FALSE)
    }

    import <- RODBC::sqlFetch(connexion, table, as.is = TRUE) %>%
      dplyr::as_data_frame()

    RODBC::odbcClose(connexion)

  } else if (Sys.info()['sysname'] == "Linux") {
    base_access <- stringr::str_replace_all(base_access, stringr::fixed(" "), "\\ ")

    liste_tables <- system2("mdb-tables", base_access, stdout = TRUE) %>%
      strsplit(" ") %>% .[[1]]

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

    diff_type <- dplyr::data_frame(champ = names(import), col_type, import = purrr::map_chr(import, class)) %>%
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

  import <- importr::normaliser_nom_champs(import) %>%
    importr::caracteres_vides_na()

  if (any(purrr::map_chr(import, class) == "character")) {
    import <-dplyr::mutate_at(import, .vars = dplyr::vars(which(purrr::map_chr(import, class) == "character")), iconv, to = "UTF-8")
  }

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

  if(is.null(table_access)) table_access <- deparse(substitute(table))

  connexion <- RODBC::odbcConnectAccess2007(base_access)

  liste_tables <- importr::liste_tables_odbc(connexion)

  if (intersect(liste_tables, table_access) %>% length() != 0 & ecraser){
    message("Table ", table_access, " ecrasee")
    RODBC::sqlDrop(connexion, table_access)
  }

  colnames(table) <- toupper(colnames(table))

  RODBC::sqlSave(connexion, dat = table, tablename = table_access, rownames = F)

  RODBC::odbcClose(connexion)
}
