#' Importer une table LibreOffice
#'
#' Importer une table LibreOffice.
#'
#' @param table Nom de la table à importer.
#' @param base_odb Chemin de la base LibreOffice.
#'
#' @return Un data frame correspondant à la table LibreOffice.\cr
#'
#' Ne fonctionne pas sous Windows.
#'
#' @examples
#' importr::importer_table_odb("importr",
#'   base_odb = paste0(racine_packages, "importr/inst/extdata/importr.odb"))
#'
#' @export
importer_table_odb <- function(table, base_odb = "Tables_ref.odb"){

  connexion <- ODB::odb.open(base_odb)

  liste_tables <- ODB::odb.tables(connexion) %>% names()

  if (match(table, liste_tables) %>% .[!is.na(.)] %>% length() == 0) {
    import <- NULL
  } else {
    import <- ODB::odb.read(connexion, paste0('SELECT * FROM "', table, '"')) %>%
      dplyr::as_data_frame() %>%
      normaliser_nom_champs() %>%
      caracteres_vides_na()
  }

  ODB::odb.close(connexion)

  return(import)

}
