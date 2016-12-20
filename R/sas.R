#' Importer une table SAS
#'
#' Importer une table SAS
#'
#' @param fichier Chemin de table sas.
#'
#' @return Un data frame correspondant à la table SAS.
#'
#' @examples
#' importr::importer_table_sas(paste0(racine_packages, "importr/inst/extdata/importr.sas7bdat"))
#'
#' @export
importer_table_sas <- function(fichier) {

  importer_table_sas <- haven::read_sas(fichier) %>%
    normaliser_nom_champs() %>%
    caracteres_vides_na()

  return(importer_table_sas)
}
