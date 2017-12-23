#' Importer une table SAS
#'
#' Importer une table SAS
#'
#' @param fichier Chemin de table sas.
#'
#' @return Un data frame correspondant à la table SAS.
#'
#' @examples
#' importr::sas_importer(paste0(racine_packages, "importr/inst/extdata/importr.sas7bdat"))
#'
#' @export
sas_importer <- function(fichier) {

  sas_importer <- haven::read_sas(fichier) %>%
    normaliser_nom_champs() %>%
    caracteres_vides_na()

  return(sas_importer)
}
