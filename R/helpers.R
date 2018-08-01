#' Normaliser les noms des champs d'une table importee
#'
#' Normaliser les noms des champs d'une table importée.
#'
#' @param table Un data frame.
#'
#' @return Un data frame avec les noms de champ sont normalisés.\cr
#'
#' La fonction \code{caractr::str_normalise_field()} est appliquée sur chacun des noms de champ.
#'
#' @examples
#' table <- dplyr::data_frame(
#'   "Type d'unité Sirus : entreprise profilée ou unité légale" = NA_character_,
#'   "Nic du siège"= NA_character_
#' )
#' impexp::normaliser_nom_champs(table)
#'
#' @export
normaliser_nom_champs <- function(table){

  colnames(table) <- colnames(table) %>%
    caractr::str_normalise_colnames()

  if(length(names(table)) != length(unique(names(table)))) {
    names(table) <- make.unique(names(table), sep = "_")
  }

  return(table)
}

#' Remplacer les NA par des caracteres vides
#'
#' Remplacer les NA par des caractères vides.
#'
#' @param table Un data frame.
#'
#' @return Un data frame dont les champ à \code{""} sont transformés en \code{NA}.
#'
#' @examples
#' table <- dplyr::data_frame(
#'   champ_1 = "",
#'   champ_2 = ""
#' )
#' impexp::caracteres_vides_na(table)
#'
#' @export
#' @keywords internal
caracteres_vides_na <- function(table){

  if (nrow(table) != 0) {

    num_champs_character <- lapply(table, class) %>%
      purrr::map_chr(1) %>%
      { which(. == "character") }

    if (length(num_champs_character) != 0) {
      is.na(table[, num_champs_character]) <- table[, num_champs_character] == ''
    }

    return(table)

  } else {
    return(table)
  }

}
