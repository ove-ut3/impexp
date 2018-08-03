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
