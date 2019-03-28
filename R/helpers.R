#' Normalise column names of a data frame.
#'
#' @param data A data frame.
#'
#' @return A data frame with normalised colmun names.\cr
#'
#' Function \code{caractr::str_normalise_field()} is applied for each data column.
#'
#' @examples
#' data <- dplyr::data_frame(
#'   "Type d'unité Sirus : entreprise profilée ou unité légale" = NA_character_,
#'   "Nic du siège"= NA_character_
#' )
#' impexp::normalise_colnames(data)
#'
#' @export
normalise_colnames <- function(data){

  colnames(data) <- str_normalise_colnames(colnames(data))

  return(data)
}
