#' Import an objet from a R data file.
#'
#' @param r_file R data file path.
#' @param object Object name to load.
#'
#' @return If parameter \code{object} stays \code{NULL} then all objects of the R data file loaded in a named list.\cr
#' Otherwise, only the object specified is loaded.
#'
#' @examples
#' impexp::r_import(paste0(find.package("impexp"), "/extdata/impexp.RData"))
#' impexp::r_import(paste0(find.package("impexp"), "/extdata/impexp.RData"), "data")
#'
#' @export
r_import <- function(r_data, object = NULL){

  if (!file.exists(r_data)) {
    stop(glue::glue("File \"{r_data}\" doas not exist"), call. = FALSE)
  }

  env = new.env()
  load(file = r_data, envir = env)

  r_data_file <- stringr::str_match(r_data, "([^\\/]+?)\\.(rda|RData)$")[, 2]

  if (length(env) == 1 & names(env)[1] == r_data_file) {
    r_data_import <- env[[names(env)]]
    return(r_data_import)
  }

  if (!is.null(object)) {
    r_data_import <- env[[object]]
  } else {
    r_data_import <- as.list(env)
  }

  return(r_data_import)
}
