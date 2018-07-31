#' Import an objet from a RData file.
#'
#' @param rdata RData file path.
#' @param object Object name to load.
#'
#' @return If parameter \code{object} stays \code{NULL} then all objects of the RData file loaded in a named list.\cr
#' Otherwise, only the object specified is loaded.
#'
#' @examples
#' divr::rdata_import(paste0(racine_packages, "divr/data/rdata_import.RData"))
#' divr::rdata_import(paste0(racine_packages, "divr/data/rdata_import.RData"), "lib_num_mois")
#'
#' @export
rdata_import <- function(rdata, object = NULL){

  if (!file.exists(rdata)) {
    stop(paste0("File \"", rdata, "\" doas not exist"), call. = FALSE)
  }

  env = new.env()
  load(file = rdata, envir = env)

  rdata_file <- stringr::str_match(rdata, "([^\\/]+?\\.RData)$")[, 2]

  if (length(env) == 1 & paste0(names(env)[1], ".RData") == rdata_file) {
    rdata_import <- env[[names(env)]]
    return(rdata_import)
  }

  if (!is.null(object)) {
    rdata_import <- env[[object]]
  } else {
    rdata_import <- as.list(env)
  }

  return(rdata_import)
}
