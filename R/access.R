#' List tables within a Microsoft Access.
#'
#' @param path Path to the Access database.
#' @param sys_tables If \code{TRUE} then returns also system tables.
#'
#' @return A character vector containing table names.
#'
#' @examples
#' impexp::access_list_tables(path = paste0(find.package("impexp"), "/extdata/impexp.accdb))
#'
#' @export
access_list_tables <- function(path, sys_tables = FALSE){

  connection <- access_connect(path)

  tables <- DBI::dbListTables(connection)

  DBI::dbDisconnect(connection)

  if (sys_tables == FALSE) {
    tables <- tables[which(!stringr::str_detect(tables, "^MSys"))]
  }

  return(tables)
}

#' Import a Microsoft Access database table.
#'
#' @param table Name of the table to import as a character.
#' @param path Path to the Access database.
#'
#' @return A data frame.
#'
#' @examples
#' impexp::access_import("Table_impexp",
#'   path = paste0(find.package("impexp"), "/extdata/impexp.accdb"))
#'
#' @export
access_import <- function(table, path){

  connection <- access_connect(path)

  if (match(table, DBI::dbListTables(connection)) %>% .[!is.na(.)] %>% length() == 0) {
    stop(glue::glue("Table \"{table}\" not found in the database"), call. = FALSE)
  }

  import <- DBI::dbReadTable(connection, table) %>%
    dplyr::as_tibble()

  DBI::dbDisconnect(connection)

  import <- import %>%
    dplyr::rename_all(tolower) %>%
    dplyr::mutate_if(is.character, dplyr::na_if, "") %>%
    dplyr::mutate_at(.vars = dplyr::vars(which(purrr::map_lgl(., ~ any(class(.) == "POSIXct")))), lubridate::as_date)

  return(import)
}
