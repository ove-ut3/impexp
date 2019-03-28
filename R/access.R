#' List tables within a Microsoft Access.
#'
#' @param path Path to the Access database.
#' @param sys_tables If \code{TRUE} then returns also system tables.
#'
#' @return A character vector containing table names.
#'
#' @examples
#' impexp::access_tables(path = paste0(find.package("impexp"), "/extdata/impexp.accdb"))
#'
#' @export
access_tables <- function(path, sys_tables = FALSE){

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
    dplyr::mutate_if(data, is.character, dplyr::na_if, "") %>%
    dplyr::mutate_at(.vars = dplyr::vars(which(purrr::map_lgl(., ~ any(class(.) == "POSIXct")))), lubridate::as_date)

  return(import)
}

#' Export a data frame to a Microsoft Access database (using RODBC).
#'
#' @param data Data frame to export, unquoted.
#' @param path Path to the Access database.
#' @param table_name Optional name of the table to export as a character. By default, the name of the data frame is used.
#' @param override If \code{TRUE} then the new data frame override the Access table if it already exists in the database.
#'
#' @examples
#' impexp::access_export(data_frame, path = "Path/to/the/database.accdb")
#' impexp::access_export(data_frame, path = "Path/to/the/database.accdb"), table_name = "export")
#'
#' @export
access_export <- function(data, path, table_name = NULL, override = TRUE){

  if (is.null(table_name)) {
    table_name <- deparse(substitute(data))
  }

  # https://github.com/r-dbi/odbc/issues/79
  #
  # connection <- access_connect(path)
  #
  # if (intersect(DBI::dbListTables(connection), table_name) %>% length() != 0 & override) {
  #   message("Table \"", table_name, "\" overridden")
  #   DBI::dbRemoveTable(connection, table_name)
  # }
  #
  # colnames(data) <- toupper(colnames(data))
  #
  # DBI::dbWriteTable(connection, name = table_name, value = data)
  #
  # DBI::dbDisconnect(connection)

  if (!"RODBC" %in% installed.packages()[, 1]) {
    stop("RODBC package needs to be installed", call. = FALSE)
  }

  connection <- RODBC::odbcConnectAccess2007(path)
  connection_dbi <- access_connect(path)

  if (intersect(DBI::dbListTables(connection_dbi), table_name) %>% length() != 0 & override) {
    message("Table \"", table_name, "\" overridden")
    RODBC::sqlDrop(connection, table_name)
  }
  DBI::dbDisconnect(connection_dbi)

  colnames(data) <- toupper(colnames(data))

  RODBC::sqlSave(connection, dat = data, tablename = table_name, rownames = FALSE)

  RODBC::odbcClose(connection)
}

#' Execute SQL queries in a Microsoft Access database.
#'
#' @param sql SQL queries as a character vector.
#' @param path Path to the Access database.
#'
#' @export
access_execute_sql <- function(sql, path) {

  connexion <- access_connect(path)

  purrr::walk(sql, ~ DBI::dbExecute(connexion, .))

  DBI::dbDisconnect(connexion)
}
