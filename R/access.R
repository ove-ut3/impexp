#' Connect to a Microsoft Access database.
#'
#' @param path Path to the Access database.
#'
#' @return A DBI connection.
#'
#' @export
#' @keywords internal
access_connect <- function(path) {

  if (!file.exists(path)) {
    stop(paste0("The Access database\"", path, "\" does not exist"), call. = FALSE)
  }

  if(!stringr::str_detect(path, "[A-Z]:\\/")) {
    dbq <- paste0(getwd(), "/", path)
  } else {
    dbq <- path
  }

  connection <- DBI::dbConnect(odbc::odbc(),
                               driver = "Microsoft Access Driver (*.mdb, *.accdb)",
                               dbq = iconv(dbq, to = "Windows-1252"),
                               encoding = "Windows-1252")
  return(connection)
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
#'   path = paste0(path.package("impexp"), "/extdata/impexp.accdb"))
#'
#' @export
access_import <- function(table, path){

  connection <- impexp::access_connect(path)

  if (match(table, DBI::dbListTables(connection)) %>% .[!is.na(.)] %>% length() == 0) {
    stop(paste0("Table \"", table, "\" not found in the database"), call. = FALSE)
  }

  import <- DBI::dbReadTable(connection, table) %>%
    dplyr::as_tibble()

  DBI::dbDisconnect(connection)

  import <- import %>%
    impexp::normaliser_nom_champs() %>%
    impexp::caracteres_vides_na() %>%
    dplyr::mutate_at(.vars = dplyr::vars(which(purrr::map_lgl(., ~ any(class(.) == "POSIXct")))), lubridate::as_date)

  return(import)
}

#' Export a data frame to a Microsoft Access database.
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
  # connection <- impexp::access_connect(path)
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

  connection <- RODBC::odbcConnectAccess2007(path)
  connection_dbi <- impexp::access_connect(path)

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

  connexion <- impexp::access_connect(path)

  purrr::walk(sql, ~ DBI::dbExecute(connexion, .))

  DBI::dbDisconnect(connexion)
}
