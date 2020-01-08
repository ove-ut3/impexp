#' Create an empty SQLite database.
#'
#' @param path SQLite database path.
#'
#' @export
sqlite_create <- function(path) {

  creation <- DBI::dbConnect(RSQLite::SQLite(), dbname = path)
  DBI::dbDisconnect(creation)

  return(path)
}

#' Return table names from a SQLite database.
#'
#' @param path SQLite database path.
#'
#' @export
sqlite_list_tables <- function(path) {

  if (!file.exists(path)) {
    stop("SQLite database \"", path,"\" does not exist.", call. = FALSE)
  }

  connection <- DBI::dbConnect(RSQLite::SQLite(), dbname = path)

  sqlite_liste_tables <- DBI::dbListTables(connection)

  DBI::dbDisconnect(connection)

  return(sqlite_liste_tables)
}

#' Import a SQLite database table.
#'
#' @param path SQLite database path.
#' @param table Name of the table to import as a character.
#' @param wait_unlock Wait until SQLite database is unlocked.
#' @param \dots Additional arguments passed to \code{DBI::dbReadTable}.
#'
#' @return A tibble.
#'
#' @export
sqlite_import <- function(path, table = NULL, wait_unlock = TRUE, ...) {

  if (!file.exists(path)) {
    stop("SQLite database \"", path,"\" does not exist.", call. = FALSE)
  }

  connection <- DBI::dbConnect(RSQLite::SQLite(), dbname = path)

  if (is.null(table)) {
    table <- path %>%
      stringr::str_extract("([^/]+?)$") %>%
      tools::file_path_sans_ext()
  }

  if (!table %in% DBI::dbListTables(connection)) {
    DBI::dbDisconnect(connection)
    stop("Table \"", table,"\" does not exist in \"", path,"\".", call. = FALSE)
  }

  if (wait_unlock == FALSE) {
    table <- DBI::dbReadTable(connection, table, ...)
  } else {
    table <- sqlite_wait_unlock(DBI::dbReadTable, connection, name = table, ...)
  }

  table <- dplyr::as_tibble(table)

  DBI::dbDisconnect(connection)

  return(table)
}

#' Export a table to a SQLite database.
#'
#' @param path SQLite database path.
#' @param data Data frame to export, unquoted.
#' @param table_name Optional name of the table to export as a character. By default, the name of sqlite database is used.
#' @param wait_unlock Wait until SQLite database is unlocked.
#' @param \dots Additional arguments to \code{DBI::dbWriteTable}.
#'
#' @export
sqlite_export <- function(path, data, table_name = NULL, wait_unlock = TRUE, ...) {

  if (!file.exists(path)) {
    stop("SQLite database \"", path,"\" does not exist.", call. = FALSE)
  }

  connection <- DBI::dbConnect(RSQLite::SQLite(), dbname = path)

  if (is.null(table_name)) {
    table_name <- path %>%
      stringr::str_extract("([^/]+?)$") %>%
      tools::file_path_sans_ext()
  }

  if (wait_unlock == FALSE) {
    DBI::dbWriteTable(connection, name = table_name, value = data, row.names = FALSE, ...)
  } else {
    sqlite_wait_unlock(DBI::dbWriteTable, connection, name = table_name, value = data, row.names = FALSE, ...)
  }

  DBI::dbDisconnect(connection)
}

#' Append rows to a table in a SQLite database.
#'
#' @param path SQLite database path.
#' @param data Data frame to append, unquoted.
#' @param table_name Table name in SQLite database to append data.
#' @param wait_unlock Wait until SQLite database is unlocked.
#'
#' @export
sqlite_append_rows <- function(path, data, table_name = NULL, wait_unlock = TRUE) {

  if (!file.exists(path)) {
    stop("SQLite database \"", path,"\" does not exist.", call. = FALSE)
  }

  if (is.null(table_name)) {
    table_name <- path %>%
      stringr::str_extract("([^/]+?)$") %>%
      tools::file_path_sans_ext()
  }

  if (ncol(sqlite_import(path, table_name)) != ncol(data)) {
    stop("data and table_name data must have the same number of columns.", call. = FALSE)
  }

  connection <- DBI::dbConnect(RSQLite::SQLite(), dbname = path)

  if (wait_unlock == FALSE) {
    DBI::dbWriteTable(connection, table_name, data, append = TRUE, row.names = FALSE)
  } else {
    sqlite_wait_unlock(DBI::dbWriteTable, connection, name = table_name, value = data, append = TRUE, row.names = FALSE)
  }

  DBI::dbDisconnect(connection)
}

#' Execute SQL queries in a SQLite database.
#'
#' @param path SQLite database path.
#' @param sql A SQL query.
#' @param wait_unlock Wait until SQLite database is unlocked.
#'
#' @export
sqlite_execute_sql <- function(path, sql, wait_unlock = TRUE) {

  if (!file.exists(path)) {
    stop("SQLite database \"", path,"\" does not exist.", call. = FALSE)
  }

  connection <- DBI::dbConnect(RSQLite::SQLite(), dbname = path)

  sql <- stringr::str_replace_all(sql, "([^'])'([^'])", "\\1''\\2")

  if (wait_unlock == FALSE) {
    DBI::dbExecute(connection, sql)
  } else {
    sqlite_wait_unlock(DBI::dbExecute, connection, statement = sql)
  }

  DBI::dbDisconnect(connection)
}
