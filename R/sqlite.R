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
    stop("SQLite databse \"", path,"\" does not exist.", call. = FALSE)
  }

  connection <- DBI::dbConnect(RSQLite::SQLite(), dbname = path)

  sqlite_liste_tables <- DBI::dbListTables(connection)

  DBI::dbDisconnect(connection)

  return(sqlite_liste_tables)
}

#' Import a SQLite database table.
#'
#' @param table Name of the table to import as a character.
#' @param path SQLite database path.
#'
#' @return A tibble.
#'
#' @export
sqlite_import <- function(table, path) {

  if (!file.exists(path)) {
    stop("SQLite databse \"", path,"\" does not exist.", call. = FALSE)
  }

  connection <- DBI::dbConnect(RSQLite::SQLite(), dbname = path)

  if (!table %in% DBI::dbListTables(connection)) {
    DBI::dbDisconnect(connection)
    stop("Table \"", table,"\" does not exist in \"", path,"\".", call. = FALSE)
  }

  table <- DBI::dbReadTable(connection, table) %>%
    dplyr::as_tibble()

  DBI::dbDisconnect(connection)

  return(table)
}

#' Export a table to a SQLite database.
#'
#' @param data Data frame to export, unquoted.
#' @param path SQLite database path.
#' @param table_name Optional name of the table to export as a character. By default, the name of the data frame is used.
#' @param override If \code{TRUE} then the new data frame override the SQLite table if it already exists in the database.
#'
#' @export
sqlite_export <- function(table, path, table_name = NULL, override = FALSE, message = TRUE) {

  if (!file.exists(path)) {
    stop("SQLite databse \"", path,"\" does not exist.", call. = FALSE)
  }

  connection <- DBI::dbConnect(RSQLite::SQLite(), dbname = path)

  if (is.null(table_name)) {
    table_name <- deparse(substitute(table))
  }

  DBI::dbWriteTable(connection, name = table_name, value = table, row.names = FALSE, overwrite = override)

  DBI::dbDisconnect(connection)
}

#' Append rows to a table in a SQLite database.
#'
#' @param data Data frame to append, unquoted.
#' @param table Table name in SQLite database to append data.
#' @param path SQLite database path.
#'
#' @export
sqlite_append_rows <- function(data, table_name, path) {

  if (!file.exists(path)) {
    stop("SQLite databse \"", path,"\" does not exist.", call. = FALSE)
  }

  if (ncol(sqlite_import(table_name, path)) != ncol(data)) {
    stop("data and table_name data must have the same number of columns.", call. = FALSE)
  }

  connection <- DBI::dbConnect(RSQLite::SQLite(), dbname = path)

  DBI::dbWriteTable(connection, table_name, data, append = TRUE, row.names = FALSE)

  DBI::dbDisconnect(connection)
}

#' Execute SQL queries in a SQLite database.
#'
#' @param sql_list A vector of SQL queries.
#' @param path SQLite database path.
#' @param wait_unlock Wait until SQLite database is unlocked.
#'
#' @export
sqlite_execute_sql <- function(sql_list, path, wait_unlock = TRUE) {

  if (!file.exists(path)) {
    stop("SQLite databse \"", path,"\" does not exist.", call. = FALSE)
  }

  connection <- DBI::dbConnect(RSQLite::SQLite(), dbname = path)

  if (wait_unlock == FALSE) {
    purrr::walk(sql_list, ~ DBI::dbExecute(connection, .))

  } else {
    execute_sql_safe <- purrr::safely(DBI::dbExecute)

    execute_sql <- purrr::map(sql_list, ~ execute_sql_safe(connection, .))

    temoin_erreur <- purrr::map_lgl(execute_sql, ~ !is.null(.$error))

    while(any(temoin_erreur)) {

      sql_list <- sql_list[temoin_erreur]

      execute_sql <- purrr::map(sql_list, ~ execute_sql_safe(connection, .))

      temoin_erreur <- purrr::map_lgl(execute_sql, ~ !is.null(.$error))
    }
  }

  DBI::dbDisconnect(connection)
}
