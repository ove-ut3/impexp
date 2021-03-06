% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/sqlite.R
\name{sqlite_execute_sql}
\alias{sqlite_execute_sql}
\title{Execute SQL queries in a SQLite database.}
\usage{
sqlite_execute_sql(path, sql, wait_unlock = TRUE)
}
\arguments{
\item{path}{SQLite database path.}

\item{sql}{A SQL query.}

\item{wait_unlock}{Wait until SQLite database is unlocked.}
}
\description{
Execute SQL queries in a SQLite database.
}
