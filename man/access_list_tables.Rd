% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/access.R
\name{access_list_tables}
\alias{access_list_tables}
\title{List tables within a Microsoft Access.}
\usage{
access_list_tables(path, sys_tables = FALSE)
}
\arguments{
\item{path}{Path to the Access database.}

\item{sys_tables}{If \code{TRUE} then returns also system tables.}
}
\value{
A character vector containing table names.
}
\description{
List tables within a Microsoft Access.
}
\examples{
impexp::access_list_tables(system.file("extdata/impexp.accdb", package = "impexp"))

}
