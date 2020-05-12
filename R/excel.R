#' Export a Microsoft Excel file (using openxlsx).
#'
#' @param data A data frame or a named list of data frames that will be sheets in the Excel file.
#' @param path Path to the created Excel file.
#' @param create_dir if \code{TRUE} then the directory containing the Excel file is created.
#' @param sheet Sheet name if data is a data frame, or overrides list names.
#' @param footer Footer notes to place beneath tables in sheets.
#' @param n_cols_rowname Number of columns containg row names. These columns usualy don't need header.
#'
#' @export
excel_export <- function(data, path, create_dir = FALSE, sheet = NULL, footer = NULL, n_cols_rowname = 0) {

  if (!"openxlsx" %in% utils::installed.packages()[, 1]) {
    stop("openxlsx package needs to be installed", call. = FALSE)
  }

  if (any(class(data) == "data.frame")) {
    data <- list("data" = data)
  }

  if (any(purrr::map_lgl(data, ~ !any(class(.) == "data.frame")))) {
    stop("At least one of the objects is not a data frame", call. = FALSE)
  }

  if (is.null(sheet)) {
    sheet <- names(data)
  }

  workbook <- openxlsx::createWorkbook()

  if (length(footer) == 1) {
    footer <- rep(footer, length(data)) %>% as.list()

  } else if (!is.null(footer) & length(footer) != length(data)) {
    stop("Length of footer must be equal to the length of data", call. = FALSE)

  } else if (is.null(footer)) {
    footer <- rep(NA_character_, length(data)) %>% as.list()
  }

  purrr::pwalk(list(data, sheet, footer), excel_sheet, workbook = workbook, n_cols_rowname = n_cols_rowname)

  if (create_dir == TRUE) {
    stringr::str_match(path, "(.+)/[^/]+?$")[, 2] %>%
      dir.create(showWarnings = FALSE, recursive = TRUE)
  }

  openxlsx::saveWorkbook(workbook, path, overwrite = TRUE)

  return(path)
}
