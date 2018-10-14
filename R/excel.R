#' Import Microsoft Excel files located in a path.
#'
#' @param pattern A regular expression. Only file names matching the regular expression will be imported.
#' @param path Path where the excel files are located (recursive).
#' @param pattern_sheet A regular expression. Only sheet names in excel files matching the regular expression will be imported.
#' @param parallel If \code{TRUE} then excel files are imported using all CPU cores.
#' @param zip If \code{TRUE} then excel files within zip files are also imported.
#' @param message If \code{TRUE} then a message indicates how many files are imported.
#' @param \dots Optional arguments from \code{readxl::read_excel} function.
#'
#' @return A data frame whith a column-list "import" containing all tibbles.
#'
#' @examples
#' impexp::excel_import_path(path = paste0(find.package("impexp"), "/extdata"), pattern_sheet = "impexp")
#' impexp::excel_import_path(path = paste0(find.package("impexp"), "/extdata"), pattern_sheet = "impexp", skip = 1)
#' impexp::excel_import_path(path = paste0(find.package("impexp"), "/extdata"), pattern_sheet = "impexp", zip = TRUE)
#'
#' @export
excel_import_path <- function(pattern = "\\.xlsx?$", path = ".", pattern_sheet = ".", parallel = FALSE, zip = FALSE, message = TRUE, ...) {

  if (!dir.exists(path)) {
    stop("The path \"", path,"\" does not exist", call. = FALSE)
  }

  files <- dplyr::tibble(file = list.files(path, recursive = TRUE, full.names = TRUE) %>%
                           stringr::str_subset(pattern) %>%
                           iconv(from = "UTF-8"))

  # If zip files are included
  if (zip == TRUE) {

    zip_files <- zip_extract_path(path, pattern = pattern, parallel = parallel) %>%
      dplyr::select(-exdir)

    files <- dplyr::bind_rows(zip_files, files) %>%
      dplyr::arrange(file)

  } else {
    files <- files %>%
      dplyr::mutate(zip_file = NA_character_)
  }

  if (nrow(files) == 0) {
    message("No files matches the parameters")

    return(files)
  }

  if (message == TRUE) {
    message(length(unique(files$file))," Excel files imported...")
    pbapply::pboptions(type = "timer")
  } else {
    pbapply::pboptions(type = "none")
  }

  if (parallel == TRUE) {
    cluster <- parallel::makeCluster(parallel::detectCores())
  } else {
    cluster <- NULL
  }

  excel_import_path <- files %>%
    dplyr::mutate(sheet = purrr::map(file, readxl::excel_sheets)) %>%
    tidyr::unnest() %>%
    dplyr::filter(stringr::str_detect(sheet, pattern_sheet)) %>%
    dplyr::mutate(import = pbapply::pblapply(split(., 1:nrow(.)), function(import) {
      readxl::read_excel(import$file, import$sheet, ...)
    }, cl = cluster))

  suppression <- dplyr::filter(files, !is.na(zip_file)) %>%
    dplyr::pull(file) %>%
    file.remove()

  if (parallel == TRUE) {
    parallel::stopCluster(cluster)
  }

  if (zip == TRUE) {
    excel_import_path <- dplyr::left_join(files, excel_import_path, by = c("zip_file", "file")) %>%
      dplyr::mutate(file = stringr::str_match(file, "/(.+)")[, 2])

  }

  return(excel_import_path)
}

#' Export a Microsoft Excel file.
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
