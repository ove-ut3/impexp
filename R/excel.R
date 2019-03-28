#' Import Microsoft Excel files located in a path.
#'
#' @param pattern A regular expression. Only file names matching the regular expression will be imported.
#' @param path Path where the excel files are located (recursive).
#' @param pattern_sheet A regular expression. Only sheet names in excel files matching the regular expression will be imported.
#' @param parallel If \code{TRUE} then excel files are imported using all CPU cores (using parallel & pbapply packages).
#' @param zip If \code{TRUE} then excel files within zip files are also imported.
#' @param progress_bar If \code{TRUE} then a progress bar is displayed (using pbapply package).
#' @param message If \code{TRUE} then a message indicates how many files are imported.
#' @param \dots Optional arguments from \code{readxl::read_excel} function.
#'
#' @return A data frame with a column-list "import" containing all tibbles.
#'
#' @examples
#' impexp::excel_import_path(paste0(find.package("impexp"), "/extdata"), pattern_sheet = "impexp")
#' impexp::excel_import_path(paste0(find.package("impexp"), "/extdata"), pattern_sheet = "impexp", skip = 1)
#' impexp::excel_import_path(paste0(find.package("impexp"), "/extdata"), pattern_sheet = "impexp", zip = TRUE)
#'
#' @export
excel_import_path <- function(path = ".", pattern = "\\.xlsx?$", pattern_sheet = ".", parallel = FALSE, zip = FALSE, progress_bar = FALSE, message = FALSE, ...) {

  if (!dir.exists(path)) {
    stop("The path \"", path,"\" does not exist", call. = FALSE)
  }

  files <- dplyr::tibble(file = list.files(path, recursive = TRUE, full.names = TRUE) %>%
                           stringr::str_subset(pattern) %>%
                           iconv(from = "UTF-8") %>%
                           purrr::map_chr(tools::file_path_as_absolute),
                         zip_files = NA_character_)

  # If zip files are included
  if (zip == TRUE) {

    zip_files <- zip_extract_path(path, pattern = pattern, parallel = parallel, progress_bar = progress_bar, message = message, files = TRUE)

    files <- dplyr::bind_rows(zip_files, files) %>%
      dplyr::arrange(file)

  } else {
    files <- dplyr::mutate(files, zip_file = NA_character_)
  }

  if (nrow(files) == 0) {
    message("No files matches the parameters")

    return(files)
  }

  if (progress_bar == TRUE & !"pbapply" %in% installed.packages()[, 1]) {
    stop("pbapply package needs to be installed", call. = FALSE)
  }

  if (message == TRUE) {
    message(length(unique(files$file))," excel file(s) imported...")
  }

  if (parallel == TRUE & !all(c("parallel", "pbapply") %in% installed.packages()[, 1])) {
    stop("parallel and pbapply packages need to be installed", call. = FALSE)
  }

  if (parallel == TRUE) {
    cluster <- parallel::makeCluster(parallel::detectCores())
  } else {
    cluster <- NULL
  }

  if (progress_bar == TRUE | parallel == TRUE) {

    excel_import_path <- files %>%
      dplyr::mutate(sheet = purrr::map(file, readxl::excel_sheets)) %>%
      tidyr::unnest() %>%
      dplyr::filter(stringr::str_detect(sheet, pattern_sheet)) %>%
      dplyr::mutate(import = pbapply::pblapply(split(., 1:nrow(.)), function(import) {

        import(import$zip_file, import$file, pattern, "excel", ...)

      }, cl = cluster))

  } else {

    excel_import_path <- files %>%
      dplyr::mutate(sheet = purrr::map(file, readxl::excel_sheets)) %>%
      tidyr::unnest() %>%
      dplyr::filter(stringr::str_detect(sheet, pattern_sheet)) %>%
      dplyr::mutate(import = lapply(split(., 1:nrow(.)), function(import) {

        import(import$zip_file, import$file, pattern, "excel", ...)

      }))

  }

  if (parallel == TRUE) {
    parallel::stopCluster(cluster)
  }

  return(excel_import_path)
}

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

  if (!"openxlsx" %in% installed.packages()[, 1]) {
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
