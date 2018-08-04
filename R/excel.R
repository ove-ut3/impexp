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

    zip_files <- impexp::zip_extract_path(path, pattern = pattern, parallel = parallel) %>%
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

#' Create a sheet in a Microsoft Excel file.
#'
#' @param workbook An \code{openxlsx} worbook object.
#' @param data A data frame.
#' @param sheet A sheet name.
#' @param footer Footer notes to place beneath tables in sheet.
#' @param n_cols_rowname Number of columns containg row names. These columns usualy don't need header.
#'
#' @export
#' @keywords internal
excel_sheet <- function(workbook, data, sheet, footer = NULL, n_cols_rowname = 0) {

  openxlsx::addWorksheet(workbook, sheet)

  #n_cols_rowname <- length(stringr::str_subset(colnames(data), "^lib"))

  if (!is.null(data[["sous_titre_"]])) {
    line_subtitle <- which(data$sous_titre_ == "O")
    data <- dplyr::select(data, -sous_titre_)
  }

  if (!is.null(data[["indentation_"]])) {
    data <- data %>%
      dplyr::mutate(lib = paste0(purrr::map_chr(indentation_, ~ paste0(rep("    ", . - 1), collapse = "")), lib)) %>%
      dplyr::select(-indentation_)
  }

  if (!is.null(data[["dont_"]])) {
    line_subtotal <- which(data$dont_ == "O")
    data <- dplyr::select(data, -dont_)
  }

  if (!is.null(data[["dont_derniere_ligne_"]])) {
    line_subtotal_last <- which(data$dont_derniere_ligne_ == "O")
    data <- dplyr::select(data, -dont_derniere_ligne_)
  }

  style_figures <- openxlsx::createStyle(halign = "right")
  style_title1 <- openxlsx::createStyle(textDecoration = "bold", halign = "center")
  style_title2 <- openxlsx::createStyle(textDecoration = "bold", halign = "center", border = "bottom")
  style_subtitle <- openxlsx::createStyle(textDecoration = "bold", border = "bottom")
  style_subtotal <- openxlsx::createStyle(textDecoration = "italic")
  style_subtotal_last <- openxlsx::createStyle(textDecoration = "italic", border = "bottom")
  style_col_border <- openxlsx::createStyle(border = "right")

  #### Title ####

  full_title <- colnames(data) %>%
    paste0("##")

  if (n_cols_rowname != 0){
    full_title[1:n_cols_rowname] <- ""
  }

  line_title = 0

  while(any(stringr::str_detect(full_title, "##"), na.rm = TRUE)) {

    line_title <- line_title + 1

    title <- full_title %>%
      stringr::str_match("^(.+?)##") %>% .[, 2] %>%
      dplyr::tibble(title = .)

    full_title <- full_title %>%
      stringr::str_match("##(.+)") %>% .[, 2]

    if (line_title == 1) {
      full_title_bck <- title$title
    } else {
      full_title_bck <- paste(full_title_bck, title$title, sep = "##")
    }

    if (any(stringr::str_detect(full_title, "##"), na.rm = TRUE)) {
      merge <- title %>%
        dplyr::mutate(full_title_bck = full_title_bck) %>%
        dplyr::mutate(merge = dplyr::row_number()) %>%
        split(x = .$merge, f = .$full_title_bck)
      line_title_last <- FALSE
      style <- style_title1

      if (line_title == 1) {
        col_border <- purrr::map_int(merge, tail, 1)
      }

    } else {
      line_title_last <- TRUE
      style <- style_title2
    }

    title <- title %>%
      t() %>%
      dplyr::as_tibble()

    start_col <- ifelse(n_cols_rowname == 0, 1, n_cols_rowname)

    openxlsx::writeData(workbook, sheet, title, startCol = start_col, startRow = line_title, colNames = FALSE)
    openxlsx::addStyle(workbook, sheet, style, rows = line_title, 1:ncol(data), gridExpand = TRUE, stack = TRUE)

    if (line_title_last == FALSE) {
      purrr::walk(merge, ~ openxlsx::mergeCells(workbook, sheet, cols = ., rows = line_title))
    }

  }

  #### Body ####

  openxlsx::writeData(workbook, sheet, data, startRow = line_title + 1, colNames = FALSE)

  if (n_cols_rowname != 0) {
    openxlsx::addStyle(workbook, sheet, style_figures, rows = (line_title + 1):(line_title + nrow(data)), n_cols_rowname + 1:(ncol(data) + 1), gridExpand = TRUE, stack = TRUE)

  } else {
    col_figures <- purrr::map_int(data, ~ class(.) %in% c("integer", "double")) %>%
    { which(. != 0) }

    if (length(col_figures) != 0) {
      openxlsx::addStyle(workbook, sheet, style_figures, rows = (line_title + 1):(line_title + nrow(data)), col_figures, gridExpand = TRUE, stack = TRUE)
    }

  }

  # Auto-adjust column width
  if (n_cols_rowname != 0) {
    openxlsx::setColWidths(workbook, sheet, cols = 1:n_cols_rowname, widths = "auto")
  } else {
    openxlsx::setColWidths(workbook, sheet, cols = 1:ncol(data), widths = "auto")
  }

  if (exists("col_border")) {
    openxlsx::addStyle(workbook, sheet, style_col_border, rows = 1:(line_title + nrow(data)), cols = col_border, gridExpand = TRUE, stack = TRUE)

  }

  if (exists("line_subtitle")) {
    openxlsx::addStyle(workbook, sheet, style_subtitle, rows = line_title + line_subtitle, cols = 1:ncol(data), gridExpand = TRUE, stack = TRUE)

  }

  if (exists("line_subtotal")) {
    openxlsx::addStyle(workbook, sheet, style_subtotal, rows = line_title + line_subtotal, cols = 1:ncol(data), gridExpand = TRUE, stack = TRUE)

  }

  if (exists("line_subtotal_last")) {
    openxlsx::addStyle(workbook, sheet, style_subtotal_last, rows = line_title + line_subtotal_last, cols = 1:ncol(data), gridExpand = TRUE, stack = TRUE)

  }

  #### Footer ####

  if (!is.null(footer)) {
    note <- dplyr::tibble(note1 = footer)
    openxlsx::writeData(workbook, sheet, note, startRow = line_title + nrow(data) + 2, colNames = FALSE)
  }

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

  purrr::pwalk(list(data, sheet, footer), impexp::excel_sheet, workbook = workbook, n_cols_rowname = n_cols_rowname)

  if (create_dir == TRUE) {
    stringr::str_match(path, "(.+)/[^/]+?$")[, 2] %>%
      dir.create(showWarnings = FALSE, recursive = TRUE)
  }

  openxlsx::saveWorkbook(workbook, path, overwrite = TRUE)

  return(path)
}
