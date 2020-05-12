access_connect <- function(path) {

  path <- stringr::str_replace(path, "/$", "") %>%
    tools::file_path_as_absolute()

  if (!file.exists(path)) {
    stop(glue::glue("The Access database\"{path}\" does not exist"), call. = FALSE)
  }

  connection <- DBI::dbConnect(odbc::odbc(),
                               driver = "Microsoft Access Driver (*.mdb, *.accdb)",
                               dbq = iconv(path, to = "Windows-1252"),
                               encoding = "Windows-1252")
  return(connection)
}

excel_sheet <- function(workbook, data, sheet, footer = NULL, n_cols_rowname = 0) {

  openxlsx::addWorksheet(workbook, sheet)

  #n_cols_rowname <- length(stringr::str_subset(colnames(data), "^lib"))

  if (!is.null(data[["sous_titre_"]])) {
    line_subtitle <- which(data$sous_titre_ == "O")
    data <- dplyr::select(data, -.data$sous_titre_)
  }

  if (!is.null(data[["indentation_"]])) {
    data <- data %>%
      dplyr::mutate(lib = paste0(purrr::map_chr(.data$indentation_, ~ paste0(rep("    ", . - 1), collapse = "")), .data$lib)) %>%
      dplyr::select(-.data$indentation_)
  }

  if (!is.null(data[["dont_"]])) {
    line_subtotal <- which(data$dont_ == "O")
    data <- dplyr::select(data, -.data$dont_)
  }

  if (!is.null(data[["dont_derniere_ligne_"]])) {
    line_subtotal_last <- which(data$dont_derniere_ligne_ == "O")
    data <- dplyr::select(data, -.data$dont_derniere_ligne_)
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
        col_border <- purrr::map_int(merge, utils::tail, 1)
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

sqlite_wait_unlock <- function(f, connection, ...) {

  execute_sql_safe <- purrr::safely(f)

  execute_sql <- execute_sql_safe(connection, ...)

  if (!is.null(execute_sql$error)) {

    error_message <- execute_sql$error

    if (stringr::str_detect(error_message, "database is locked", negate = TRUE)) {
      DBI::dbDisconnect(connection)
      stop(error_message, call. = FALSE)
    }

    while(stringr::str_detect(error_message, "database is locked")) {

      Sys.sleep(1)

      execute_sql <- execute_sql_safe(connection, ...)

      if (!is.null(execute_sql$error)) {
        error_message <- execute_sql$error
      } else {
        error_message <- "ok"
      }

    }

  } else {
    return(execute_sql$result)
  }

}
