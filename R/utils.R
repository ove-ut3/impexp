access_connect <- function(path) {

  if (!file.exists(path)) {
    stop(paste0("The Access database\" ", path, "\" does not exist"), call. = FALSE)
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

zip_extract <- function(zip_file, pattern = NULL, exdir = NULL, remove_zip = FALSE) {

  if (!file.exists(zip_file)) {
    stop("The zip file \"", zip_file,"\" does not exist", call. = FALSE)
  }

  files <- unzip(zip_file, list = TRUE)[["Name"]]

  if (!is.null(pattern)) files <- files[stringr::str_detect(files, pattern)]

  if (is.null(exdir)) {
    exdir <- stringr::str_match(zip_file, "(.+)/")[, 2]
  }

  unzip(zip_file, files, exdir = exdir)

  if (remove_zip == TRUE) {
    file.remove(zip_file) %>%
      invisible()
  }
}

zip_extract_path <- function(path, pattern, pattern_zip = "\\.zip$", n_files = Inf, parallel = FALSE) {

  if (!dir.exists(path)) {
    stop("The path \"", path,"\" does not exist", call. = FALSE)
  }

  zip_files <- dplyr::tibble(zip_file = list.files(path, recursive = TRUE, full.names = TRUE) %>%
                               stringr::str_subset(pattern_zip))

  if (nrow(zip_files) == 0) {
    message("There is no zip file in the path : \"", path,"\"")
    return(invisible(NULL))
  }

  zip_files <- zip_files %>%
    dplyr::mutate(id_zip = dplyr::row_number() %>% as.character())

  zip_files <- purrr::map_df(zip_files$zip_file, unzip, list = TRUE, .id = "id_zip") %>%
    dplyr::select(id_zip, file = Name) %>%
    dplyr::filter(stringr::str_detect(file, pattern)) %>%
    dplyr::inner_join(zip_files, ., by = "id_zip") %>%
    dplyr::select(-id_zip)

  if (nrow(zip_files) > abs(n_files)) {

    if (n_files > 0) {
      zip_files <- dplyr::filter(zip_files, dplyr::row_number() <= n_files)
    } else if (n_files < 0) {
      zip_files <- zip_files %>%
        dplyr::filter(dplyr::row_number() > n() + n_files)
    }

  }

  message("Extraction from ", length(zip_files$zip_file), " zip file(s)...")

  if (parallel == TRUE) {
    cluster <- parallel::makeCluster(parallel::detectCores())
  } else {
    cluster <- NULL
  }

  decompression <- zip_files %>%
    split(1:nrow(.)) %>%
    pbapply::pblapply(function(ligne) {

      zip_extract(ligne$zip_file, pattern = pattern)

    }, cl = cluster)

  if (parallel == TRUE) {
    parallel::stopCluster(cluster)
  }

  zip_files <- zip_files %>%
    dplyr::mutate(exdir = stringr::str_match(zip_file, "(.+)/")[, 2],
                  file = paste0(exdir, "/", file) %>%
                    iconv(from = "UTF-8"))

  return(zip_files)
}
