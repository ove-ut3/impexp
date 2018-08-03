#' Import Microsoft Excel files located in a path.
#'
#' @param pattern A regular expression. Only file names matching the regular expression will be imported.
#' @param path Path where the excel files are located (recursive).
#' @param pattern_sheet A regular expression. Only sheet names in excel files matching the regular expression will be imported.
#' @param skip Inherits from \code{readxl::read_excel}.
#' @param na Inherits from \code{readxl::read_excel}.
#' @param col_types Inherits from \code{readxl::read_excel}.
#' @param parallel If \code{TRUE}, a excel files are imported using all CPU cores.
#' @param zip If \code{TRUE} then excel files within zip files are also imported.
#' @param message If \code{TRUE} then a message indicates how many files are imported.
#'
#' @return A data frame whith a column-list "import" containing all tibbles.
#'
#' @examples
#' impexp::excel_import_path(path = paste0(find.package("impexp"), "/extdata"), pattern_sheet = "impexp")
#' impexp::excel_import_path(path = paste0(find.package("impexp"), "/extdata"), pattern_sheet = "impexp", zip = TRUE)
#'
#' @export
excel_import_path <- function(pattern = "\\.xlsx?$", path = ".", pattern_sheet = ".", skip = 0, na = "", col_types = NULL, parallel = FALSE, zip = FALSE, message = TRUE) {

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
      readxl::read_excel(import$file, import$sheet, skip = skip, na = na, col_types = col_types)
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

#' Creer un onglet dans un fichier excel
#'
#' Créer un onglet dans un fichier excel.
#'
#' @param classeur \dots
#' @param table \dots
#' @param nom_onglet \dots
#' @param notes \dots
#' @param n_colonnes_lib \dots
#'
#' @export
#' @keywords internal
excel_onglet <- function(classeur, table, nom_onglet, notes = NULL, n_colonnes_lib = 0) {

  openxlsx::addWorksheet(classeur, nom_onglet)

  #n_colonnes_lib <- length(stringr::str_subset(colnames(table), "^lib"))

  if (!is.null(table[["sous_titre_"]])) {
    num_ligne_sous_titre <- which(table$sous_titre_ == "O")
    table <- dplyr::select(table, -sous_titre_)
  }

  if (!is.null(table[["indentation_"]])) {
    table <- table %>%
      dplyr::mutate(lib = paste0(purrr::map_chr(indentation_, ~ paste0(rep("    ", . - 1), collapse = "")), lib)) %>%
      dplyr::select(-indentation_)
  }

  if (!is.null(table[["dont_"]])) {
    num_ligne_dont <- which(table$dont_ == "O")
    table <- dplyr::select(table, -dont_)
  }

  if (!is.null(table[["dont_derniere_ligne_"]])) {
    num_ligne_dont_derniere_ligne <- which(table$dont_derniere_ligne_ == "O")
    table <- dplyr::select(table, -dont_derniere_ligne_)
  }

  style_donnees_numerique <- openxlsx::createStyle(halign = "right")
  style_titre1 <- openxlsx::createStyle(textDecoration = "bold", halign = "center")
  style_titre2 <- openxlsx::createStyle(textDecoration = "bold", halign = "center", border = "bottom")
  style_sous_titre <- openxlsx::createStyle(textDecoration = "bold", border = "bottom")
  style_dont <- openxlsx::createStyle(textDecoration = "italic")
  style_dont_derniere_ligne <- openxlsx::createStyle(textDecoration = "italic", border = "bottom")
  style_colonnes_bordure <- openxlsx::createStyle(border = "right")

  #### Titre ####

  titre_complet <- colnames(table) %>%
    paste0("##")

  if (n_colonnes_lib != 0){
    titre_complet[1:n_colonnes_lib] <- ""
  }

  num_ligne_titre = 0

  while(any(stringr::str_detect(titre_complet, "##"), na.rm = TRUE)) {

    num_ligne_titre <- num_ligne_titre + 1

    titre_ligne <- titre_complet %>%
      stringr::str_match("^(.+?)##") %>% .[, 2] %>%
      dplyr::tibble(titre = .)

    titre_complet <- titre_complet %>%
      stringr::str_match("##(.+)") %>% .[, 2]

    if (num_ligne_titre == 1) {
      titre_complet_bck <- titre_ligne$titre
    } else {
      titre_complet_bck <- paste(titre_complet_bck, titre_ligne$titre, sep = "##")
    }

    if (any(stringr::str_detect(titre_complet, "##"), na.rm = TRUE)) {
      merge <- titre_ligne %>%
        dplyr::mutate(titre_complet_bck = titre_complet_bck) %>%
        dplyr::mutate(merge = dplyr::row_number()) %>%
        split(x = .$merge, f = .$titre_complet_bck)
      derniere_ligne_titre <- FALSE
      style <- style_titre1

      if (num_ligne_titre == 1) {
        num_colonnes_bordure <- purrr::map_int(merge, tail, 1)
      }

    } else {
      derniere_ligne_titre <- TRUE
      style <- style_titre2
    }

    titre_ligne <- titre_ligne %>%
      t() %>%
      dplyr::as_tibble()

    start_col <- ifelse(n_colonnes_lib == 0, 1, n_colonnes_lib)

    openxlsx::writeData(classeur, nom_onglet, titre_ligne, startCol = start_col, startRow = num_ligne_titre, colNames = FALSE)
    openxlsx::addStyle(classeur, nom_onglet, style, rows = num_ligne_titre, 1:ncol(table), gridExpand = TRUE, stack = TRUE)

    if (derniere_ligne_titre == FALSE) {
      purrr::walk(merge, ~ openxlsx::mergeCells(classeur, nom_onglet, cols = ., rows = num_ligne_titre))
    }

  }

  #### Corps ####

  openxlsx::writeData(classeur, nom_onglet, table, startRow = num_ligne_titre + 1, colNames = FALSE)

  if (n_colonnes_lib != 0) {
    openxlsx::addStyle(classeur, nom_onglet, style_donnees_numerique, rows = (num_ligne_titre + 1):(num_ligne_titre + nrow(table)), n_colonnes_lib + 1:(ncol(table) + 1), gridExpand = TRUE, stack = TRUE)

  } else {
    champs_numeriques <- purrr::map_int(table, ~ class(.) %in% c("integer", "double")) %>%
    { which(. != 0) }

    if (length(champs_numeriques) != 0) {
      openxlsx::addStyle(classeur, nom_onglet, style_donnees_numerique, rows = (num_ligne_titre + 1):(num_ligne_titre + nrow(table)), champs_numeriques, gridExpand = TRUE, stack = TRUE)
    }

  }


  # Largeur de colonne auto-ajustée
  if (n_colonnes_lib != 0) {
    openxlsx::setColWidths(classeur, nom_onglet, cols = 1:n_colonnes_lib, widths = "auto")
  } else {
    openxlsx::setColWidths(classeur, nom_onglet, cols = 1:ncol(table), widths = "auto")
  }

  if (exists("num_colonnes_bordure")) {
    openxlsx::addStyle(classeur, nom_onglet, style_colonnes_bordure, rows = 1:(num_ligne_titre + nrow(table)), cols = num_colonnes_bordure, gridExpand = TRUE, stack = TRUE)

  }

  if (exists("num_ligne_sous_titre")) {
    openxlsx::addStyle(classeur, nom_onglet, style_sous_titre, rows = num_ligne_titre + num_ligne_sous_titre, cols = 1:ncol(table), gridExpand = TRUE, stack = TRUE)

  }

  if (exists("num_ligne_dont")) {
    openxlsx::addStyle(classeur, nom_onglet, style_dont, rows = num_ligne_titre + num_ligne_dont, cols = 1:ncol(table), gridExpand = TRUE, stack = TRUE)

  }

  if (exists("num_ligne_dont_derniere_ligne")) {
    openxlsx::addStyle(classeur, nom_onglet, style_dont_derniere_ligne, rows = num_ligne_titre + num_ligne_dont_derniere_ligne, cols = 1:ncol(table), gridExpand = TRUE, stack = TRUE)

  }

  #### Note ####

  if (!is.null(notes)) {
    note <- dplyr::tibble(note1 = notes)
    openxlsx::writeData(classeur, nom_onglet, note, startRow = num_ligne_titre + nrow(table) + 2, colNames = FALSE)
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

  purrr::pwalk(list(data, sheet, footer), impexp::excel_onglet, workbook = workbook, n_cols_rowname = n_cols_rowname)

  if (create_dir == TRUE) {
    stringr::str_match(path, "(.+)/[^/]+?$")[, 2] %>%
      dir.create(showWarnings = FALSE, recursive = TRUE)
  }

  openxlsx::saveWorkbook(workbook, path, overwrite = TRUE)

  return(path)
}
