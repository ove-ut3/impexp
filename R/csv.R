#' Importer un fichier CSV
#'
#' Importer un fichier CSV.
#'
#' @param fichier Chemin vers le fichier CSV
#' @param fonction Fonction à utiliser pour l'import CSV.
#' @param ligne_debut Ligne de début à partir duquel importer.
#' @param encoding Encodage du fichier CSV.
#' @param na Caractères à considérer comme vide en plus de \code{c("NA", "", " ")}.
#' @param dec Cractère de décimale, par défaut la virgule.
#' @param col_types Type des champs (utilisé par \code{data.table::fread}.
#' @param normaliser Normaliser les noms de champ de table.
#' @param warning_type Affichage des warnings liés aux types des champs importés.
#'
#' @return Un data frame.\cr
#'
#' @export
#' @keywords internal
csv_importer <- function(fichier, fonction = "read.csv2", ligne_debut = 1, encoding = "Latin-1", na = NULL, dec = ",", col_types = NULL, normaliser = TRUE, warning_type = TRUE) {

  if (!file.exists(fichier)) {
    stop("Le fichier \"", fichier,"\" n'existe pas.", call. = FALSE)
  }

  if (fonction == "read.csv2") {
    encoding <- dplyr::recode(encoding, "Latin-1" = "Windows-1252")
    csv_importer <- read.csv2(iconv(fichier, from = "UTF-8"), na.strings = c("NA", "", " ", na), fileEncoding = encoding, check.names = FALSE)

  } else if (fonction == "fread") {

    if (warning_type == FALSE) {
      fonction_import <- purrr::quietly(data.table::fread)
    } else {
      fonction_import <- data.table::fread
    }

    csv_importer <- fonction_import(iconv(fichier, from = "UTF-8"), sep = ";", encoding = encoding, na.strings = c("NA", "", na), dec = dec)

    if (warning_type == FALSE) {
      if (any(!stringr::str_detect(csv_importer$warnings, "Bumped column \\d+? to type"))) {
        warnings <- csv_importer$warnings[!stringr::str_detect(csv_importer$warnings, "Bumped column \\d+? to type")]
        purrr::walk(warnings, message)
      }
      csv_importer <- csv_importer[["result"]]
    }

  }

  if (normaliser == TRUE) {
    csv_importer <- patchr::normalise_colnames(csv_importer)
  }

  csv_importer <- dplyr::as_tibble(csv_importer)

  return(csv_importer)
}

#' Import csv files located in a path.
#'
#' @param pattern A regular expression. Only file names matching the regular expression will be imported.
#' @param path Path where the csv files are located (recursive).
#' @param n_csv Number of csv files to extract. A negative value will starts from the bottom of the files list.
#' @param parallel If \code{TRUE} then csv files are imported using all CPU cores.
#' @param zip If \code{TRUE} then csv files within zip files are also imported.
#' @param pattern_zip A regular expression. Only zip files matching the regular expression will be extracted.
#' @param message If \code{TRUE} then a message indicates how many files are imported.
#' @param \dots Optional arguments from \code{data.table::fread} function.
#'
#' @return A data frame whith a column-list "import" containing all tibbles.
#'
#' @export
csv_import_path <- function(pattern, path = ".", n_csv = Inf, parallel = FALSE, zip = FALSE, pattern_zip = "\\.zip$", message = TRUE, ...) {

  if (!dir.exists(path)) {
    stop("The path \"", path,"\" does not exist", call. = FALSE)
  }

  files <- dplyr::tibble(file = list.files(path, recursive = TRUE, full.names = TRUE) %>%
                           stringr::str_subset(pattern) %>%
                           iconv(from = "UTF-8"),
                         zip_file = NA_character_)

  # If zip files are included
  if (zip == TRUE) {

    zip_files <- impexp::zip_extract_path(path, pattern = pattern, pattern_zip = pattern_zip, n_files = n_csv, parallel = parallel) %>%
      dplyr::select(-exdir)

    files <- dplyr::bind_rows(zip_files, files) %>%
      dplyr::arrange(file)

  }

  if (nrow(files) > abs(n_csv)) {

    if (n_csv > 0) {
      files <- dplyr::filter(files, dplyr::row_number() <= n_csv)
    } else if (n_csv < 0) {
      files <- files %>%
        dplyr::filter(dplyr::row_number() > n() + n_csv)
    }

  }

  if (nrow(files) == 0) {
    message("No files matches the parameters")

    return(files)
  }

  if (message == TRUE) {
    message(length(unique(files$file))," csv files imported...")
    pbapply::pboptions(type = "timer")
  } else {
    pbapply::pboptions(type = "none")
  }

  if (parallel == TRUE) {
    cluster <- parallel::makeCluster(parallel::detectCores())
  } else {
    cluster <- NULL
  }

  csv_import_path <- files %>%
    dplyr::mutate(import = pbapply::pblapply(split(., 1:nrow(.)), function(import) {
      data.table::fread(import$file, ...) %>%
        dplyr::as_tibble()
    }, cl = cluster))

  remove <- dplyr::filter(files, !is.na(zip_file)) %>%
    dplyr::pull(file) %>%
    file.remove()

  if (parallel == TRUE) {
    parallel::stopCluster(cluster)
  }

  if (zip == TRUE) {
    csv_import_path <- files %>%
      dplyr::left_join(csv_import_path, by = c("zip_file", "file")) %>%
      dplyr::mutate(file = stringr::str_match(file, "/(.+)")[, 2])

  }

  return(csv_import_path)
}
