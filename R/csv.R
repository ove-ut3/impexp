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
    csv_importer <- impexp::normaliser_nom_champs(csv_importer)
  }

  csv_importer <- dplyr::as_tibble(csv_importer)

  return(csv_importer)
}

#' Importer les fichiers CSV d'un repertoire (recursif)
#'
#' Importer les fichiers CSV d'un répertoire (récursif).
#'
#' @param regex_fichier Expression régulière à partir de laquelle les fichiers dont le nom matche sont importés.
#' @param chemin Chemin du répertoire à partir duquel seront importés les fichiers excel (récursif).
#' @param fonction Fonction à utiliser pour l'import CSV.
#' @param ligne_debut Ligne de début à partir duquel importer.
#' @param encoding Encodage des fichiers CSV.
#' @param na Caractères à considérer comme vide en plus de \code{c("NA", "", " ")}.
#' @param col_types Type des champs (utilisé par \code{data.table::fread}).
#' @param normaliser Normaliser les noms de champ de table.
#' @param n_csv Nombre de CSV à importer. Une valeur négative correspond au nombre de CSV à partir de la fin dans la liste.
#' @param paralleliser \code{TRUE}, import parallelisé des fichiers CSV.
#' @param archive_zip \code{TRUE}, les fichiers CSV contenus dans des archives zip sont également importés; \code{FALSE} les archives zip sont ignorées.
#' @param regex_zip Expression régulière pour filtrer les archives zip à traiter.
#' @param warning_type Affichage des warnings liés aux types des champs importés.
#' @param message_import \code{TRUE}, affichage du message d'import
#'
#' @return Un data frame dont le champ "import" est la liste des data frame importés.
#'
#' @export
csv_importer_masse <- function(regex_fichier, chemin = ".", fonction = "read.csv2", ligne_debut = 1, encoding = "Latin-1", na = NULL, col_types = NULL, normaliser = TRUE, n_csv = Inf, paralleliser = FALSE, archive_zip = FALSE, regex_zip = "\\.zip$", warning_type = FALSE, message_import = TRUE) {

  if (!dir.exists(chemin)) {
    stop("Le répertoire \"", chemin,"\" n'existe pas.", call. = FALSE)
  }

  fichiers <- dplyr::tibble(fichier = list.files(chemin, recursive = TRUE, full.names = TRUE) %>%
                              stringr::str_subset(regex_fichier) %>%
                              iconv(from = "UTF-8"),
                            archive_zip = NA_character_)

  # Si l'on inclut les archives zip
  if (archive_zip == TRUE) {

    archives_zip <- divr::zip_extract_path(chemin, regex_fichier = regex_fichier, regex_zip = regex_zip, n_fichiers = n_csv, paralleliser = paralleliser)

    fichiers <- dplyr::bind_rows(archives_zip, fichiers) %>%
      dplyr::arrange(fichier)

  }

  if (nrow(fichiers) > abs(n_csv)) {

    if (n_csv > 0) {
      fichiers <- dplyr::filter(fichiers, dplyr::row_number() <= n_csv)
    } else if (n_csv < 0) {
      fichiers <- fichiers %>%
        dplyr::filter(dplyr::row_number() > n() + n_csv)
    }

  }

  if (nrow(fichiers) == 0) {
    message("Aucun fichier ne correspond aux paramètres saisis")

    return(fichiers)
  }

  if (message_import == TRUE) {
    message("Import de ", length(unique(fichiers$fichier))," fichier(s) csv...")
    pbapply::pboptions(type = "timer")
  } else {
    pbapply::pboptions(type = "none")
  }

  if (paralleliser == TRUE) {
    cluster <- divr::cl_initialise()
  } else {
    cluster <- NULL
  }

  csv_importer_masse <- dplyr::tibble(fichier = unique(fichiers$fichier),
                                      import = pbapply::pblapply(unique(fichiers$fichier), impexp::csv_importer, fonction = fonction, ligne_debut = ligne_debut, encoding = encoding, na = na, col_types = col_types, normaliser = normaliser, warning_type = warning_type, cl = cluster))

  suppression <- dplyr::filter(fichiers, !is.na(archive_zip)) %>%
    dplyr::pull(fichier) %>%
    file.remove()

  if (paralleliser == TRUE) {
    divr::cl_stop(cluster)
  }

  if (archive_zip == TRUE) {
    csv_importer_masse <- fichiers %>%
      dplyr::select(-repertoire_sortie) %>%
      dplyr::left_join(csv_importer_masse, by = "fichier") %>%
      dplyr::mutate(fichier = stringr::str_match(fichier, "/(.+)")[, 2])

  }

  return(csv_importer_masse)
}
