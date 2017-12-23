#' Importer un fichier CSV
#'
#' Importer un fichier CSV.
#'
#' @param fichier Chemin vers le fichier CSV
#' @param ligne_debut Ligne de début à partir duquel importer.
#' @param encoding Encodage du fichier CSV.
#' @param na Caractères à considérer comme vide en plus de \code{c("NA", "", " ")}.
#' @param dec Cractère de décimale, par défaut la virgule.
#' @param col_types Type des champs (utilisé par \code{data.table::fread}.
#' @param warning_type Affichage des warnings liés aux types des champs importés.
#'
#' @return Un data frame.\cr
#'
#' @export
csv_importer <- function(fichier, ligne_debut = 1, encoding = "Latin-1", na = NULL, dec = ",", col_types = NULL, warning_type = TRUE) {

  if (!file.exists(fichier)) {
    stop("Le fichier \"", fichier,"\" n'existe pas.", call. = FALSE)
  }

  if (warning_type == FALSE) {
    fonction_import <- purrr::quietly(data.table::fread)
  } else {
    fonction_import <- data.table::fread
  }

  csv_importer <- fonction_import(iconv(fichier, from = "UTF-8"), sep = ";", encoding = encoding, na.strings = c("NA", "", " ", na), dec = dec)

  if (warning_type == FALSE) {
    if (any(!stringr::str_detect(csv_importer$warnings, "Bumped column \\d+? to type"))) {
      warnings <- csv_importer$warnings[!stringr::str_detect(csv_importer$warnings, "Bumped column \\d+? to type")]
      purrr::walk(warnings, message)
    }
    csv_importer <- csv_importer[["result"]]
  }

  csv_importer <- csv_importer %>%
    impexp::normaliser_nom_champs() %>%
    tibble::as_tibble()

  return(csv_importer)
}

#' Importer les fichiers CSV d'un repertoire (recursif)
#'
#' Importer les fichiers CSV d'un répertoire (récursif).
#'
#' @param regex_fichier Expression régulière à partir de laquelle les fichiers dont le nom matche sont importés.
#' @param chemin Chemin du répertoire à partir duquel seront importés les fichiers excel (récursif).
#' @param ligne_debut Ligne de début à partir duquel importer.
#' @param encoding Encodage des fichiers CSV.
#' @param na Caractères à considérer comme vide en plus de \code{c("NA", "", " ")}.
#' @param col_types Type des champs (utilisé par \code{data.table::fread}).
#' @param paralleliser \code{TRUE}, import parallelisé des fichiers CSV.
#' @param archive_zip \code{TRUE}, les fichiers CSV contenus dans des archives zip sont également importés; \code{FALSE} les archives zip sont ignorées.
#' @param regex_zip Expression régulière pour filtrer les archives zip à traiter.
#' @param warning_type Affichage des warnings liés aux types des champs importés.
#' @param message_import \code{TRUE}, affichage du message d'import
#'
#' @return Un data frame dont le champ "import" est la liste des data frame importés.
#'
#' @export
csv_importer_masse <- function(regex_fichier, chemin = ".", ligne_debut = 1, encoding = "Latin-1", na = NULL, col_types = NULL, paralleliser = FALSE, archive_zip = FALSE, regex_zip = NULL, warning_type = FALSE, message_import = TRUE) {

  if (!dir.exists(chemin)) {
    stop("Le répertoire \"", chemin,"\" n'existe pas.", call. = FALSE)
  }

  fichiers <- dplyr::tibble(fichier = list.files(chemin, recursive = TRUE, full.names = TRUE) %>%
                              stringr::str_subset(regex_fichier) %>%
                              iconv(from = "UTF-8"))

  # Si l'on inclut les archives zip
  if (archive_zip == TRUE) {

    archives_zip <- divr::extraire_masse_zip(chemin, regex_fichier = regex_fichier, regex_zip = regex_zip, paralleliser = paralleliser)

    fichiers <- dplyr::bind_rows(archives_zip, fichiers) %>%
      dplyr::arrange(fichier)

  } else {
    fichiers <- fichiers %>%
      dplyr::mutate(archive_zip = NA_character_)
  }

  if (nrow(fichiers) == 0) {
    message("Aucun fichier ne correspond aux paramètres saisis")
    return(NULL)
  }

  if (message_import == TRUE) {
    message("Import de ", length(unique(fichiers$fichier))," fichier(s) csv...")
    pbapply::pboptions(type = "timer")
  } else {
    pbapply::pboptions(type = "none")
  }

  if (paralleliser == TRUE) {
    cluster <- divr::initialiser_cluster()
  } else {
    cluster <- NULL
  }

  csv_importer_masse <- dplyr::tibble(fichier = unique(fichiers$fichier),
                                      import = pbapply::pblapply(unique(fichiers$fichier), csv_importer, ligne_debut = ligne_debut, encoding = encoding, na = na, col_types = col_types, warning_type = warning_type, cl = cluster))

  dplyr::filter(fichiers, !is.na(archive_zip)) %>%
    dplyr::pull(fichier) %>%
    file.remove() %>%
    invisible()

  if (paralleliser == TRUE) {
    divr::stopper_cluster(cluster)
  }

  if (archive_zip == TRUE) {
    csv_importer_masse <- fichiers %>%
      dplyr::select(-repertoire_sortie) %>%
      dplyr::left_join(csv_importer_masse, by = "fichier") %>%
      dplyr::mutate(fichier = stringr::str_match(fichier, "/(.+)")[, 2])

  }

  return(csv_importer_masse)
}
