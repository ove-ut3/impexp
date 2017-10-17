#' Importer un fichier CSV
#'
#' Importer un fichier CSV.
#'
#' @param fichier Chemin vers le fichier CSV
#' @param ligne_debut Ligne de début à partir duquel importer.
#' @param encoding Encodage du fichier CSV.
#' @param na Caractères à considérer comme vide en plus de \code{c("NA", "", " ")}.
#' @param col_types Type des champs (utilisé par \code{data.table::fread}.
#' @param warning_type Affichage des warnings liés aux types des champs importés.
#'
#' @return Un data frame.\cr
#'
#' @export
importer_fichier_csv <- function(fichier, ligne_debut = 1, encoding = "Latin-1", na = NULL, col_types = NULL, warning_type = TRUE) {

  if (!file.exists(fichier)) {
    stop("Le fichier \"", fichier,"\" n'existe pas.", call. = FALSE)
  }

  if (warning_type == FALSE) {
    fonction_import <- purrr::quietly(data.table::fread)
  } else {
    fonction_import <- data.table::fread
  }

  importer_fichier_csv <- fonction_import(iconv(fichier, from = "UTF-8"), sep = ";", encoding = encoding, na.strings = c("NA", "", " ", na))

  if (warning_type == FALSE) {
    if (any(!stringr::str_detect(importer_fichier_csv$warnings, "Bumped column \\d+? to type"))) {
      warnings <- importer_fichier_csv$warnings[!stringr::str_detect(importer_fichier_csv$warnings, "Bumped column \\d+? to type")]
      purrr::walk(warnings, message)
    }
    importer_fichier_csv <- importer_fichier_csv[["result"]]
  }

  importer_fichier_csv <- importer_fichier_csv %>%
    tibble::as_tibble() %>%
    importr::normaliser_nom_champs()

  return(importer_fichier_csv)
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
#' @param warning_type Affichage des warnings liés aux types des champs importés.
#' @param message_import \code{TRUE}, affichage du message d'import
#'
#' @return Un data frame dont le champ "import" est la liste des data frame importés.
#'
#' @export
importer_masse_csv <- function(regex_fichier, chemin = ".", ligne_debut = 1, encoding = "Latin-1", na = NULL, col_types = NULL, paralleliser = FALSE, archive_zip = FALSE, warning_type = FALSE, message_import = TRUE) {

  if (!dir.exists(chemin)) {
    stop("Le répertoire \"", chemin,"\" n'existe pas.", call. = FALSE)
  }

  fichiers <- dplyr::tibble(fichier = list.files(chemin, recursive = TRUE, full.names = TRUE) %>%
                              .[which(stringr::str_detect(., regex_fichier))])

  # Si l'on inclut les archives zip
  if (archive_zip == TRUE) {

    archives_zip <- divr::extraire_masse_zip(chemin, regex_fichier = regex_fichier, paralleliser = paralleliser)

    fichiers <- dplyr::bind_rows(archives_zip, fichiers) %>%
      dplyr::arrange(fichier)
  }

  if (nrow(fichiers) == 0) {
    message("Aucun fichier ne correspond aux paramètres saisis")
    return(NULL)
  }

  if (message_import) message("Import de ", length(unique(fichiers$fichier))," fichier(s) csv...")

  if (paralleliser == TRUE) {
    cluster <- divr::initialiser_cluster()
  } else {
    cluster <- NULL
  }

  importer_masse_csv <- dplyr::tibble(fichier = unique(fichiers$fichier),
                                      import = pbapply::pblapply(unique(fichiers$fichier), importer_fichier_csv, ligne_debut = ligne_debut, encoding = encoding, na = na, col_types = col_types, warning_type = warning_type, cl = cluster))

  file.remove(fichiers$fichier) %>%
    invisible()

  if (paralleliser == TRUE) {
    divr::stopper_cluster(cluster)
  }

  if (archive_zip == TRUE) {
    importer_masse_csv <- fichiers %>%
      dplyr::select(-repertoire_sortie) %>%
      dplyr::left_join(importer_masse_csv, by = "fichier") %>%
      dplyr::mutate(fichier = stringr::str_match(fichier, "/(.+)")[, 2])

  }

  return(importer_masse_csv)

}
