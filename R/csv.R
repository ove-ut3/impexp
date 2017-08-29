#' Importer un fichier CSV
#'
#' Importer un fichier CSV.
#'
#' @param fichier Chemin vers le fichier CSV
#' @param ligne_debut Ligne de début à partir duquel importer.
#' @param encoding Encodage du fichier CSV.
#' @param col_types Type des champs (utilisé par \code{data.table::fread}.
#'
#' @return Un data frame.\cr
#'
#' @export
importer_fichier_csv <- function(fichier, ligne_debut = 1, encoding = "Latin-1", col_types = NULL) {

  importer_fichier_csv <- data.table::fread(fichier, sep = ";", encoding = encoding) %>%
    importr::normaliser_nom_champs()

  return(importer_fichier_csv)
}

#' Importer les fichiers CSV d'un repertoire (recursif)
#'
#' Importer les fichiers CSV d'un répertoire (récursif).
#'
#' @param regex_fichier Expression régulière à partir de laquelle les onglet dont le nom matche sont importés.
#' @param chemin Chemin du répertoire à partir duquel seront importés les fichiers excel (récursif).
#' @param regex_onglet Expression régulière à partir de laquelle les onglet dont le nom matche sont importés.
#' @param ligne_debut Ligne de début à partir duquel importer.
#' @param encoding Encodage des fichiers CSV.
#' @param col_types Type des champs (utilisé par \code{data.table::fread}.
#' @param paralleliser \code{TRUE}, import parallelisé des fichiers excel.
#' @param archive_zip \code{TRUE}, les fichiers excel contenus dans des archives zip sont également importés; \code{FALSE} les archives zip sont ignorées.
#' @param message_import Le message à afficher pendant l'importation (par défaut : "Import des fichiers csv:")
#'
#' @return Un data frame dont le champ "import" est la liste des data frame importés.
#'
#' @export
importer_masse_csv <- function(regex_fichier, chemin = ".", regex_onglet = ".", ligne_debut = 1, encoding = "Latin-1", col_types = NULL, paralleliser = FALSE, archive_zip = FALSE, message_import = "Import des fichiers csv:") {

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

  message(message_import)

  if (paralleliser == TRUE) {
    cluster <- divr::initialiser_cluster()
  } else {
    cluster <- NULL
  }

  importer_masse_csv <- dplyr::tibble(fichier = unique(fichiers$fichier),
                                      import = pbapply::pblapply(unique(fichiers$fichier), importer_fichier_csv, ligne_debut = ligne_debut, encoding = encoding, col_types = col_types, cl = cluster))

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
