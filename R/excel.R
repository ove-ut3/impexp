#' Lister les onglets d'un fichier excel
#'
#' Lister les onglets d'un fichier excel.
#'
#' @param fichier Chemin vers le fichier excel.
#'
#' @return un vecteur de type caractère contenant les noms des onglets du fichier excel.
#'
#' @examples
#' importr::liste_onglets_excel(paste0(racine_packages, "importr/inst/extdata/importr.xlsx"))
#'
#' @export
liste_onglets_excel <- function(fichier) {

  if (!file.exists(fichier)) {
    stop(paste0("Le fichier \"", fichier, "\" n'existe pas"), call. = FALSE)
  }

  extension <- divr::extension_fichier(fichier)

  if (!stringr::str_detect(extension, "^xls[xm]?$")) {
    stop(paste0("Le fichier \"", fichier, "\" n'est pas un fichier excel"), call. = FALSE)
  }

  quiet_excel_sheets <- purrr::quietly(readxl::excel_sheets)
  liste_onglets <- quiet_excel_sheets(fichier) %>%
    .[["result"]]

  if (any(class(liste_onglets) == "character") == TRUE) {
    return(liste_onglets)
  }

  if (extension == "xls") {

    chargement_excel <- tryCatch( {
      xlsx::loadWorkbook(fichier)
      }
      , error = function(cond) {
      return(cond)
      }
    )

    if (any(class(chargement_excel) == "error") == TRUE) {
      return(NULL)
    }

    liste_onglets <- xlsx::getSheets(chargement_excel) %>%
      names() %>%
      iconv(from = "UTF-8") %>%
      trimws()

    return(liste_onglets)

  } else {
    liste_onglets <- openxlsx::getSheetNames(fichier)

    return(liste_onglets)
  }
}

#' Importer un fichier Excel
#'
#' Importer un fichier Excel.
#'
#' @param fichier Chemin vers le fichier excel.
#' @param nom_onglet Nom de l'onglet à importer.
#' @param regex_onglet Expression régulière à partir de laquelle les onglet dont le nom matche sont importés.
#' @param num_onglet Numéro de l'onglet à importer.
#' @param ligne_debut Ligne de début à partir duquel importer.
#' @param col_types Type des champs (utilisé par \code{readxl::read_excel}.
#' @param test_champ_manquant Un nom de champ du fichier. Importer un fichier Excel sans connaitre la ligne de début, mais à partir de la première ligne non-vide du champ.
#'
#' @return Un data frame ou une liste de data frames.\cr
#'
#' Si les paramètres \code{nom_onglet} et \code{regex_onglet} sont laissés à NULL, alors le premier onglet du fichier excel est importé.\cr
#' Si le paramètre \code{nom_onglet} est rempli, alors cet onglet est importé sans prendre en compte le paramètre \code{regex_onglet}.\cr
#' Si le paramètre \code{regex_onglet} est rempli et le champ \code{nom_onglet} laissé à vide, alors les onglets dont le nom matche avec l'expression régulière sont importés dans une liste "import" stockée dans un data frame.
#'
#' @examples
#' # Import du premier onglet
#' importr::importer_fichier_excel(paste0(racine_packages, "importr/inst/extdata/importr.xlsx"))
#'
#' # Import de l'onglet "Onglet 2"
#' importr::importer_fichier_excel(paste0(racine_packages, "importr/inst/extdata/importr.xlsx"),
#'   nom_onglet = "Onglet 2")
#'
#' # Import des onglets comprenant le mot "Autre" en début de chaine
#' importr::importer_fichier_excel(paste0(racine_packages, "importr/inst/extdata/importr.xlsx"),
#'   regex_onglet = "^Autre")
#'
#' @export
importer_fichier_excel <- function(fichier, nom_onglet = NULL, regex_onglet = NULL, num_onglet = 1, ligne_debut = 1, col_types = NULL, test_champ_manquant = NULL) {

  if (!file.exists(fichier)) {
    stop(paste0("Le fichier \"", fichier, "\" n'existe pas"), call. = FALSE)
  }

  extension <- divr::extension_fichier(fichier)

  if (!stringr::str_detect(extension, "^xls[xm]?$")) {
    stop(paste0("Le fichier \"", fichier, "\" n'est pas un fichier excel"), call. = FALSE)
  }

  noms_onglets <- liste_onglets_excel(fichier)

  if (is.null(noms_onglets)) {
    if (!is.null(regex_onglet)) {
      import <- dplyr::tibble(fichier = fichier,
                                import = list(NULL))
      attr(import, "erreur") <- "Chargement des onglets impossibles"
      return(import)
    } else {
      message("Fichier ", fichier, " : Pas d'onglet dans le fichier (NULL)")
      return(NULL)
    }
  }

  if (!is.null(nom_onglet)) {
    num_onglet <- which(noms_onglets == nom_onglet)
  } else if (!is.null(regex_onglet)) {
    num_onglet <- stringr::str_detect(noms_onglets, regex_onglet) %>%
      which()
  }

  if (length(num_onglet) == 0) {
    if (!is.null(regex_onglet)) {
      import <- dplyr::tibble(fichier = fichier,
                                import = list(NULL))
      attr(import, "erreur") <- paste0("Pas d'onglet matchant le pattern: ", regex_onglet[1])
      return(import)
    } else {
      message("Fichier ", fichier, " : Pas d'onglet matchant le nom: ", nom_onglet)
      return(NULL)
    }
  }

  if (!is.null(regex_onglet)) {

    import <- dplyr::tibble(fichier = fichier,
                             onglet = noms_onglets[num_onglet],
                             import = purrr::map(num_onglet, importer_fichier_excel_,
                               fichier = fichier,
                               ligne_debut = ligne_debut,
                               col_types = col_types,
                               test_champ_manquant = test_champ_manquant)
                             )

    return(import)
  }

  import <- importr::importer_fichier_excel_(fichier, num_onglet, ligne_debut, col_types, test_champ_manquant)

  return(import)

}

#' Importer un fichier Excel (fonction générique).
#'
#' @param fichier Chemin vers le fichier excel.
#' @param num_onglet Numéro de l'onglet à importer.
#' @param ligne_debut Ligne de début à partir duquel importer.
#' @param col_types Type des champs (utilisé par \code{readxl::read_excel}.
#' @param test_champ_manquant Un nom de champ du fichier. Importer un fichier Excel sans connaitre la ligne de début, mais à partir de la première ligne non-vide du champ.
#'
#' @return Un data frame correspondant à la feuille excel.
#'
#' @export
#' @keywords internal
importer_fichier_excel_ <- function(fichier, num_onglet, ligne_debut = 1, col_types = NULL, test_champ_manquant = NULL) {

  quiet_read_excel <- purrr::quietly(readxl::read_excel)
  import <- quiet_read_excel(fichier, sheet = num_onglet, skip = ligne_debut - 1, col_types = col_types) %>%
    .[["result"]]

  if (any(class(import) == "tbl_df") == TRUE) {

    if (nrow(import) == 0) {
      import <- dplyr::tibble("warning")
      attr(import, "warning") <- "Pas de ligne dans l'onglet"
    } else {

      import <- importr::normaliser_nom_champs(import)

      if (!is.null(test_champ_manquant)) {
        import <- import_test_champ_manquant(table = import, test_champ_manquant = test_champ_manquant, fichier = fichier, num_onglet = num_onglet, diff_ligne_package = 1)
      }
    }

    attr(import, "package") <- "readxl"

    return(import)
  }

  extension <- divr::extension_fichier(fichier)

  if (extension == "xls"){

    import <- tryCatch(
      {
        xlsx::read.xlsx(fichier, sheetIndex = num_onglet, startRow = ligne_debut, check.names = TRUE, stringsAsFactors = FALSE, encoding="UTF-8")
      }
      , error = function(cond) {
        return(cond)
      }
    )

    if (class(import)[1] == "data.frame") {

      import <- dplyr::as_tibble(import) %>%
        normaliser_nom_champs() %>%
        caracteres_vides_na()

      if (!is.null(test_champ_manquant)) {
        import <- import_test_champ_manquant(table = import, test_champ_manquant = test_champ_manquant, fichier = fichier, num_onglet = num_onglet, diff_ligne_package = 3)
      }

    } else {

      import <- dplyr::tibble("erreur")
      attr(import, "erreur") <- "Impossible de lire le fichier excel"
    }

    attr(import, "package") <- "xlsx"

    return(import)

  } else {

    import <- tryCatch(
      {
        openxlsx::read.xlsx(fichier, sheet = num_onglet, startRow = ligne_debut, check.names = TRUE)
      }
      , error = function(cond) {
        return(cond)
      }
    )

    if (any(class(import) == "data.frame") == TRUE) {

      import <- dplyr::as_tibble(import) %>%
        normaliser_nom_champs() %>%
        caracteres_vides_na()

      if (!is.null(test_champ_manquant)) {
        import <- import_test_champ_manquant(table = import, test_champ_manquant = test_champ_manquant, fichier = fichier, num_onglet = num_onglet, diff_ligne_package = 3)
      }
    } else {
      erreur <- import$message
      import <- dplyr::tibble("erreur")
      attr(import, "erreur") <- erreur
    }

    attr(import, "package") <- "openxlsx"

    return(import)
  }

}

#' Importer un fichier Excel sans connaitre la ligne de début, mais à partir de la première ligne non-vide d'un champ (fonction générique).
#'
#' @param table Table pour laquelle doit être déterminée la première ligne non-vide.
#' @param test_champ_manquant Un nom de champ du fichier. Importer un fichier Excel sans connaitre la ligne de début, mais à partir de la première ligne non-vide du champ.
#' @param fichier Chemin vers le fichier excel.
#' @param num_onglet Numéro de l'onglet à importer.
#' @param diff_ligne_package Correction entre packages sur le paramètre de première ligne d'import.
#' @param col_types Type des champs (utilisé par \code{readxl::read_excel}.
#'
#' @return Un data frame importé à partir de la première ligne non-vide.
#'
#' @export
#' @keywords internal
import_test_champ_manquant <- function(table, test_champ_manquant, fichier, num_onglet, diff_ligne_package, col_types) {

  #Test si un champ est attendu dans l'import mais n'est pas présent

  test_champ_manquant <- caractr::normaliser_char(test_champ_manquant)

  if (intersect(colnames(table), test_champ_manquant) %>% length() == 1) {
    # Le champ est trouvé
    return(table)
  }

  test <- importr::importer_fichier_excel_(fichier, num_onglet, ligne_debut = 1, col_types)
  colnames(test) <- paste0("champ_", as.character(1:ncol(test)))

  normaliser_char <- caractr::normaliser_char

  ligne_en_tete <- dplyr::mutate_all(test, .funs = "as.character") %>%
    dplyr::mutate_all(.funs = "normaliser_char") %>%
    dplyr::mutate(num_ligne = row_number() + diff_ligne_package)

  ligne_en_tete$paste2 <- apply(ligne_en_tete[, colnames(ligne_en_tete)], 1, caractr::paste2, collapse = "##")
  ligne_en_tete <- dplyr::select(ligne_en_tete, num_ligne, paste2) %>%
    dplyr::mutate(paste2 = paste0("##", paste2, "##")) %>%
    dplyr::filter(stringr::str_detect(paste2, fixed(paste0("##", test_champ_manquant, "##")))) %>%
    .$num_ligne

  if (length(ligne_en_tete) != 1) {

    import <- dplyr::tibble("erreur")
    attr(import, "erreur") <- paste0("Fichier ", fichier, " / Onglet ", num_onglet, " : Tentative de nouvel import mais la valeur '", test_champ_manquant, "' n'a pas ete trouvée dans la première colonne")

    return(import)
  }

  import <- importr::importer_fichier_excel_(fichier, num_onglet, ligne_debut = ligne_en_tete, col_types) %>%
    importr::normaliser_nom_champs() %>%
    importr::caracteres_vides_na()

  attr(import, "info") <- paste0("Fichier ", fichier, " / Onglet ", num_onglet, " : Nouvel import a partir de la ligne ", ligne_en_tete)

  return(import)
}

#' Importer les fichiers Excel d'un repertoire (recursif)
#'
#' Importer les fichiers Excel d'un répertoire (récursif).
#'
#' @param chemin Chemin du répertoire à partir duquel seront importés les fichiers excel (récursif).
#' @param regex_fichier Expression régulière à partir de laquelle les onglet dont le nom matche sont importés.
#' @param regex_onglet Expression régulière à partir de laquelle les onglet dont le nom matche sont importés.
#' @param ligne_debut Ligne de début à partir duquel importer.
#' @param col_types Type des champs (utilisé par \code{readxl::read_excel}.
#' @param fichier_log Chemin du fichier de log.
#' @param archive_zip \code{TRUE}, les fichiers excel contenus dans des archives zip sont également importés; \code{FALSE} les archives zip sont ignorées.
#' @param test_champ_manquant Un nom de champ du fichier. Importer un fichier Excel sans connaitre la ligne de début, mais à partir de la première ligne non-vide du champ.
#' @param message_import Le message à afficher pendant l'importation (par défaut : "Import des fichiers excels:")
#'
#' @return Un data frame dont le champ "import" est la liste des data frame importés.
#'
#' @examples
#' importr::importer_masse_xlsx(paste0(racine_packages, "importr/inst/extdata"), regex_fichier = "xlsx$", regex_onglet = "importr")
#'
#' @export
importer_masse_excel <- function(regex_fichier, chemin = ".", regex_onglet = ".", ligne_debut = 1, col_types = NULL, paralleliser = FALSE, archive_zip = FALSE, test_champ_manquant = NULL, archive_zip_repertoire_sortie = "import_masse_excel", message_import = "Import des fichiers excels:") {

  fichiers <- dplyr::tibble(fichier = list.files(chemin, recursive = TRUE, full.names = TRUE) %>%
                              .[which(stringr::str_detect(., regex_fichier))])

  # Si l'on inclut les archives zip
  if (archive_zip == TRUE) {

    divr::vider_repertoire(archive_zip_repertoire_sortie)

    archives_zip <- divr::extraire_masse_zip(chemin, regex_fichier = regex_fichier, repertoire_sortie = archive_zip_repertoire_sortie, paralleliser = paralleliser)

    fichiers <- dplyr::bind_rows(archives_zip, fichiers) %>%
      arrange(fichier)
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

  import_masse_xlsx <- pbapply::pblapply(fichiers$fichier %>% unique, importer_fichier_excel, regex_onglet = regex_onglet, ligne_debut = ligne_debut, col_types = col_types, test_champ_manquant = test_champ_manquant, cl = cluster) %>%
    dplyr::bind_rows() %>%
    dplyr::mutate(package = purrr::map(import, attributes) %>% purrr::map_chr( ~ ifelse(!is.null(.$package), .$package, NA_character_)),
                  erreur = purrr::map(import, attributes) %>% purrr::map_chr( ~ ifelse(!is.null(.$erreur), .$erreur, NA_character_)),
                  warning = purrr::map(import, attributes) %>% purrr::map_chr( ~ ifelse(!is.null(.$warning), .$warning, NA_character_)),
                  info = purrr::map(import, attributes) %>% purrr::map_chr( ~ ifelse(!is.null(.$info), .$info, NA_character_))
           )

  if (paralleliser == TRUE) {
    divr::stopper_cluster(cluster)
  }

  if (archive_zip == TRUE) {
    import_masse_xlsx <- left_join(fichiers, import_masse_xlsx, by = "fichier") %>%
      dplyr::mutate(fichier = stringr::str_match(fichier, "/(.+)")[, 2])

  }

  return(import_masse_xlsx)

}
