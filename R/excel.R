#' Importer un fichier Excel
#'
#' Importer un fichier Excel.
#'
#' @param fichier Chemin vers le fichier excel.
#' @param nom_onglet Nom de l'onglet à importer.
#' @param regex_onglet Expression régulière à partir de laquelle les onglet dont le nom matche sont importés.
#' @param num_onglet Numéro de l'onglet à importer.
#' @param ligne_debut Ligne de début à partir duquel importer.
#' @param na Caractères à considérer comme vide en plus de \code{c("")}.
#' @param col_types Type des champs (utilisé par \code{readxl::read_excel}.
#' @param normaliser Normaliser les noms de champ de table.
#'
#' @return Un data frame ou une liste de data frames.\cr
#'
#' Si les paramètres \code{nom_onglet} et \code{regex_onglet} sont laissés à NULL, alors le premier onglet du fichier excel est importé.\cr
#' Si le paramètre \code{nom_onglet} est rempli, alors cet onglet est importé sans prendre en compte le paramètre \code{regex_onglet}.\cr
#' Si le paramètre \code{regex_onglet} est rempli et le champ \code{nom_onglet} laissé à vide, alors les onglets dont le nom matche avec l'expression régulière sont importés dans une liste "import" stockée dans un data frame.
#'
#' @examples
#' # Import du premier onglet
#' impexp::excel_importer(paste0(racine_packages, "impexp/inst/extdata/impexp.xlsx"))
#'
#' # Import de l'onglet "Onglet 2"
#' impexp::excel_importer(paste0(racine_packages, "impexp/inst/extdata/impexp.xlsx"),
#'   nom_onglet = "Onglet 2")
#'
#' # Import des onglets comprenant le mot "Autre" en début de chaine
#' impexp::excel_importer(paste0(racine_packages, "impexp/inst/extdata/impexp.xlsx"),
#'   regex_onglet = "^Autre")
#'
#' @export
excel_importer <- function(fichier, nom_onglet = NULL, regex_onglet = NULL, num_onglet = 1, ligne_debut = 1, na = NULL, col_types = NULL, normaliser = TRUE) {

  if (!file.exists(fichier)) {
    stop("Le fichier \"", fichier,"\" n'existe pas.", call. = FALSE)
  }

  noms_onglets <- readxl::excel_sheets(fichier)

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

    num_onglet <- stringr::str_which(noms_onglets, regex_onglet)

    import <- dplyr::tibble(fichier = fichier,
                            onglet = noms_onglets[num_onglet])

    import$import <- lapply(num_onglet, impexp::excel_importer_,
                            fichier = fichier,
                            ligne_debut = ligne_debut,
                            na = na,
                            col_types = col_types,
                            normaliser = normaliser)

  } else {

    import <- impexp::excel_importer_(fichier = fichier,
                                      num_onglet = num_onglet,
                                      ligne_debut = ligne_debut,
                                      na = na,
                                      col_types = col_types,
                                      normaliser = normaliser)
  }

  return(import)
}

#' Importer un fichier Excel (fonction generique)
#'
#' Importer un fichier Excel (fonction générique).
#'
#' @param fichier Chemin vers le fichier excel.
#' @param num_onglet Numéro de l'onglet à importer.
#' @param ligne_debut Ligne de début à partir duquel importer.
#' @param na Caractères à considérer comme vide en plus de \code{c("")}.
#' @param col_types Type des champs (utilisé par \code{readxl::read_excel}.
#' @param normaliser Normaliser les noms de champ de table.
#'
#' @return Un data frame ou une liste de data frames.\cr
#'
#' @export
#' @keywords internal
excel_importer_ <- function(fichier, num_onglet = 1, ligne_debut = 1, na = NULL, col_types = NULL, normaliser = TRUE) {

  quiet_read_excel <- purrr::quietly(readxl::read_excel)

  import <- quiet_read_excel(fichier, sheet = num_onglet, skip = ligne_debut - 1, na = c("", na), col_types = col_types) %>%
    .[["result"]]

  if (any(class(import) == "tbl_df") == TRUE) {

    if (nrow(import) == 0) {
      import <- dplyr::tibble("warning")
      attr(import, "warning") <- "Pas de ligne dans l'onglet"

    }

    if (normaliser == TRUE) {
      import <- impexp::normaliser_nom_champs(import)
    }

  } else {
    import <- dplyr::tibble("erreur")
    attr(import, "erreur") <- "Impossible de lire le fichier excel"

  }

  return(import)
}

#' Importer les fichiers Excel d'un repertoire (recursif)
#'
#' Importer les fichiers Excel d'un répertoire (récursif).
#'
#' @param regex_fichier Expression régulière à partir de laquelle les onglet dont le nom matche sont importés.
#' @param chemin Chemin du répertoire à partir duquel seront importés les fichiers excel (récursif).
#' @param regex_onglet Expression régulière à partir de laquelle les onglet dont le nom matche sont importés.
#' @param ligne_debut Ligne de début à partir duquel importer.
#' @param na Caractères à considérer comme vide en plus de \code{c("")}.
#' @param col_types Type des champs (utilisé par \code{readxl::read_excel}.
#' @param normaliser Normaliser les noms de champ de table.
#' @param paralleliser \code{TRUE}, import parallelisé des fichiers excel.
#' @param archive_zip \code{TRUE}, les fichiers excel contenus dans des archives zip sont également importés; \code{FALSE} les archives zip sont ignorées.
#' @param message_import \code{TRUE}, affichage du message d'import
#'
#' @return Un data frame dont le champ "import" est la liste des data frame importés.
#'
#' @examples
#' impexp::importer_masse_xlsx(paste0(racine_packages, "impexp/inst/extdata"), regex_fichier = "xlsx$", regex_onglet = "impexp")
#'
#' @export
excel_importer_masse <- function(regex_fichier, chemin = ".", regex_onglet = ".", ligne_debut = 1, na = NULL, col_types = NULL, normaliser = TRUE, paralleliser = FALSE, archive_zip = FALSE, message_import = TRUE) {

  if (!dir.exists(chemin)) {
    stop("Le répertoire \"", chemin,"\" n'existe pas.", call. = FALSE)
  }

  fichiers <- dplyr::tibble(fichier = list.files(chemin, recursive = TRUE, full.names = TRUE) %>%
                              stringr::str_subset(regex_fichier) %>%
                              iconv(from = "UTF-8"))

  # Si l'on inclut les archives zip
  if (archive_zip == TRUE) {

    archives_zip <- impexp::zip_extract_path(chemin, regex_fichier = regex_fichier, paralleliser = paralleliser)

    fichiers <- dplyr::bind_rows(archives_zip, fichiers) %>%
      arrange(fichier)

  } else {
    fichiers <- fichiers %>%
      dplyr::mutate(archive_zip = NA_character_)
  }

  if (nrow(fichiers) == 0) {
    message("Aucun fichier ne correspond aux paramètres saisis")

    return(fichiers)
  }

  if (message_import == TRUE) {
    message("Import de ", length(unique(fichiers$fichier))," fichiers excel...")
    pbapply::pboptions(type = "timer")
  } else {
    pbapply::pboptions(type = "none")
  }

  if (paralleliser == TRUE) {
    cluster <- divr::cl_initialise()
  } else {
    cluster <- NULL
  }

  import_masse_xlsx <- pbapply::pblapply(fichiers$fichier %>% unique, excel_importer, regex_onglet = regex_onglet, ligne_debut = ligne_debut, na = na, col_types = col_types, normaliser = normaliser, cl = cluster) %>%
    dplyr::bind_rows() %>%
    dplyr::mutate(erreur = lapply(import, attributes) %>%
                    purrr::map_chr( ~ ifelse(!is.null(.$erreur), .$erreur, NA_character_)),
                  warning = lapply(import, attributes) %>%
                    purrr::map_chr( ~ ifelse(!is.null(.$warning), .$warning, NA_character_)),
                  info = lapply(import, attributes) %>%
                    purrr::map_chr( ~ ifelse(!is.null(.$info), .$info, NA_character_))
           )

  suppression <- dplyr::filter(fichiers, !is.na(archive_zip)) %>%
    dplyr::pull(fichier) %>%
    file.remove()

  if (paralleliser == TRUE) {
    divr::cl_stop(cluster)
  }

  if (archive_zip == TRUE) {
    import_masse_xlsx <- dplyr::left_join(fichiers, import_masse_xlsx, by = "fichier") %>%
      dplyr::mutate(fichier = stringr::str_match(fichier, "/(.+)")[, 2])

  }

  return(import_masse_xlsx)

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

#' Exporter un fichier excel
#'
#' Exporter un fichier excel.
#'
#' @param table data.frame ou liste de data.frame à exporter.
#' @param nom_fichier \dots
#' @param creer_repertoire \dots
#' @param nom_onglet \dots
#' @param notes \dots
#' @param n_colonnes_lib \dots
#'
#' @export
excel_exporter <- function(table, nom_fichier, creer_repertoire = FALSE, nom_onglet = NULL, notes = NULL, n_colonnes_lib = 0) {

  if (any(class(table) == "data.frame")) {
    table <- list("table" = table)
  }

  if (any(purrr::map_lgl(table, ~ !any(class(.) == "data.frame")))) {
    stop("Au moins un des objets n'est pas un data.frame", call. = FALSE)
  }

  if (is.null(nom_onglet)) {
    nom_onglet <- names(table)
  }

  classeur <- openxlsx::createWorkbook()

  if (length(notes) == 1) {
    notes <- rep(notes, length(table)) %>% as.list()

  } else if (!is.null(notes) & length(notes) != length(table)) {
    stop("Le nombre de notes doit être égale au nombre de table", call. = FALSE)

  } else if (is.null(notes)) {
    notes <- rep(NA_character_, length(table)) %>% as.list()
  }

  purrr::pwalk(list(table, nom_onglet, notes), impexp::excel_onglet, classeur = classeur, n_colonnes_lib = n_colonnes_lib)

  if (creer_repertoire == TRUE) {
    stringr::str_match(nom_fichier, "(.+)/[^/]+?$")[, 2] %>%
      dir.create(showWarnings = FALSE, recursive = TRUE)
  }

  openxlsx::saveWorkbook(classeur, nom_fichier, overwrite = TRUE)

  return(nom_fichier)
}
