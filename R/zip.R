#' Extract files from a zip.
#'
#' Files within the zip file can be selected with the pattern argument.
#'
#' @param zip_file Path to the zip file.
#' @param pattern an optional regular expression. Only file names which match the regular expression will be extracted.
#' @param exdir The directory to extract files to. It will be created if necessary.
#' @param remove_zip If \code{TRUE}, the zip file is removed.
#'
#' @examples
#' impexp::zip_extract("archive.zip", pattern = "pdf$", exdir = "test_zip")
#'
#' @export
#' @keywords internal
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

#' Extract files from zip located in a path.
#'
#' @param path Path where the zip files are located (recursive).
#' @param pattern an optional regular expression. Only file names which match the regular expression will be extracted.
#' @param pattern_zip an optional regular expression. Only zip file names which match the regular expression will be unziped.
#' @param n_files Number of files to extract. A negative value will starts from the bottom of the files list.
#' @param parallel If \code{TRUE}, a parallelised extraction is performed.
#'
#' @examples
#' impexp::zip_extract_path("Chemin/vers/un/répartoire", pattern = "pdf$")
#'
#' @export
#' @keywords internal
zip_extract_path <- function(path, pattern, pattern_zip = "\\.zip$", n_files = Inf, parallel = FALSE) {

  if (!dir.exists(path)) {
    stop("The path \"", path,"\" does not exist", call. = FALSE)
  }

  if (parallel == TRUE) {
    clusters <- divr::cl_initialise()
  } else {
    clusters <- NULL
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

  decompression <- zip_files %>%
    split(1:nrow(.)) %>%
    pbapply::pblapply(function(ligne) {

      impexp::zip_extract(ligne$zip_file, pattern = pattern)

    }, cl = clusters)

  if (parallel == TRUE) {
    divr::cl_stop(clusters)
  }

  zip_files <- zip_files %>%
    dplyr::mutate(exdir = stringr::str_match(zip_file, "(.+)/")[, 2],
                  file = paste0(exdir, "/", file))

  return(zip_files)
}
