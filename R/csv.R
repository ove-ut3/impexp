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
