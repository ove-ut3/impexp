#' Import csv files located in a path.
#'
#' @param pattern A regular expression. Only file names matching the regular expression will be imported.
#' @param path Path where the csv files are located (recursive).
#' @param n_csv Number of csv files to extract. A negative value will starts from the bottom of the files list.
#' @param parallel If \code{TRUE} then csv files are imported using all CPU cores (using parallel & pbapply packages).
#' @param zip If \code{TRUE} then csv files within zip files are also imported.
#' @param pattern_zip A regular expression. Only zip files matching the regular expression will be extracted.
#' @param progress_bar If \code{TRUE} then a progress bar is displayed (using pbapply package).
#' @param message If \code{TRUE} then a message indicates how many files are imported.
#' @param \dots Optional arguments from \code{data.table::fread} function.
#'
#' @return A data frame with a column-list "import" containing all tibbles.
#'
#' @export
csv_import_path <- function(pattern, path = ".", n_csv = Inf, parallel = FALSE, zip = FALSE, pattern_zip = "\\.zip$", progress_bar = FALSE, message = FALSE, ...) {

  if (!dir.exists(path)) {
    stop("The path \"", path,"\" does not exist", call. = FALSE)
  }

  files <- dplyr::tibble(file = list.files(path, recursive = TRUE, full.names = TRUE) %>%
                           stringr::str_subset(pattern) %>%
                           iconv(from = "UTF-8") %>%
                           purrr::map_chr(tools::file_path_as_absolute),
                         zip_file = NA_character_)

  # If zip files are included
  if (zip == TRUE) {

    zip_files <- zip_extract_path(path, pattern = pattern, pattern_zip = pattern_zip, n_files = n_csv, parallel = parallel, progress_bar = progress_bar, message = message, files = TRUE)

    files <- dplyr::bind_rows(zip_files, files) %>%
      dplyr::arrange(file)

  }

  if (nrow(files) > abs(n_csv)) {

    if (n_csv > 0) {
      files <- dplyr::filter(files, dplyr::row_number() <= n_csv)
    } else if (n_csv < 0) {
      files <- dplyr::filter(files, dplyr::row_number() > n() + n_csv)
    }

  }

  if (nrow(files) == 0) {
    message("No files matches the parameters")

    return(files)
  }

  if (progress_bar == TRUE & !"pbapply" %in% installed.packages()[, 1]) {
    stop("pbapply package needs to be installed", call. = FALSE)
  }

  if (message == TRUE) {
    message(length(unique(files$file))," csv file(s) imported...")
  }

  if (parallel == TRUE & !all(c("parallel", "pbapply") %in% installed.packages()[, 1])) {
    stop("parallel and pbapply packages need to be installed", call. = FALSE)
  }

  if (parallel == TRUE) {
    cluster <- parallel::makeCluster(parallel::detectCores())
  } else {
    cluster <- NULL
  }

  if (progress_bar == TRUE | parallel == TRUE) {

    csv_import_path <- files %>%
      dplyr::mutate(import = pbapply::pblapply(split(., 1:nrow(.)), function(import) {

        import(import$zip_file, import$file, "csv")

      }, cl = cluster))

  } else {

    csv_import_path <- files %>%
      dplyr::mutate(import = lapply(split(., 1:nrow(.)), function(import) {

        import(import$zip_file, import$file, "csv")

      }))

  }

  if (parallel == TRUE) {
    parallel::stopCluster(cluster)
  }

  return(csv_import_path)
}
