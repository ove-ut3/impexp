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

#' @keywords internal
zip_extract_path <- function(path, pattern, pattern_zip = "\\.zip$", n_files = Inf, parallel = FALSE) {

  if (!dir.exists(path)) {
    stop("The path \"", path,"\" does not exist", call. = FALSE)
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

  if (parallel == TRUE) {
    cluster <- parallel::makeCluster(parallel::detectCores())
  } else {
    cluster <- NULL
  }

  decompression <- zip_files %>%
    split(1:nrow(.)) %>%
    pbapply::pblapply(function(ligne) {

      zip_extract(ligne$zip_file, pattern = pattern)

    }, cl = cluster)

  if (parallel == TRUE) {
    parallel::stopCluster(cluster)
  }

  zip_files <- zip_files %>%
    dplyr::mutate(exdir = stringr::str_match(zip_file, "(.+)/")[, 2],
                  file = paste0(exdir, "/", file) %>%
                    iconv(from = "UTF-8"))

  return(zip_files)
}
