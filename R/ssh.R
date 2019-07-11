#' ssh_import_file
#'
#' @param path \dots
#' @param ssh_key_file \dots
#' @param to \dots
#'
#' @export
ssh_import_file <- function(path, ssh_key_file, to = ".") {

  session <- ssh::ssh_connect("root@shiny.le-henanff.fr", keyfile = ssh_key_file)
  ssh::scp_download(session, path, to = to, verbose = FALSE)
  ssh::ssh_disconnect(session)

  file <- stringr::str_extract(path, "([^/]+?)$")

  path <- paste0(to, "/", file) %>%
    tools::file_path_as_absolute()

  return(path)
}
