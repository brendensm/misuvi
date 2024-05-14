#'View the MI-SUVI Documentation
#'
#'This function opens up a PDF of the MI-SUVI documentation in your web browser.
#'
#'@importFrom utils browseURL
#'@importFrom curl has_internet
#'
#'@examples
#'documentation()
#'
#'@export documentation
#'@returns NULL, opens a web browser of the technical documentation.


documentation <- function(){

  if(!curl::has_internet()){
    message("Documentation could not be opened. No internet connection.")
    return(invisible(NULL))
  }

  utils::browseURL("https://www.michigan.gov/opioids/-/media/Project/Websites/opioids/documents/edc32Michigan-2022-SUVI-Documentation-562024.pdf?rev=3cd9b9477c194f3fb616292157918cc2")

}
