#'Add shape files to your data
#'
#'This function queries shape files from the tigris package. This function returns an sf object.
#'
#'@param x a data frame for which you want to add shape files to. Must be one of the MI-SUVI data sets.
#'
#'@importFrom tigris zctas counties
#'@importFrom sf st_as_sf
#'
#'@examples
#'my_data <- misuvi_load() |> add_geometry()
#'
#'
#'@export add_geometry
#'@returns an sf data.frame of MI-SUVI data.

add_geometry <- function(x){

  if(nrow(x) > 900){

    # ZCTA

    geom <- tigris::zctas(state = "MI", year = 2010)[,c("ZCTA5CE10", "geometry")]

    final <- merge(x, geom, by.x = "zcta", by.y = "ZCTA5CE10") |>
      sf::st_as_sf()

    return(final)

  }else{

    geom <- tigris::counties(state = "MI", year = 2020, cb = TRUE)[,"GEOID", "geometry"]

    final <- merge(x, geom, by.x = "fips", by.y = "GEOID", all = TRUE) |>
      sf::st_as_sf()

    return(final)

  }

}
