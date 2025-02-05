#'Load MI-SUVI Data
#'
#'This function allows the user to select the MI-SUVI data set they want returned.
#'
#'@param geography Choose the geographic level you want returned to you. "county" or "zcta".
#'@param type Select the type of data frame you would like returned. Options include "metrics", "percentiles", "zscores", and "ranks".
#'@param moe Logical -- choose whether margin of error variables are returned. This applies to the "metrics" data sets only.
#'
#'@examples
#'zcta_metrics <- misuvi_load("zcta", "metrics")
#'county_zscores <- misuvi_load("county", "zscores")
#'
#'
#'@export misuvi_load
#'@returns A clean data.frame of MI-SUVI data.

misuvi_load <- function(geography = "county", type = "zscores", moe = FALSE){

  if(!(type %in% c("metrics", "percentiles", "zscores", "ranks"))){
    stop("Please select a type from the following options: 'metrics', 'percentiles', 'zscores', 'ranks'")
  }

  if(!(geography %in% c("county", "zcta"))){
    stop("Please select a geography of 'county' or 'zcta'.")
  }

  if(type == "metrics"){

    if(isFALSE(moe)){

      if(geography == "county"){

        re_county <- county_metrics[, !grepl("^moe", names(county_metrics))]

        return(re_county)

      }else{
        re_zcta <- zcta_metrics[, !grepl("^moe", names(zcta_metrics))]

        return(re_zcta)
      }

    }else{
      if(geography == "county"){
        return(county_metrics)
      }else{
        return(zcta_metrics)
      }
    }

  }else if(type == "percentiles"){

    if(geography == "county"){
      return(county_percentiles)
    }else{
      return(zcta_percentiles)
    }
  }else if(type == "zscores"){

    if(geography == "county"){
      return(county_zscores)
    }else{
      return(zcta_zscores)
    }
  }else{

    if(geography == "county"){
      return(county_ranks)
    }else{
      return(zcta_ranks)
    }
  }

}
