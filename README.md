# Access the 'Michigan Substance Use Vulnerablility Index' with misuvi

This package allows users to easily access the MI-SUVI data sets. The user can import data sets with full metrics, percentiles, Z-scores, or rankings. Data is available at both the County and Zip Code Tabulation Area (ZCTA) levels.

"The MI-SUVI was created to consider the diverse factors that influence a community’s vulnerability related to substance use. The MI-SUVI is a single, standardized score that considers multiple factors that influence a community’s vulnerability related to substance use, including indicators related to substance use burden, resources, and social vulnerability." (MDHHS, 2024)

## Installation

To install from CRAN:
```r
install.packages("misuvi")
```

Or install from GitHub:
```r
devtools::install_github("brendensm/misuvi")
```

## Main Functions
-   `misuvi_load` &mdash; Loads the MI-SUVI data set of your choice. Defaults to County metrics.
-   `add_geometry` &mdash; Adds shape files to your MI-SUVI data.
-   `documentation` &mdash; Launches a browser with the full technical documentation of the MI-SUVI data.
-   `dictionary` &mdash; Provides a data frame with a more detailed name of the abbreviated variables that `misuvi_load` returns.

## Usage

### Loading Data

By default, the `misuvi_load` function returns the raw metrics of Michigan Counties without the margin of error variables.

```r
library(misuvi)

misuvi_load(geography = "county", type = "metrics", moe = "FALSE")
```

To return ZCTAs, simply change geography to "zcta". Likewise we can choose to return ranks, zscores, or percentile ranks too.

```r
misuvi_load("zcta", type = "ranks")

misuvi_load("zcta", type = "zcores")

misuvi_load("zcta", type = "percentiles")
```

## References
Michigan Department of Health and Human Services. (2024). Michigan 2022 Substance Use Vulnerability Index Documentation(Version 1). https://www.michigan.gov/opioids/-/media/Project/Websites/opioids/documents/edc32Michigan-2022-SUVI-Documentation-562024.pdf?rev=3cd9b9477c194f3fb616292157918cc2.

