# Data wrangling for MISUVI

library(readxl)
library(janitor)
library(dplyr)
library(utils)

# Last Downloaded: 5/8/2024

# Percentiles -------------------------------------------------------------
headers_text <- c(
  `user-agent` = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/102.0.5005.61 Safari/537.36'
)

temp_c <- tempfile(fileext = ".xlsx")
utils::download.file("https://www.michigan.gov/opioids/-/media/Project/Websites/opioids/documents/db00a2022-MI-County-Substance-Use-Vulnerability-Index-Results-Web-Version-V1Updated-4302024.xlsx?rev=a072a21b48c045c1b00fc20d07e93891",
                     temp_c, mode = "wb", headers = headers_text)

temp_z <- tempfile(fileext = ".xlsx")
utils::download.file("https://www.michigan.gov/opioids/-/media/Project/Websites/opioids/documents/2ef3b2022-MI-ZCTA-ZIP-Code-Substance-Use-Vulnerability-Index-Results-Web-Version--V1Updated-562024.xlsx?rev=6fc15a677e4b48ed8258d824bbb74e53",
                     temp_z, mode = "wb", headers = headers_text)

county_p <- readxl::read_xlsx(temp_c, sheet = 6, skip = 1)[-c(1:2),] |>
  clean_names() |>
  rename(
    fatal_od = x5_year_average_fatal_overdose_rate_per_100_000_2018_2022,
    od_ed = x3_year_average_nonfatal_overdose_emergency_healthcare_visit_rate_per_100_000_2020_2022,
    op_script = opioid_prescription_unit_rate_per_1_000_2022,
    drug_arrest = drug_related_arrest_rate_per_100_000_2022,
    sud_30min = percent_of_population_within_30_minute_drive_of_sud_treatment_center_2022,
    ssp_15min = percent_of_population_within_15_minute_drive_of_syringe_service_program_2022,
    bup_script = buprenorphine_prescription_unit_rate_per_1_000_2022,
    mod_svi_score = modified_social_vulnerability_index_score_2022,
    burden_score = mi_suvi_burden_score_2022,
    resource_score = mi_suvi_resource_score_2022,
    svi_score = mi_suvi_social_vulnerability_score_2022,
    misuvi_score = mi_suvi_score_2022
  )


county_p[,3:14] <- lapply(county_p[,3:14], as.numeric)

county_percentiles <- county_p

#usethis::use_data(county_percentiles, compress = "xz", overwrite = TRUE, internal = TRUE)


zcta_percentiles <- readxl::read_xlsx(temp_z, sheet = 8, skip = 1)[-c(1:2),] |>
  clean_names() |>
  rename(

    fatal_od = x5_year_average_fatal_overdose_rate_per_100_000_2018_2022,
    od_ed = x3_year_average_nonfatal_overdose_emergency_healthcare_visit_rate_per_100_000_2020_2022,
    op_script = opioid_prescription_unit_rate_per_1_000_2022,
    sud_30min = percent_of_population_within_30_minute_drive_of_sud_treatment_center_2022,
    ssp_15min = percent_of_population_within_15_minute_drive_of_syringe_service_program_2022,
    bup_script = buprenorphine_prescription_unit_rate_per_1_000_2022,
    mod_svi_score = modified_social_vulnerability_index_score_2022,
    burden_score = mi_suvi_burden_score_2022,
    resource_score = mi_suvi_resource_score_2022,
    svi_score = mi_suvi_social_vulnerability_score_2022,
    misuvi_score = mi_suvi_score_2022

  )


zcta_percentiles[,4:14] <- lapply(zcta_percentiles[,4:14], as.numeric)
zcta_percentiles[,1] <- lapply(zcta_percentiles[,1], as.character)

#usethis::use_data(zcta_percentiles, compress = "xz", overwrite = TRUE, internal = TRUE)


# Metrics -----------------------------------------------------------------



county_suvi <- readxl::read_xlsx(temp_c, sheet = 3)[3:85,] |>
  clean_names() |>
  rename(

    pov150 = percent_of_individuals_below_150_percent_poverty_estimate,
    unemployed = percent_of_civilian_population_16_unemployed,
    high_housing = percent_of_households_with_high_housing_cost_burden,
    no_hs = percent_of_individuals_with_no_high_school_diploma,
    no_insurance = percent_of_civilian_population_without_health_insurance,
    age65_older = percent_of_individuals_65,
    lessthan18 = percent_of_individuals_18,
    disability = percent_of_civilian_population_with_a_disability,
    single_parent = percent_of_households_with_single_parent_and_children_18,
    eng_less = percent_individuals_5_who_speak_english_less_than_well,
    nw_hisp = percent_of_individuals_not_white_non_hispanic,
    housing_10unit = percent_of_housing_structures_with_10_units,
    mobile_homes = percent_of_housing_units_that_are_mobile_homes,
    more_people_rooms = percent_of_households_with_more_people_than_rooms,
    no_vehicle = percent_of_households_with_no_vehicle,
    group_quarters = percent_of_individuals_living_in_group_quarters,
    no_broadband = percent_of_households_without_a_computer_with_broadband_internet,
    pharm_15min = percent_of_population_within_15_min_drive_of_pharmacy,
    hosp_30min = percent_of_population_within_30_min_drive_of_hospital)
county_suvi <- county_suvi |>
  rename(
    !!paste0("moe_", names(county_suvi)[4]) := x5,
    !!paste0("moe_", names(county_suvi)[6]) := x7,
    !!paste0("moe_", names(county_suvi)[8]) := x9,
    !!paste0("moe_", names(county_suvi)[10]) := x11,
    !!paste0("moe_", names(county_suvi)[12]) := x13,
    !!paste0("moe_", names(county_suvi)[14]) := x15,
    !!paste0("moe_", names(county_suvi)[16]) := x17,
    !!paste0("moe_", names(county_suvi)[18]) := x19,
    !!paste0("moe_", names(county_suvi)[20]) := x21,
    !!paste0("moe_", names(county_suvi)[22]) := x23,
    !!paste0("moe_", names(county_suvi)[24]) := x25,
    !!paste0("moe_", names(county_suvi)[26]) := x27,
    !!paste0("moe_", names(county_suvi)[28]) := x29,
    !!paste0("moe_", names(county_suvi)[30]) := x31,
    !!paste0("moe_", names(county_suvi)[32]) := x33,
    !!paste0("moe_", names(county_suvi)[34]) := x35,
    !!paste0("moe_", names(county_suvi)[36]) := x37,
  ) |> select(-county)

county_suvi[,3:38] <- lapply(county_suvi[,3:38], as.numeric)
county_suvi[,1] <- lapply(county_suvi[,1], as.character)

county_od <- readxl::read_xlsx(temp_c, sheet = 4, skip=1)[-c(1:2),] |>
  clean_names() |>
  rename(

    fatal_od = x5_year_average_fatal_overdose_rate_per_100_000_2018_2022,
    od_ed = x3_year_average_nonfatal_overdose_emergency_healthcare_visit_rate_per_100_000_2020_2022,
    op_script = opioid_prescription_unit_rate_per_1_000_2022,
    drug_arrest = drug_related_arrest_rate_per_100_000_2022,
    sud_30min = percent_of_population_within_30_minute_drive_of_sud_treatment_center_2022,
    ssp_15min = percent_of_population_within_15_minute_drive_of_syringe_service_program_2022,
    bup_script = buprenorphine_prescription_unit_rate_per_1_000_2022,
    mod_svi_score = modified_social_vulnerability_index_score_2022

  )

county_od[,3:10] <- lapply(county_od[,3:10], as.numeric)

county_metrics <- merge(county_od, county_suvi, by = "fips", all = TRUE) |> as_tibble()

#usethis::use_data(county_metrics, compress = "xz", overwrite = TRUE, internal = TRUE)






zcta_suvi <- readxl::read_xlsx(temp_z, sheet = 5, skip = 1)[-c(1, 970:976),] |>
  clean_names() |>
  rename(

    pov150 = percent_of_individuals_below_150_percent_poverty_estimate,
    unemployed = percent_of_civilian_population_16_unemployed,
    high_housing = percent_of_households_with_high_housing_cost_burden,
    no_hs = percent_of_individuals_with_no_high_school_diploma,
    no_insurance = percent_of_civilian_population_without_health_insurance,
    age65_older = percent_of_individuals_65,
    lessthan18 = percent_of_individuals_18,
    disability = percent_of_civilian_population_with_a_disability,
    single_parent = percent_of_households_with_single_parent_and_children_18,
    eng_less = percent_individuals_5_who_speak_english_less_than_well,
    nw_hisp = percent_of_individuals_not_white_non_hispanic,
    housing_10unit = percent_of_housing_structures_with_10_units,
    mobile_homes = percent_of_housing_units_that_are_mobile_homes,
    more_people_rooms = percent_of_households_with_more_people_than_rooms,
    no_vehicle = percent_of_households_with_no_vehicle,
    group_quarters = percent_of_individuals_living_in_group_quarters,
    no_broadband = percent_of_households_without_a_computer_with_broadband_internet,
    pharm_15min = percent_of_population_within_15_min_drive_of_pharmacy,
    hosp_30min = percent_of_population_within_30_min_drive_of_hospital)
zcta_suvi <- zcta_suvi |>
  rename(
    !!paste0("moe_", names(zcta_suvi)[5]) := x6,
    !!paste0("moe_", names(zcta_suvi)[7]) := x8,
    !!paste0("moe_", names(zcta_suvi)[9]) := x10,
    !!paste0("moe_", names(zcta_suvi)[11]) := x12,
    !!paste0("moe_", names(zcta_suvi)[13]) := x14,
    !!paste0("moe_", names(zcta_suvi)[15]) := x16,
    !!paste0("moe_", names(zcta_suvi)[17]) := x18,
    !!paste0("moe_", names(zcta_suvi)[19]) := x20,
    !!paste0("moe_", names(zcta_suvi)[21]) := x22,
    !!paste0("moe_", names(zcta_suvi)[23]) := x24,
    !!paste0("moe_", names(zcta_suvi)[25]) := x26,
    !!paste0("moe_", names(zcta_suvi)[27]) := x28,
    !!paste0("moe_", names(zcta_suvi)[29]) := x30,
    !!paste0("moe_", names(zcta_suvi)[31]) := x32,
    !!paste0("moe_", names(zcta_suvi)[33]) := x34,
    !!paste0("moe_", names(zcta_suvi)[35]) := x36,
    !!paste0("moe_", names(zcta_suvi)[37]) := x38,
  ) |> select(-c(associated_counties, associated_county_subdivisions))

zcta_suvi[,4:38] <- lapply(zcta_suvi[,4:38], as.numeric)
zcta_suvi[,1] <- lapply(zcta_suvi[,1], as.character)

zcta_od <- readxl::read_xlsx(temp_z, sheet = 6, skip = 1)[-c(1:2, 971:975),] |>
  clean_names() |>
  rename(

    fatal_od = x5_year_average_fatal_overdose_rate_per_100_000_2018_2022,
    od_ed = x3_year_average_nonfatal_overdose_emergency_healthcare_visit_rate_per_100_000_2020_2022,
    op_script = opioid_prescription_unit_rate_per_1_000_2022,
    sud_30min = percent_of_population_within_30_minute_drive_of_sud_treatment_center_2022,
    ssp_15min = percent_of_population_within_15_minute_drive_of_syringe_service_program_2022,
    bup_script = buprenorphine_prescription_unit_rate_per_1_000_2022,
    mod_svi_score = modified_social_vulnerability_index_score_2022

  )

zcta_od[,4:10] <- lapply(zcta_od[,4:10], as.numeric)

zcta_metrics <- merge(zcta_od, zcta_suvi, by = "zcta", all = TRUE)

#usethis::use_data(zcta_metrics, compress = "xz", overwrite = TRUE, internal = TRUE)


# Z-scores ----------------------------------------------------------------

county_z <- readxl::read_xlsx(temp_c, sheet = 5, skip = 1)[-c(1:2),] |>
  clean_names() |>
  rename(
    fatal_od = x5_year_average_fatal_overdose_rate_per_100_000_2018_2022,
    od_ed = x3_year_average_nonfatal_overdose_emergency_healthcare_visit_rate_per_100_000_2020_2022,
    op_script = opioid_prescription_unit_rate_per_1_000_2022,
    drug_arrest = drug_related_arrest_rate_per_100_000_2022,
    sud_30min = percent_of_population_within_30_minute_drive_of_sud_treatment_center_2022,
    ssp_15min = percent_of_population_within_15_minute_drive_of_syringe_service_program_2022,
    bup_script = buprenorphine_prescription_unit_rate_per_1_000_2022,
    mod_svi_score = modified_social_vulnerability_index_score_2022,
    burden_score = mi_suvi_burden_score_2022,
    resource_score = mi_suvi_resource_score_2022,
    svi_score = mi_suvi_social_vulnerability_score_2022,
    misuvi_score = mi_suvi_score_2022
  )


county_z[,3:14] <- lapply(county_z[,3:14], as.numeric)

county_zscores <- county_z

#usethis::use_data(county_zscores, compress = "xz", overwrite = TRUE, internal = TRUE)


zcta_zscores <- readxl::read_xlsx(temp_z, sheet = 7, skip = 1)[-c(1:2),] |>
  clean_names() |>
  rename(

    fatal_od = x5_year_average_fatal_overdose_rate_per_100_000_2018_2022,
    od_ed = x3_year_average_nonfatal_overdose_emergency_healthcare_visit_rate_per_100_000_2020_2022,
    op_script = opioid_prescription_unit_rate_per_1_000_2022,
    sud_30min = percent_of_population_within_30_minute_drive_of_sud_treatment_center_2022,
    ssp_15min = percent_of_population_within_15_minute_drive_of_syringe_service_program_2022,
    bup_script = buprenorphine_prescription_unit_rate_per_1_000_2022,
    mod_svi_score = modified_social_vulnerability_index_score_2022,
    burden_score = mi_suvi_burden_score_2022,
    resource_score = mi_suvi_resource_score_2022,
    svi_score = mi_suvi_social_vulnerability_score_2022,
    misuvi_score = mi_suvi_score_2022

  )


zcta_zscores[,4:14] <- lapply(zcta_zscores[,4:14], as.numeric)
zcta_zscores[,1] <- lapply(zcta_zscores[,1], as.character)

#usethis::use_data(zcta_zscores, compress = "xz", overwrite = TRUE, internal = TRUE)




# Rankings ----------------------------------------------------------------

county_r <- readxl::read_xlsx(temp_c, sheet = 7, skip = 1)[-c(1:2),] |>
  clean_names() |>
  rename(
    fatal_od = x5_year_average_fatal_overdose_rate_per_100_000_2018_2022,
    od_ed = x3_year_average_nonfatal_overdose_emergency_healthcare_visit_rate_per_100_000_2020_2022,
    op_script = opioid_prescription_unit_rate_per_1_000_2022,
    drug_arrest = drug_related_arrest_rate_per_100_000_2022,
    sud_30min = percent_of_population_within_30_minute_drive_of_sud_treatment_center_2022,
    ssp_15min = percent_of_population_within_15_minute_drive_of_syringe_service_program_2022,
    bup_script = buprenorphine_prescription_unit_rate_per_1_000_2022,
    mod_svi_score = modified_social_vulnerability_index_score_2022,
    burden_score = mi_suvi_burden_score_2022,
    resource_score = mi_suvi_resource_score_2022,
    svi_score = mi_suvi_social_vulnerability_score_2022,
    misuvi_score = mi_suvi_score_2022
  )


county_r[,3:14] <- lapply(county_r[,3:14], as.numeric)

county_ranks <- county_r

#usethis::use_data(county_ranks, compress = "xz", overwrite = TRUE, internal = TRUE)


zcta_ranks <- readxl::read_xlsx(temp_z, sheet = 9, skip = 1)[-c(1:2),] |>
  clean_names() |>
  rename(

    fatal_od = x5_year_average_fatal_overdose_rate_per_100_000_2018_2022,
    od_ed = x3_year_average_nonfatal_overdose_emergency_healthcare_visit_rate_per_100_000_2020_2022,
    op_script = opioid_prescription_unit_rate_per_1_000_2022,
    sud_30min = percent_of_population_within_30_minute_drive_of_sud_treatment_center_2022,
    ssp_15min = percent_of_population_within_15_minute_drive_of_syringe_service_program_2022,
    bup_script = buprenorphine_prescription_unit_rate_per_1_000_2022,
    mod_svi_score = modified_social_vulnerability_index_score_2022,
    burden_score = mi_suvi_burden_score_2022,
    resource_score = mi_suvi_resource_score_2022,
    svi_score = mi_suvi_social_vulnerability_score_2022,
    misuvi_score = mi_suvi_score_2022

  )


zcta_ranks[,4:14] <- lapply(zcta_ranks[,4:14], as.numeric)
zcta_ranks[,1] <- lapply(zcta_ranks[,1], as.character)


#### Regional Data

temp_reg <- tempfile(fileext = ".xlsx")
utils::download.file("https://www.michigan.gov/opioids/-/media/Project/Websites/opioids/documents/Public-Use-Dataset-County-Population-and-Regional-Groupings.xlsx?rev=6ac0d413e9164876b1bcd32acfc30f0c",
                     temp_reg, mode = "wb", headers = headers_text)

reg_raw <- read_xlsx(temp_reg, sheet = 2, skip = 1)

var_labels = c(
  county = "County Name",
  fips = "FIPS",
  #county_code = "County Code",
  pop18 = "2018 Population",
  pop19 = "2019 Population",
  pop20 = "2020 Population",
  pihp = "Prepaid Inpatient Health Plan (PIHP) Region",
  region = "5 Region Grouping",
  urbanicity = "3 Level Urbanicity Grouping",
  epr = "Emergency Preparedness Region",
  mspd = "Michigan State Police Districts"
)

regional_groupings <- reg_raw |>
  clean_names() |>
  rename(
    county = county_name,
    pop18 = x2018_population,
    pop19 = x2019_population,
    pop20 = x2020_population,
    pihp = prepaid_inpatient_health_plan_pihp_region,
    region = x5_region_grouping,
    urbanicity = x3_level_urbanicity_grouping,
    epr = emergency_preparedness_region,
    mspd = michigan_state_police_districts
  ) #|> mutate(county_code = substrRight(fips, 2))

Hmisc::label(regional_groupings) = as.list(var_labels[match(names(regional_groupings), names(var_labels))])



# Dictionary --------------------------------------------------------------



c_metric_label <- c(
  "FIPS",
  "County",
  "5-Year Average Fatal Overdose Rate per 100,000 (2018-2022)",
  "3-Year Average Nonfatal Overdose Emergency Healthcare Visit Rate per 100,000 (2020-2022)",
  "Opioid Prescription Unit Rate per 1,000 (2022)",
  "Drug Related Arrest Rate per 100,000 (2022)",
  "Percent of Population within 30 Minute Drive of SUD Treatment Center (2022)",
  "Percent of Population within 15 Minute Drive of Syringe Service Program (2022)",
  "Buprenorphine Prescription Unit Rate per 1,000 (2022)",
  "Modified Social Vulnerability Index Score (2022)",
  "County Population",
  "% of Individuals Below 150% Poverty Estimate",
  "% of Civilian Population 16+ Unemployed",
  "% of Households with High Housing Cost Burden",
  "% of Individuals with No High School Diploma",
  "% of Civilian Population without Health Insurance",
  "% of Individuals 65+",
  "% of Individuals <18",
  "% of Civilian Population with a Disability",
  "% of Households with Single Parent and Children <18",
  "% Individuals 5+ Who Speak English 'Less than Well'",
  "% of Individuals Not White, non-Hispanic",
  "% of Housing Structures with 10+ Units",
  "% of Housing Units that are Mobile Homes",
  "% of Households with More People Than Rooms",
  "% of Households with No Vehicle",
  "% of Individuals Living in Group Quarters",
  "% of Households Without a Computer with Broadband Internet",
  "% of Population Within 15 Min Drive of Pharmacy",
  "% of Population Within 30 Min Drive of Hospital"
)

z_metric_label <- c(
  "ZCTA",
  "Associated Counties",
  "Assoicated County Subdivisions",
  "5-Year Average Fatal Overdose Rate per 100,000 (2018-2022)",
  "3-Year Average Nonfatal Overdose Emergency Healthcare Visit Rate per 100,000 (2020-2022)",
  "Opioid Prescription Unit Rate per 1,000 (2022)",
  "Percent of Population within 30 Minute Drive of SUD Treatment Center (2022)",
  "Percent of Population within 15 Minute Drive of Syringe Service Program (2022)",
  "Buprenorphine Prescription Unit Rate per 1,000 (2022)",
  "Modified Social Vulnerability Index Score (2022)",
  "County Population",
  "% of Individuals Below 150% Poverty Estimate",
  "% of Civilian Population 16+ Unemployed",
  "% of Households with High Housing Cost Burden",
  "% of Individuals with No High School Diploma",
  "% of Civilian Population without Health Insurance",
  "% of Individuals 65+",
  "% of Individuals <18",
  "% of Civilian Population with a Disability",
  "% of Households with Single Parent and Children <18",
  "% Individuals 5+ Who Speak English 'Less than Well'",
  "% of Individuals Not White, non-Hispanic",
  "% of Housing Structures with 10+ Units",
  "% of Housing Units that are Mobile Homes",
  "% of Households with More People Than Rooms",
  "% of Households with No Vehicle",
  "% of Individuals Living in Group Quarters",
  "% of Households Without a Computer with Broadband Internet",
  "% of Population Within 15 Min Drive of Pharmacy",
  "% of Population Within 30 Min Drive of Hospital"
)



cleaned_names <-
  c(
    "fatal_od",
    "od_ed",
    "op_script",
    "drug_arrest",
    "sud_30min",
    "ssp_15min",
    "bup_script",
    "mod_svi_score",
    "burden_score",
    "resource_score",
    "svi_score",
    "misuvi_score",
    "pov150",
    "unemployed",
    "high_housing",
    "no_hs",
    "no_insurance",
    "age65_older",
    "lessthan18",
    "disability",
    "single_parent",
    "eng_less",
    "nw_hisp",
    "housing_10unit",
    "mobile_homes",
    "more_people_rooms",
    "no_vehicle",
    "group_quarters",
    "no_broadband",
    "pharm_15min",
    "hosp_30min"
  )


original_names <-
  c(
    "5_year_average_fatal_overdose_rate_per_100_000_2018_2022",
    "3_year_average_nonfatal_overdose_emergency_healthcare_visit_rate_per_100_000_2020_2022",
    "opioid_prescription_unit_rate_per_1_000_2022",
    "drug_related_arrest_rate_per_100_000_2022",
    "percent_of_population_within_30_minute_drive_of_sud_treatment_center_2022",
    "percent_of_population_within_15_minute_drive_of_syringe_service_program_2022",
    "buprenorphine_prescription_unit_rate_per_1_000_2022",
    "modified_social_vulnerability_index_score_2022",
    "mi_suvi_burden_score_2022",
    "mi_suvi_resource_score_2022",
    "mi_suvi_social_vulnerability_score_2022",
    "mi_suvi_score_2022",
    "percent_of_individuals_below_150_percent_poverty_estimate",
    "percent_of_civilian_population_16_unemployed",
    "percent_of_households_with_high_housing_cost_burden",
    "percent_of_individuals_with_no_high_school_diploma",
    "percent_of_civilian_population_without_health_insurance",
    "percent_of_individuals_65",
    "percent_of_individuals_18",
    "percent_of_civilian_population_with_a_disability",
    "percent_of_households_with_single_parent_and_children_18",
    "percent_individuals_5_who_speak_english_less_than_well",
    "percent_of_individuals_not_white_non_hispanic",
    "percent_of_housing_structures_with_10_units",
    "percent_of_housing_units_that_are_mobile_homes",
    "percent_of_households_with_more_people_than_rooms",
    "percent_of_households_with_no_vehicle",
    "percent_of_individuals_living_in_group_quarters",
    "percent_of_households_without_a_computer_with_broadband_internet",
    "percent_of_population_within_15_min_drive_of_pharmacy",
    "percent_of_population_within_30_min_drive_of_hospital"
  )

dict <- tibble::tibble(cleaned_names, original_names)




usethis::use_data(dict, county_metrics, zcta_metrics, county_percentiles,
                  zcta_percentiles, county_zscores, zcta_zscores,
                  county_ranks, zcta_ranks, internal = TRUE,
                  compress = "xz", overwrite = TRUE)

usethis::use_data(regional_groupings, internal = FALSE,
                  compress = "xz", overwrite = TRUE)
