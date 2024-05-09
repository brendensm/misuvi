

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

dict <- data.frame(cleaned_names, original_names)


#'View misuvi data dictionary
#'
#'This package provides a tidy version of the MISUVI data sets. This function returns a data frame of the cleaned variable names and the original variable names.
#'
#'@export dictionary

dictionary <- function(){
  return(dict)
}
