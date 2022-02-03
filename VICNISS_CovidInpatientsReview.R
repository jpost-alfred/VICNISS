#VICNISS COVID Inpatients Review
#Date updated: 2022-01-31

# 0. R Setup --------------------------------------------------------------

# clears anything currently loaded in R
rm(list = ls()); gc()

#Set library and working directory
.libPaths("H:/OBS/Documentation/09. R Library/win-library/4.0")
setwd("I:/infectn/IC/SURVEILLANCE/ALFRED/Outbreaks Significant organisms/COVID-19/COVID VICNISS Reporting")

#Load packages
library(dplyr)
library(openxlsx)


# 1. Read in Data ---------------------------------------------------------
df_PC <- read.xlsx("Extract_PowerChartCovidInpatients.xlsx") %>%
  mutate(MRN = as.character(MRN))
df_VICNISS <- read.xlsx("Extract_VICNISSCovidInpatients.xlsx") %>%
  filter(Discharged == "No")  #Remove discharged patients

#Select patients present in VICNISS but not Powerchart
discharged <- anti_join(df_VICNISS, df_PC, by = c("URNo" = "MRN")) %>%
  select(URNo, Family.Name, First.Name, Gender, Admission.Date, Most.Recent.Location, Most.Recent.Location.Recorded.Date) %>%
  filter(!grepl("Isolation lifted", Most.Recent.Location, ignore.case = TRUE))

#Select patients present in both VICNISS and Powerchart
stayed <-inner_join(df_VICNISS, df_PC, by = c("URNo" = "MRN")) %>%
  select(URNo, Family.Name, First.Name, Gender, Admission.Date, Most.Recent.Location, WARD) %>%
  arrange(WARD) %>%
  rename(PrevLocation = Most.Recent.Location,
         NewLocation = WARD)

#Select patients present in Powerchart but not VICNISS
new <- anti_join(df_PC, df_VICNISS, by = c("MRN" = "URNo")) %>%
  select(MRN, NAME, DOB, AGE, SEX, SITE, WARD, REASON) %>%
  filter(!grepl("HITH", WARD),
         !grepl("Better at Home", WARD, ignore.case = TRUE),
         !grepl("GlenHuntly", WARD, ignore.case = TRUE),
         !grepl("ETQC", NAME, ignore.case = TRUE))


# 2. Export Data ----------------------------------------------------------

sheets <- list(stayed, new, discharged)
wb <- buildWorkbook(sheets, sheetName = c("Stayed", "New", "Discharged"))
saveWorkbook(wb, "VICNISS_DailyList.xlsx", overwrite = TRUE)
