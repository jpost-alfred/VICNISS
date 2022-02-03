# VICNISS
This repository contains code for VICNISS COVID Inpatient Reporting. It is run daily to prepare the New, Stayed, and Discharged patient lists.

Project scripts:
1. VICNISS_CovidInpatientsReview.R

The code is located here: __I:\infectn\IC\SURVEILLANCE\ALFRED\Outbreaks Significant organisms\COVID-19\COVID VICNISS Reporting__ 

Process:
1.	Download the current patient list from VICNISS, delete the first two rows, and then save as ‘Extract_VICNISSCovidInpatients.xlsx’ in the folder above (overwrite the current file)
2.	Copy the patient list from PowerChart, paste into an excel and delete any blank columns, and then save as ‘Extract_PowerChartCovidInpatients.xlsx’ in the folder above (overwrite the current file)
3.	Run the script: **VICNISS_CovidInpatientsReview.R**
4.	Lists are exported to the file **VICNISS_DailyList.xlsx** in the same folder. 


Data source(s): Powerchart (current COVID inpatients), VICNISS inpatients report

Last Updated: 2022-02-03