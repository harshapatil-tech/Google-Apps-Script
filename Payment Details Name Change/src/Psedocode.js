/**
 *  Payment Details for SMEs for the last month when this sheet is getting created.
 * 
 *  Script runs 5th of every month
 *    - It checks whether the month is May.
 *    - If month is May ->
 *          - create a folder for this financial year
 *          - copy previous year's spreadsheets from previous fin-year folder to this fin-year folder
 *              (  folder nomenclature is 2024-25, 2025-26 etc.  )
 *          - every spreadsheet will contain 12 sheets -> each sheet for each month (month_yy)
 *    
 *    - Get employee data and the organizational structure from "Employee Details" spreadsheet.
 *      and then verify with the previous month Timetable Tutor_Names for that department.
 *          - For each sheet (every department)
 *                - Senior SMEs for each department are listed at the top of every month
 *                - Each SME below the Senior SMEs will be ordered alphabetically and will have the junior SMEs working under them come right below the SMEs
 *                    (format 
 *                        SME 1
 *                          > Junior SME 1
 *                          > Junior SME 2
 *                        SME 2
 *                          > Junior SME 1
 *                          > Junior SME 2
 *                    )
 */