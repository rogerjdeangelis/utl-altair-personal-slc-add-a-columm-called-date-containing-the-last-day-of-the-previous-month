%let pgm=utl-altair-personal-slc-add-a-columm-called-date-containing-the-last-day-of-the-previous-month;

%stop_submission;

RE: Altair personal slc add a columm called date containing the last day of the previous month

Too long to post in listserv, see github

github
https://github.com/rogerjdeangelis/utl-altair-personal-slc-add-a-columm-called-date-containing-the-last-day-of-the-previous-month

I received an excel workbook today, 2025-11-01, and I need to add a column called last_day_prev_month that contains the last
day of the previous month, 2025-10-31.

Although you can update an excel sheet inplace using SAS, Python or R, I suggest you add a new sheet with the last_day_prev_month.added.
R and Python also have the ability to preserve styles, even read formats and excel color coding.

https://community.altair.com/discussion/24183

/*               _     _
 _ __  _ __ ___ | |__ | | ___ _ __ ___
| `_ \| `__/ _ \| `_ \| |/ _ \ `_ ` _ \
| |_) | | | (_) | |_) | |  __/ | | | | |
| .__/|_|  \___/|_.__/|_|\___|_| |_| |_|
|_|
*/

/********************************************************************************************************************/
/*                    INPUT                       |       OUTPUT(aaded columns last_day_prev_month & current_date   */
/*                    =====                       |       =======================================================   */
/*                                                |                                                                 */
/* d:/xls/students.xlsx sheet=STUDENTS            | d:/xls/students.xlsx sheet=STUDENTS_UPDATE                      */
/*                                                |                                                                 */
/* ----------------------+                        | ----------------------------------+                             */
/* | A1| fx     |STUDENT |                        | | A1| fx | LAST_DAY_PREV_MONTH    |                             */
/* ---------------------------------------------+ | --------------------------------------------------------------+ */
/* [_] |    A   |   B    |   c  |   D   |    E  | | [_] |    A     |    B     |    C  |  D  |   E |    F  |    G  | */
/* ---------------------------------------------| | --------------------------------------------------------------| */
/*  1  |STUDENT | YEAR   | STATE| GRADE1| GRADE2| |  1  |LAST_DAY  |CURRENT   |       |     |     |       |       | */
/*  -- |--------+--------+------+-------+-------| |  1  |PREV_MONTH|DATE      |STUDENT|YEAR |STATE| GRADE1| GRADE2| */
/*  2  |  JACK  | 2020   |  NC  |  85   |  87   | |  -- |----------+----------+-------+-----+-----+-------+-------| */
/*  -- |--------+--------+------+-------+-------| |  2  |2025-10-31|2025-11-01|  JACK |2020 | NC  |  85   |  87   | */
/*  3  |  ALEX  | 2025   |  MS  |  91   |  92   | |  -- |----------+--------- +-------+-----+-----+-------+-------| */
/*  -- |--------+--------+------+-------+-------| |  3  |2025-10-31|2025-11-01|  ALEX |2025 | MS  |  91   |  92   | */
/*  4  |  BARB  | 2018   |  TN  |  78   |  92   | |  -- |----------+----------+-------+-----+-----+-------+-------| */
/*  -- |--------+--------+------+-------+-------| |  4  |2025-10-31|2025-11-01|  BARB |2018 | TN  |  78   |  92   | */
/*  5  |  MARY  | 2020   |  NY  |  87   |  95   | |  -- |----------+----------+-------+-----+-----+-------+-------| */
/*  -- |--------+--------+------+-------+-------| |  5  |2025-10-31|2025-11-01|  MARY |2020 | NY  |  87   |  95   | */
/*  6  |  JEFF  | 2025   |  NC  |  96   |  98   | |  -- |----------+----------+-------+-----+-----+-------+-------| */
/*  -- ---------+--------+------+-------+-------+ |  6  |2025-10-31|2025-11-01|  JEFF |2025 | NC  |  96   |  98   | */
/*                                                |  -- |----------+----------+-------+-----+-----+-------+-------+ */
/*  [STUDENTS]                                    |                                                                 */
/*                                                |  [STUDENTS_UPDATE]                                              */
/********************************************************************************************************************/

/*   _                   _                       _
/ | (_)_ __  _ __  _   _| |_    _____  _____ ___| |
| | | | `_ \| `_ \| | | | __|  / _ \ \/ / __/ _ \ |
| | | | | | | |_) | |_| | |_  |  __/>  < (_|  __/ |
|_| |_|_| |_| .__/ \__,_|\__|  \___/_/\_\___\___|_|
            |_|
*/

%utlfkil(d:/xls/student.xlsx);

libname xls excel "d:/xls/student.xlsx";

data xls.students;
  input student$ year $ state $ grade1 grade2;
  label year = "Year of Birth";
cards4;
JACK 2020 NC 85 87
ALEX 2025 MS 91 92
BARB 2018 TN 78 92
MARY 2020 NY 87 95
JEFF 2025 NC 96 98
;;;;
run;quit;

libname xls clear;

/*
 _ __  _ __ ___   ___ ___  ___ ___
| `_ \| `__/ _ \ / __/ _ \/ __/ __|
| |_) | | | (_) | (_|  __/\__ \__ \
| .__/|_|  \___/ \___\___||___/___/
|_|
*/

/*---- for testing and in case of rerun      ----*/
/*---- sas and the slc cannot delete a sheet ----*/
/*---- delete sheet students_update          ----*/

options set=RHOME "D:\d451";
proc r;
submit;
library(openxlsx)
path<-"d:/xls/student.xlsx"
wb <- loadWorkbook(file = path )
removeWorksheet(wb,"students_update")
saveWorkbook(wb, path, overwrite = TRUE)
endsubmit;
run;quit;

/*---- add dates current and previous        ----*/

libname xls excel "d:/xls/student.xlsx";

data xls.students_update;
   retain last_day_prev_month current_date;
   set xls.'students$'n;
   last_day_prev_month = put(intnx('month', today(), -1, 'E'),e8601da.);
   current_date = put( today(),e8601da.);
run;quit;

libname xls clear;

/*           _               _
  ___  _   _| |_ _ __  _   _| |_
 / _ \| | | | __| `_ \| | | | __|
| (_) | |_| | |_| |_) | |_| | |_
 \___/ \__,_|\__| .__/ \__,_|\__|
                |_|
*/

d:/xls/students.xlsx sheet=STUDENTS_UPDATE

----------------------------------+
| A1| fx | LAST_DAY_PREV_MONTH    |
--------------------------------------------------------------+
[_] |    A     |    B     |    C  |  D  |   E |    F  |    G  |
--------------------------------------------------------------|
 1  |LAST_DAY  |CURRENT   |       |     |     |       |       |
 1  |PREV_MONTH|DATE      |STUDENT|YEAR |STATE| GRADE1| GRADE2|
 -- |----------+----------+-------+-----+-----+-------+-------|
 2  |2025-10-31|2025-11-01|  JACK |2020 | NC  |  85   |  87   |
 -- |----------+--------- +-------+-----+-----+-------+-------|
 3  |2025-10-31|2025-11-01|  ALEX |2025 | MS  |  91   |  92   |
 -- |----------+----------+-------+-----+-----+-------+-------|
 4  |2025-10-31|2025-11-01|  BARB |2018 | TN  |  78   |  92   |
 -- |----------+----------+-------+-----+-----+-------+-------|
 5  |2025-10-31|2025-11-01|  MARY |2020 | NY  |  87   |  95   |
 -- |----------+----------+-------+-----+-----+-------+-------|
 6  |2025-10-31|2025-11-01|  JEFF |2025 | NC  |  96   |  98   |
 -- |----------+----------+-------+-----+-----+-------+-------+
                                                               *
 [STUDENTS_UPDATE]

/*              _
  ___ _ __   __| |
 / _ \ `_ \ / _` |
|  __/ | | | (_| |
 \___|_| |_|\__,_|

*/
