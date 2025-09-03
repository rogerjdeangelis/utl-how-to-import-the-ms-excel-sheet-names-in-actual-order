# utl-how-to-import-the-ms-excel-sheet-names-in-actual-order
How to import the ms excel sheet names in actual order
    %let pgm=utl-how-to-import-the-ms-excel-sheet-names-in-actual-order;

    %stop_submission;

    How to import the ms excel sheet names in actual order

    SOAPBOX ON
    Note: I still use MS excel 2010,
    because 2010 is more compatible with other languages. MS seems to change office
    formats with newer versions?
    You may need to convert your excel 365 to excel 2010?
    SOAPBOX OFF

    INPUT
    =====

    d:/xls/ages.xlsx

    Order of sheets
      age13>age11>age21

    ----------------------
    | A1| fx     |NAME   |
    ---------------------+
    [_] |   A    | B | C |
    ---------------------|
     1  | NAME   |SEX|AGE|
     -- |--------+---+---|
     2  | Alice  | F |13 |
     -- |--------+---+---|
     3  | Barbara| F |13 |
     -- |--------+---+---|
    [AGE13]

    ----------------------
    | A1| fx     |NAME   |
    ---------------------+
    [_] |   A    | B | C |
    ---------------------|
     1  | NAME   |SEX|AGE|
     -- |--------+---+---|
     2  | James  | M |11 |
     -- |--------+---+---|
     3  | James  | M |11 |
     -- |--------+---+---|
    [AGE11]

    -------------------- -
    | A1| fx     |NAME   |
    ---------------------+
    [_] |   A    | B | C |
    ---------------------|
     1  | NAME   |SEX|AGE|
     -- |--------+---+---|
     2  | Carol  | F |21 |
     -- |--------+---+---|
     3  | Henry  | M |21 |
     -- |--------+---+---|
    [AGE21]


    OUTPUT
    ======

    %put &=sheets;

    SHEETS=afe13 afe11 afe21

    github
    https://tinyurl.com/yh9k3s8f
    https://github.com/rogerjdeangelis/utl-how-to-import-the-ms-excel-sheet-names-in-actual-order

    communities.sas
    https://tinyurl.com/5y5nampj
    https://communities.sas.com/t5/SAS-Programming/how-to-display-the-sheet-in-actual-order/m-p/747540#M234636

    /**************************************************************************************************************************/
    /* INPUT                              | PROCESS                                      | OUTPUT                             */
    /* =====                              | =======                                      | ======                             */
    /* d:/xls/ages.xlsx                   | %symdel sheets / nowarn;                     | %put &=sheets;                     */
    /*                                    |                                              |                                    */
    /* Order of sheets                    | %utl_rbeginx;                                | SHEETS=afe13 afe11 afe21           */
    /*   age13>age11>age21                | parmcards4;                                  |                                    */
    /*                                    | library(openxlsx)                            |                                    */
    /* ----------------------             | source("c:/oto/fn_tosas9x.R")                |                                    */
    /* | A1| fx     |NAME   |             | sheet_names <-                               |                                    */
    /* ---------------------+             |   getSheetNames("d:/xls/ages.xlsx")          |                                    */
    /* [_] |   A    | B | C |             | sheets<- paste(sheet_names,collapse=" ")     |                                    */
    /* ---------------------|             | print(sheets)                                |                                    */
    /*  1  | NAME   |SEX|AGE|             | writeClipboard(as.character(sheets))         |                                    */
    /*  -- |--------+---+---|             | ;;;;                                         |                                    */
    /*  2  | Alice  | F |13 |             | %utl_rendx(return=sheets);                   |                                    */
    /*  -- |--------+---+---|             |                                              |                                    */
    /*  3  | Barbara| F |13 |             | %put &=sheets;                               |                                    */
    /*  -- |--------+---+---|             |                                              |                                    */
    /* [AGE13]                            |                                              |                                    */
    /*                                    |                                              |                                    */
    /* ----------------------             |                                              |                                    */
    /* | A1| fx     |NAME   |             |                                              |                                    */
    /* ---------------------+             |                                              |                                    */
    /* [_] |   A    | B | C |             |                                              |                                    */
    /* ---------------------|             |                                              |                                    */
    /*  1  | NAME   |SEX|AGE|             |                                              |                                    */
    /*  -- |--------+---+---|             |                                              |                                    */
    /*  2  | James  | M |11 |             |                                              |                                    */
    /*  -- |--------+---+---|             |                                              |                                    */
    /*  3  | James  | M |11 |             |                                              |                                    */
    /*  -- |--------+---+---|             |                                              |                                    */
    /* [AGE11]                            |                                              |                                    */
    /*                                    |                                              |                                    */
    /* -------------------- -             |                                              |                                    */
    /* | A1| fx     |NAME   |             |                                              |                                    */
    /* ---------------------+             |                                              |                                    */
    /* [_] |   A    | B | C |             |                                              |                                    */
    /* ---------------------|             |                                              |                                    */
    /*  1  | NAME   |SEX|AGE|             |                                              |                                    */
    /*  -- |--------+---+---|             |                                              |                                    */
    /*  2  | Carol  | F |21 |             |                                              |                                    */
    /*  -- |--------+---+---|             |                                              |                                    */
    /*  3  | Henry  | M |21 |             |                                              |                                    */
    /*  -- |--------+---+---|             |                                              |                                    */
    /* [AGE21]                            |                                              |                                    */
    /*                                    |                                              |                                    */
    /* %utlfkil(d:/xls/ages.xlsx) ;       |                                              |                                    */
    /*                                    |                                              |                                    */
    /* ods excel                          |                                              |                                    */
    /*  file="d:/xls/ages.xlsx"           |                                              |                                    */
    /*  options(                          |                                              |                                    */
    /*  sheet_name='afe#byval1');;        |                                              |                                    */
    /* proc print data=have;              |                                              |                                    */
    /* by age notsorted;                  |                                              |                                    */
    /* run;quit;                          |                                              |                                    */
    /* ods excel close;                   |                                              |                                    */
    /**************************************************************************************************************************/


    /*                   _
    (_)_ __  _ __  _   _| |_
    | | `_ \| `_ \| | | | __|
    | | | | | |_) | |_| | |_
    |_|_| |_| .__/ \__,_|\__|
            |_|
    */

    options
     validvarname=upcase;
    data have;
      input
        name$
        sex$ age;
    cards4;
    Alice   F 13
    Barbara F 13
    James   M 11
    James   M 11
    Carol   F 21
    Henry   M 21
    ;;;;
    run;quit;

    %utlfkil(d:/xls/ages.xlsx) ;

    ods excel
     file="d:/xls/ages.xlsx"
     options(
     sheet_name='afe#byval1');;
    proc print data=have;
    by age notsorted;
    run;quit;
    ods excel close;

    /**************************************************************************************************************************/
    /* INPUT                                                                                                                  */
    /* =====                                                                                                                  */
    /* d:/xls/ages.xlsx                                                                                                       */
    /*                                                                                                                        */
    /* Order of sheets                                                                                                        */
    /*   age13>age11>age21                                                                                                    */
    /*                                                                                                                        */
    /* ----------------------                                                                                                 */
    /* | A1| fx     |NAME   |                                                                                                 */
    /* ---------------------+                                                                                                 */
    /* [_] |   A    | B | C |                                                                                                 */
    /* ---------------------|                                                                                                 */
    /*  1  | NAME   |SEX|AGE|                                                                                                 */
    /*  -- |--------+---+---|                                                                                                 */
    /*  2  | Alice  | F |13 |                                                                                                 */
    /*  -- |--------+---+---|                                                                                                 */
    /*  3  | Barbara| F |13 |                                                                                                 */
    /*  -- |--------+---+---|                                                                                                 */
    /* [AGE13]                                                                                                                */
    /*                                                                                                                        */
    /* ----------------------                                                                                                 */
    /* | A1| fx     |NAME   |                                                                                                 */
    /* ---------------------+                                                                                                 */
    /* [_] |   A    | B | C |                                                                                                 */
    /* ---------------------|                                                                                                 */
    /*  1  | NAME   |SEX|AGE|                                                                                                 */
    /*  -- |--------+---+---|                                                                                                 */
    /*  2  | James  | M |11 |                                                                                                 */
    /*  -- |--------+---+---|                                                                                                 */
    /*  3  | James  | M |11 |                                                                                                 */
    /*  -- |--------+---+---|                                                                                                 */
    /* [AGE11]                                                                                                                */
    /*                                                                                                                        */
    /* -------------------- -                                                                                                 */
    /* | A1| fx     |NAME   |                                                                                                 */
    /* ---------------------+                                                                                                 */
    /* [_] |   A    | B | C |                                                                                                 */
    /* ---------------------|                                                                                                 */
    /*  1  | NAME   |SEX|AGE|                                                                                                 */
    /*  -- |--------+---+---|                                                                                                 */
    /*  2  | Carol  | F |21 |                                                                                                 */
    /*  -- |--------+---+---|                                                                                                 */
    /*  3  | Henry  | M |21 |                                                                                                 */
    /*  -- |--------+---+---|                                                                                                 */
    /* [AGE21]                                                                                                                */
    /**************************************************************************************************************************/

    /*
     _ __  _ __ ___   ___ ___  ___ ___
    | `_ \| `__/ _ \ / __/ _ \/ __/ __|
    | |_) | | | (_) | (_|  __/\__ \__ \
    | .__/|_|  \___/ \___\___||___/___/
    |_|
    */

    %symdel sheets / nowarn;

    %utl_rbeginx;
    parmcards4;
    library(openxlsx)
    source("c:/oto/fn_tosas9x.R")
    sheet_names <-
      getSheetNames("d:/xls/ages.xlsx")
    sheets<- paste(sheet_names,collapse=" ")
    print(sheets)
    writeClipboard(as.character(sheets))
    ;;;;
    %utl_rendx(return=sheets);

    %put &=sheets;


    /**************************************************************************************************************************/
    /* %put &=sheets;                                                                                                         */
    /*                                                                                                                        */
    /* SHEETS=afe13 afe11 afe21                                                                                               */
    /**************************************************************************************************************************/

    /*              _
      ___ _ __   __| |
     / _ \ `_ \ / _` |
    |  __/ | | | (_| |
     \___|_| |_|\__,_|

    */
