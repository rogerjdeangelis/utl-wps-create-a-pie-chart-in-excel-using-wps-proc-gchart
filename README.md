# utl-wps-create-a-pie-chart-in-excel-using-wps-proc-gchart
WPS create a pie chart in excel using wps proc gchart
    %let pgm=utl-wps-create-a-pie-chart-in-excel-using-wps-proc-gchart;

    WPS create a pie chart in excel using wps proc gchart

    image
    http://tinyurl.com/yxme2hhz
    https://github.com/rogerjdeangelis/utl-wps-create-a-pie-chart-in-excel-using-wps-proc-gchart/blob/main/gchart.png

    excel workbook
    http://tinyurl.com/2dfpzx4y
    https://github.com/rogerjdeangelis/utl-wps-create-a-pie-chart-in-excel-using-wps-proc-gchart/blob/main/gchart.xlsx

    github
    http://tinyurl.com/mtc2xaw9
    https://github.com/rogerjdeangelis/utl-wps-create-a-pie-chart-in-excel-using-wps-proc-gchart


    /*               _     _
     _ __  _ __ ___ | |__ | | ___ _ __ ___
    | `_ \| `__/ _ \| `_ \| |/ _ \ `_ ` _ \
    | |_) | | | (_) | |_) | |  __/ | | | | |
    | .__/|_|  \___/|_.__/|_|\___|_| |_| |_|
    |_|
    */

     /**************************************************************************************************************************/
     /*                                                                                                                        */
     /*            INPUT                                             PROCESS                                                   */
     /*            =====                                             ======                                                    */
     /*                                                                                                                        */
     /*  SD1.HAVE total obs=163                                       WPS                                                      */
     /*                                                                                                                        */
     /*   MODEL                               TYPE              proc gchart data=sd1.have;                                     */
     /*                                                            pie type ;                                                  */
     /*   MDX                                 SUV               run;quit;                                                      */
     /*   NSX coupe 2dr manual S              Sports            proc r;                                                        */
     /*   RS 6 4dr                            Sports            library(openxlsx);                                             */
     /*   TT 1.8 convertible 2dr (coupe)      Sports            wb <- createWorkbook("d:/xls/gchart.xlsx");                    */
     /*   TT 1.8 Quattro 2dr (convertible)    Sports            addWorksheet(wb, "pie");                                       */
     /*   TT 3.2 coupe 2dr (convertible)      Sports            insertImage(wb, "pie", "d:/png/gchart.png",                    */
     /*   A6 3.0 Avant Quattro                Wagon               width=10, height=8, units="in"  );                           */
     /*   S4 Avant Quattro                    Wagon             saveWorkbook(wb,"d:/xls/gchart.xlsx",                          */
     /*   X3 3.0i                             SUV                overwrite = TRUE);                                            */
     /*   ....                                                                                                                 */
     /*                         OUTPUT                                                                                         */
     /*                         =====                                                                                          */
     /*                                                                                                                        */
     /*                        ***********       SUV                                                                           */
     /*                    ****           ****    60                                                                           */
     /*                  **                   **                                                                               */
     /*                **                       **                                                                             */
     /*               *  ..                       *                                                                            */
     /*             **     .                       **                                                                          */
     /*            **       .                       **                                                                         */
     /*            *         ..                      *                                                                         */
     /*           *            .                      *                                                                        */
     /*           *              .                    *                                                                        */
     /*           *               .                   *                                                                        */
     /*           *                 +  . . .. . .. . .*                                                                        */
     /*   Sports  *                                   *                                                                        */
     /*    49     *               .  .                *                                                                        */
     /*           *              ..   .               *                                                                        */
     /*            *            .      .             *                                                                         */
     /*            **          .       ..           **                                                                         */
     /*              *        .          .         *                                                                           */
     /*               *      .           .        *   Wagon                                                                    */
     /*                **    .            .     **      30                                                                     */
     /*                  ** .              .  **                                                                               */
     /*                    ****           ****                                                                                 */
     /*                        ***********                                                                                     */
     /*                          Truck                                                                                         */
     /*                            24                                                                                          */
     /*                                                                                                                        */
     /**************************************************************************************************************************/

    /*                   _
    (_)_ __  _ __  _   _| |_
    | | `_ \| `_ \| | | | __|
    | | | | | |_) | |_| | |_
    |_|_| |_| .__/ \__,_|\__|
            |_|
    */

    options validvarname=upcase;
    libname sd1 "d:/sd1";
    data sd1.have;
      set sashelp.cars(keep=model type
           where=(type not in ('Hybrid','Sedan')));
    run;quit;

    /**************************************************************************************************************************/
    /*                                                                                                                        */
    /* SD1.HAVE total obs=163                                                                                                 */
    /*                                                                                                                        */
    /*  Obs    MODEL                               TYPE                                                                       */
    /*                                                                                                                        */
    /*    1    MDX                                 SUV                                                                        */
    /*    2    NSX coupe 2dr manual S              Sports                                                                     */
    /*    3    RS 6 4dr                            Sports                                                                     */
    /*    4    TT 1.8 convertible 2dr (coupe)      Sports                                                                     */
    /*    5    TT 1.8 Quattro 2dr (convertible)    Sports                                                                     */
    /*    6    TT 3.2 coupe 2dr (convertible)      Sports                                                                     */
    /*    7    A6 3.0 Avant Quattro                Wagon                                                                      */
    /*    8    S4 Avant Quattro                    Wagon                                                                      */
    /*    9    X3 3.0i                             SUV                                                                        */
    /*   10    X5 4.4i                             SUV                                                                        */
    /*   11    M3 coupe 2dr                        Sports                                                                     */
    /*   12    M3 convertible 2dr                  Sports                                                                     */
    /*   13    Z4 convertible 2.5i 2dr             Sports                                                                     */
    /*                                                                                                                        */
    /**************************************************************************************************************************/

    /*
     _ __  _ __ ___   ___ ___  ___ ___
    | `_ \| `__/ _ \ / __/ _ \/ __/ __|
    | |_) | | | (_) | (_|  __/\__ \__ \
    | .__/|_|  \___/ \___\___||___/___/
    |_|
    */

    %utlfkil(d:/png/gchart.png);
    %utlfkil(d:/xls/gchart.xlsx);

    %utl_submit_wps64x('
    libname sd1 "d:/sd1";
    filename grfout "d:/png/gchart.png";
    goptions reset=all device=png gsfname=grfout display htext=1.5;
    proc gchart data=sd1.have;
       pie type ;
    run;quit;
    filename grfout clear;

    proc r;
    submit;
    library(openxlsx);
    wb <- createWorkbook("d:/xls/gchart.xlsx");
    addWorksheet(wb, "pie");
    insertImage(wb, "pie", "d:/png/gchart.png", width=10, height=8, units="in"  );
    saveWorkbook(wb,"d:/xls/gchart.xlsx",overwrite = TRUE);
    endsubmit;
    ');

    /*                 _       _               _           _            _                _
      _____  _____ ___| |  ___| |__   ___  ___| |_   _ __ (_) ___   ___| |__   __ _ _ __| |_
     / _ \ \/ / __/ _ \ | / __| `_ \ / _ \/ _ \ __| | `_ \| |/ _ \ / __| `_ \ / _` | `__| __|
    |  __/>  < (_|  __/ | \__ \ | | |  __/  __/ |_  | |_) | |  __/| (__| | | | (_| | |  | |_
     \___/_/\_\___\___|_| |___/_| |_|\___|\___|\__| | .__/|_|\___| \___|_| |_|\__,_|_|   \__|
                                                    |_|
    */

    /**************************************************************************************************************************/
    /*                                                                                                                        */
    /*  d:/xls/gchart.xlsx[pie]                                                                                               */
    /*                                                                                                                        */
    /*                         OUTPUT                                                                                         */
    /*                         =====                                                                                          */
    /*                                                                                                                        */
    /*                        ***********       SUV                                                                           */
    /*                    ****           ****    60                                                                           */
    /*                  **                   **                                                                               */
    /*                **                       **                                                                             */
    /*               *  ..                       *                                                                            */
    /*             **     .                       **                                                                          */
    /*            **       .                       **                                                                         */
    /*            *         ..                      *                                                                         */
    /*           *            .                      *                                                                        */
    /*           *              .                    *                                                                        */
    /*           *               .                   *                                                                        */
    /*           *                 +  . . .. . .. . .*                                                                        */
    /*   Sports  *                                   *                                                                        */
    /*    49     *               .  .                *                                                                        */
    /*           *              ..   .               *                                                                        */
    /*            *            .      .             *                                                                         */
    /*            **          .       ..           **                                                                         */
    /*              *        .          .         *                                                                           */
    /*               *      .           .        *   Wagon                                                                    */
    /*                **    .            .     **      30                                                                     */
    /*                  ** .              .  **                                                                               */
    /*                    ****           ****                                                                                 */
    /*                        ***********                                                                                     */
    /*                          Truck                                                                                         */
    /*                            24                                                                                          */
    /*                                                                                                                        */
    /**************************************************************************************************************************/

     /*              _
      ___ _ __   __| |
     / _ \ `_ \ / _` |
    |  __/ | | | (_| |
     \___|_| |_|\__,_|

    */
