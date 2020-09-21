package ru.TableChanger;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class Resthings {


    /**
     * Наименование органа ПФР
     * */
    public static final String filePath = "F:\\test\\";
    public static final String files = "*.xls";
    public static final String NAMEPFR =
            "Государственное учреждение — Центр по выплате пенсий и обработке информации Пенсионного фонда Российской Федерации в г. Севастополе";
    public static final String NAMEPFR_etalon =
            "Государственное учреждение - Отделение Пенсионного фонда Российской Федерации по г.Севастополю";


    public static int           workbookSheets  = 0;         // Лист нулевой
    boolean                     ex              = false;     // Проверка завершения программы
    public static boolean       done            = false;     //

    HSSFWorkbook                workbook        = null;

}
