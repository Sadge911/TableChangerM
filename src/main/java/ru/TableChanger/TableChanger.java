package ru.TableChanger;

//https://o7planning.org/ru/11259/read-write-excel-file-in-java-using-apache-poi#a5136061
//http://java-online.ru/java-excel.xhtml
//https://tproger.ru/translations/how-to-read-write-excel-file-java-poi-example/

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;


public class TableChanger {


    public static void main(String[] args) throws IOException {
    HSSFSheet book;
        int n = 1;

            // Read XSL file
            FileInputStream inputStream = new FileInputStream(new File("F:/1.xls"));

            // Get the workbook instance for XLS file
            HSSFWorkbook workbook = new HSSFWorkbook(inputStream);

            // Get first sheet from the workbook
            HSSFSheet sheet = workbook.getSheetAt(0);

            // Get iterator to all the rows in current sheet
            Iterator<Row> it = sheet.iterator();
            while (it.hasNext()) {
                Row row = it.next();
                Iterator<Cell> cells = row.iterator();
                while (cells.hasNext()) {
                    Cell cell = cells.next();
                    System.out.println("проход " +n);
                    n++;
                }
            }
        System.out.println("");
    }

  //  Sev@$0921232
}