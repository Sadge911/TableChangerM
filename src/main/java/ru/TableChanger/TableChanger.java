package ru.TableChanger;

//https://o7planning.org/ru/11259/read-write-excel-file-in-java-using-apache-poi#a5136061
//http://java-online.ru/java-excel.xhtml
//https://tproger.ru/translations/how-to-read-write-excel-file-java-poi-example/

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;


public class TableChanger {


    public static void main(String[] args) throws IOException {
    HSSFSheet book;
        int n = 1;
        String test1 = null;
        String test2 = "";
        String test3 = "";

            // Читаем файл
            FileInputStream inputStream = new FileInputStream(new File("F:/1.xls"));

            // Берем воркбук
            HSSFWorkbook workbook = new HSSFWorkbook(inputStream);

            // Первая страница воркбука
            HSSFSheet sheet = workbook.getSheetAt(0);

            // Перебираем строки а в них ячейки.
            Iterator<Row> it = sheet.iterator();
            while (it.hasNext()) {
                Row row = it.next();
                Iterator<Cell> cells = row.iterator();
                while (cells.hasNext()) {
                    Cell cell = cells.next();
                    {
                        if(cell.getCellType() == Cell.CELL_TYPE_STRING)
                        {
                            test1 = cell.getStringCellValue();

                            if (test1 == Resthings.NAMEPFR_etalon){
                                cell.setCellValue(Resthings.NAMEPFR);
                                System.out.println("ЗАмена на ходу №  " +n);
                            }
                        }
                        if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC)
                        {

                        }
                        int cellType = cell.getCellType();

                        //System.out.println("проход " +n);

                        n++;
                    }

                }
            }
        System.out.println(" ");

        //         создаем подписи к столбцам (это будет первая строчка в листе Excel файла)
        //         записываем созданный в памяти Excel документ в файл
                try (FileOutputStream out = new FileOutputStream(new File("F:/2.xls"))) {
                    workbook.write(out);
                    System.out.println("Файл complite");
                } catch (IOException e) {
                    e.printStackTrace();
                }
    }

  //  Sev@$0921232
}