package ru.TableChanger;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.*;
import java.nio.file.FileVisitResult;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.SimpleFileVisitor;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.Iterator;
public class TableChanger {


    public static void main(String[] args){
        //TableOpener.OpenFile("F:\\test\\","1");
        try {
            HSSFWorkbook workbook;
            FileInputStream file = new FileInputStream(new File("F:\\1.xls"));
            // формируем из файла экземпляр HSSFWorkbook

            workbook = new HSSFWorkbook(file);

            // выбираем первый лист для обработки
            // нумерация начинается с 0
            HSSFSheet sheet = workbook.getSheetAt(0);
            Iterator<Row> it = sheet.iterator();
            while (it.hasNext()) {
                Row row = it.next();
                Iterator<Cell> cells = row.iterator();
                while (cells.hasNext()) {
                    Cell cell = cells.next();
                    /*
                    row.getRowNum() == 2 - Вторая строка Отчет от 1
                    row.getCell(1) - 1й Столбец

                     if (cell.getColumnIndex() == 1 && row.getRowNum() == 2) {
                        String test1 = row.getCell(1).getStringCellValue();
                        System.out.println("test111 " + test1);
                    }
                     */
                    String tempCell = cell.getStringCellValue();
                    int intCellColum = cell.getColumnIndex();
                    int intCellRow = cell.getRowIndex();
                    if (tempCell.equals( Resthings.NAMEPFR_etalon)){
                        cell.setCellValue(Resthings.NAMEPFR);
                    }
                }
            }
            //Попали в ячейку С9
//                    if (cell.getColumnIndex() == 2 && row.getRowNum() == 8) {
//                        String test1 = row.getCell(6).getStringCellValue();
//                        System.out.println("test111 " + test1);
//                        cell.setCellValue(Resthings.NAMEPFR);
//
//                    }

            workbook.write(new FileOutputStream("F:\\2.xls"));
            Resthings.done=true;

        } catch (IOException e){
            System.out.println("IO error");
        } catch (Exception e){
            System.out.println(" error");
        }

        // создаем подписи к столбцам (это будет первая строчка в листе Excel файла)
        // записываем созданный в памяти Excel документ в файл
        //        try (FileOutputStream out = new FileOutputStream(new File("F:\\Apache POI Excel File.xls"))) {
        //            workbook.write(out);
        //        } catch (IOException e) {
        //            e.printStackTrace();
        //        }
        if (Resthings.done)
        System.out.println("Excel файл успешно создан!");
    }





}