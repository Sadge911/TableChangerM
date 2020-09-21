package ru.TableChanger.api;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class openBook {
    public XSSFWorkbook book;
    public void openBookDirectly(final String path){
        File file = new File(path);
        try {
            OPCPackage pkg = OPCPackage.open(file);
            book = new XSSFWorkbook(pkg);
            pkg.close();
        } catch (FileNotFoundException e1) {
            e1.printStackTrace();
        } catch (EncryptedDocumentException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void openBook(final String path)
    {
        try {
            File file = new File(path);
            book = (XSSFWorkbook) WorkbookFactory.create(file);

        InputStream is = new FileInputStream(file);
        book = (XSSFWorkbook) WorkbookFactory.create(is);
         is.close();
        } catch (FileNotFoundException e1) {
            e1.printStackTrace();
        } catch (EncryptedDocumentException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
