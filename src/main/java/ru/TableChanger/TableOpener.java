package ru.TableChanger;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class TableOpener {

    public static void OpenFile(String path, String file){
        int num = 0;
        file = file + ".xls";

        try {
            FileInputStream fileInputStream = new FileInputStream(new File(path+file));

        } catch (IOException e){
            System.out.println("IO не открыли");
        } catch (Exception e){
            System.out.println(" просто експешен");
        }

    }
    public static void SaveFile(){
        try {
            int num = 1;
            FileOutputStream fileOutputStream = new FileOutputStream(new File("F:\\test\\out\\"+num+"\\*.xls"));
            fileOutputStream.close();
            num++;
        }catch (IOException e){
            System.out.println("IO не открыли");
        } catch (Exception e){
            System.out.println(" просто експешен");
        }
    }
}
