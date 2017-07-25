package com.xlsxEditor.controller;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

/**
 * Created by crosenlund on 7/15/2017.
 */
public class myWorkbook {

    private XSSFWorkbook workBook;
    
    public myWorkbook() {

        workBook = new XSSFWorkbook();
    }
    
    public myWorkbook(String file_loc) {

        try {
            File myFile = new File(file_loc);
            FileInputStream fis = new FileInputStream(myFile);
            workBook = new XSSFWorkbook(fis);
        } catch(FileNotFoundException e) {
            System.out.println("FileNotFoundException caught");
        } catch (IOException e) {
            System.out.println("IOException caught");
        }
    }

    public XSSFWorkbook getWorkBook() {
        return workBook;
    }
}
