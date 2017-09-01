package com.xlsxEditor.controller;

import java.io.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class xlsxEditor implements Runnable {

    private String path;
    private String output_path = "";
    private String sheetName;
    private String column;
    private xlsxEditorController xdc;

    public xlsxEditor(String path, String sheetName, String column, xlsxEditorController xdc) {
        this.path = path;
        this.sheetName = sheetName;
        this.column = column;
        this.xdc = xdc;
//        this.output_path = path;
        String[] path_split = path.split("\\.");
        for (int s = 0; s < path_split.length - 1; s++) {
            output_path += path_split[s];
            if (s < path_split.length - 2) {
                output_path += ".";
            }
        }
        output_path += "_xpaths_added." + path_split[path_split.length - 1];
    }
    
    public void run() {
        int exit_value = 0;
        int xpathColumn = 0;
        // convert the column input by the user to a number if a letter was entered
        if (column.matches("\\d+")) {
            xpathColumn = Integer.parseInt(column) - 1;
        } else {
            xpathColumn = CellReference.convertColStringToIndex(column);
        }
        
        XSSFWorkbook workbook = null;
        try {

            File myFile = new File(path);
            FileInputStream fis = new FileInputStream(myFile);
            workbook = new XSSFWorkbook(fis);
            fis.close();
        } catch (FileNotFoundException e) {
            exit_value = 1;
        } catch (IOException e) {
            exit_value = 2;
        }
        
        if (workbook != null && workbook.getSheetIndex(sheetName) != -1) {
            XSSFSheet mySheet = workbook.getSheet(sheetName);

            String groupXpath = "";
            for (int rowNumber = 0; rowNumber <= mySheet.getLastRowNum(); rowNumber++) {
                XSSFRow row = mySheet.getRow((rowNumber));
                if (row != null && row.getLastCellNum() > 0) {

                    XSSFCell cell0 = row.getCell(0);

                    // parse xpath from the generate group name (7.6 and 7.7)
                    if (cell0.toString().matches("<.*-.*?/(.+?)>.*")) {
                        groupXpath = cell0.toString().replaceAll("<.*-.*?/(.+?)>.*", "$1");
                    }
                    // RSX below 7.6 has a different generated xpath to parse from (7.2-7.5)
                    else if (cell0.toString().matches(".*-\\s(.+?)")) {
                        groupXpath = cell0.toString().replaceAll(".*-\\s(.+?)", "$1");
                        String[] xpathArray = groupXpath.split("\\.");
                        String new_groupXpath = "";
                        for (String xpath : xpathArray) {
                            if (!xpath.contains("Rep")) {
                                new_groupXpath = new_groupXpath + xpath + "/";
                            }
                        }
                        groupXpath = new_groupXpath.replaceAll("/$", "");
                    }
                    // check that we have found the xpath from the group and we only want to change the fields
                    else if (!groupXpath.equals("") && cell0.toString().matches("\\d+") && !row.getCell(1).toString().contains(" ")) {
                        XSSFCell xpathCell = row.getCell(xpathColumn);
                        xpathCell.setCellValue(groupXpath + "/" + row.getCell(1).toString());
                    }
                }
            }

            if (exit_value == 0) {
                try {
                    FileOutputStream fos = new FileOutputStream(new File(output_path));
                    workbook.write(fos);
                    fos.close();
                } catch (FileNotFoundException e) {
                    exit_value = 4;
                } catch (IOException e) {
                    exit_value = 5;
                }
            }
        } else {
            if (exit_value == 0) {
                // workSheet/tab not found in workbook
                exit_value = 3;
            }
        }
        
        //return any errors after xpaths are (attempted) added
        this.xdc.updateFroRun(exit_value);
    }
}