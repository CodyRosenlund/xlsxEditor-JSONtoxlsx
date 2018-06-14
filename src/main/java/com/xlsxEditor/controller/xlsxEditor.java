package com.xlsxEditor.controller;

import java.io.*;
import java.util.HashMap;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Cody Rosenlund, SPS Commerce
 * June 2018
 * <p>
 * xlsxEditor.java
 * <p>
 * Does the brunt of the work. Parses through the xlsx file and creates the retailers
 */
public class xlsxEditor implements Runnable {

    private String path;
    private String output_path = "";
    private String rsxVersion;
    private String docType;
    private xlsxEditorController xdc;
    private HashMap<String, Retailer> retailersHash;

    public xlsxEditor(String path, String rsxVersion, String docType, xlsxEditorController xdc) {
        this.path = path;
        this.rsxVersion = rsxVersion;
        this.docType = docType;
        this.retailersHash = new HashMap<>();
        this.xdc = xdc;
        String[] path_split = path.split("\\.");
        for (int s = 0; s < path_split.length - 2; s++) {
            output_path += path_split[s];
        }
    }

    public void run() {
        int exit_value = 0;

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

        if (workbook != null && workbook.getNumberOfSheets() > 0) {
            XSSFSheet mySheet = workbook.getSheetAt(0);

            /*
                Column 0 is where we will get the hierarchy of the file
                Column 1 is where we will get the field xpaths
                Column 2 starting here create a retailer for every 3 columns starting at index 2
                    - The first column of this subset will determine Mandatory/Optional/Conditional (for field and group)
                    - The second column of this subset will contain qualifiers
                    - The third column of this subset will contain other notes specific to the retailer
             */

            HashMap<Integer, String> hierarchyHash = new HashMap<>();
            for (int rowNumber = 0; rowNumber <= mySheet.getLastRowNum(); rowNumber++) {
                XSSFRow row = mySheet.getRow(rowNumber);

                // hierarchy of the groups for each retailer
                if (row != null && row.getLastCellNum() > 0) {
                    XSSFCell xpathCell = row.getCell(0);
                    if (xpathCell != null && xpathCell.toString() != null && xpathCell.toString().matches("^\\*.*")) {
                        String xpathCellString = xpathCell.toString();

                        /*
                            The thought here for how to turn the hierarchy from * to an xpath:
                            We know the root because the user selected it in the GUI
                            Get the group (starts with *'s) from the xlsx cell
                            Split the * from the group name
                            Count the * and use that as the key for inserting into a hash map
                                - subsequent groups will over write values in the hash as we traverse the hierarchy
                            Then for each *, grab that many values from the hash and concat to gether with '/'
                         */
                        hierarchyHash.put(xpathCellString.trim().replaceAll("^(\\**).*", "$1").length(), xpathCellString.trim().replaceAll("^\\**(.*)", "$1"));

                        String groupXpath = docType;
                        for (int i = 1; i <= xpathCellString.trim().replaceAll("^(\\**).*", "$1").length(); i++) {
                            groupXpath = groupXpath + "/" + hierarchyHash.get(i);
                        }

                        //start at column 2, every 3 columns is a new retailer if present
                        for (int r = 2; r <= row.getLastCellNum(); r = r + 3) {
                            if (mySheet.getRow(0).getLastCellNum() >= r) {
                                if (mySheet.getRow(0).getCell(r) == null || mySheet.getRow(1).getCell(r) == null) {
                                    exit_value = 6;
                                } else {
                                    String retailerHubID = mySheet.getRow(0).getCell(r).toString();
                                    String retailerName = mySheet.getRow(1).getCell(r).toString();

                                    String groupInfo = "";

                                    XSSFCell retailerCell = row.getCell(r);
                                    if (retailerCell != null) {
                                        groupInfo = retailerCell.toString();
                                    }

                                    //if matrix line does not contain any info for this field for this retailer, don't worry about it
                                    if (groupInfo != null && !groupInfo.equals("")) {
                                        //get or create the retailer, add the generated group info
                                        Retailer retailer = retailersHash.containsKey(retailerName) ? retailersHash.get(retailerName) : new Retailer(retailerName, retailerHubID);
                                        retailersHash.put(retailerName, retailer);
                                        retailer.addGroup(groupXpath, groupInfo);
                                    }
                                }
                            }
                        }
                    }
                }

                //grab xpaths and info per retailer
                if (row != null && row.getLastCellNum() > 1) {
                    XSSFCell xpathCell = row.getCell(1);
                    if (xpathCell != null) {
                        String xpathCellString = xpathCell.toString();
                        if (xpathCellString.contains("/") && !xpathCellString.equals("#N/A")) {
                            //start at column 2, every 3 columns is a new retailer if present
                            for (int r = 2; r <= row.getLastCellNum(); r = r + 3) {
                                if (mySheet.getRow(1).getLastCellNum() >= r) {
                                    // if we have blank columns
                                    if (mySheet.getRow(0).getCell(r) == null || mySheet.getRow(1).getCell(r) == null) {
                                        exit_value = 6;
                                    } else {
                                        String retailerHubID = mySheet.getRow(0).getCell(r).toString();
                                        String retailerName = mySheet.getRow(1).getCell(r).toString();

                                        String retailerMandatory = "";
                                        String retailerQuals = "";
                                        String retailerInfo = "";

                                        XSSFCell retailerCell = row.getCell(r);
                                        if (retailerCell != null) {
                                            retailerMandatory = retailerCell.toString();
                                        }
                                        XSSFCell retailerCell2 = row.getCell(r + 1);
                                        if (retailerCell2 != null) {
                                            retailerQuals = retailerCell2.toString();
                                        }
                                        XSSFCell retailerCell3 = row.getCell(r + 2);
                                        if (retailerCell3 != null) {
                                            retailerInfo = retailerCell3.toString();
                                        }

                                        //if matrix line does not contain any info for this field for this retailer, don't worry about it
                                        if ((retailerMandatory != null && !retailerMandatory.equals("")) || (retailerQuals != null && !retailerQuals.equals("")) || (retailerInfo != null && !retailerInfo.equals(""))) {
                                            //get or create the retailer, add the gathered field info
                                            Retailer retailer = retailersHash.containsKey(retailerName) ? retailersHash.get(retailerName) : new Retailer(retailerName, retailerHubID);
                                            retailersHash.put(retailerName, retailer);
                                            retailer.addField(xpathCellString, retailerMandatory, retailerQuals, retailerInfo);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }


            //pass the gathered retailer info to modify a template and create the json files
            if (!retailersHash.isEmpty()) {
                exit_value = new JSONBuilder().createFiles(rsxVersion, docType, retailersHash, output_path);
            } else {
                //Unable to gather retailer information
                exit_value = 4;
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