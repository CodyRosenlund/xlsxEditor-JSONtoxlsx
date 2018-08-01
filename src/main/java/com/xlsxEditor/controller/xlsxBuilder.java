package com.xlsxEditor.controller;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

class xlsxBuilder {

    int buildxlsx(LinkedList<Node> linkedNodes) {

        //get tradingPartners from header TradingPartnerId notes
        ArrayList<String> tradingPartners = getTradingPartners(linkedNodes);
        HashMap<String, Integer> partnerColumns = new HashMap<>();
        final int PARTNER_STARTING_COLUMN = 10;
        final int PARTNER_STARTING_COLUMN_OFFSET = 3;
        final int PARTNER_QUALIFIERS_COLUMN_OFFSET = 1;
        final int PARTNER_NOTES_COLUMN_OFFSET = 2;

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Format");

        addHowToTab(workbook);

        //cell fonts
        XSSFFont bold = workbook.createFont();
        bold.setFontName(XSSFFont.DEFAULT_FONT_NAME);
        bold.setBold(true);

        XSSFFont bold_16pt = workbook.createFont();
        bold_16pt.setFontName(XSSFFont.DEFAULT_FONT_NAME);
        bold_16pt.setBold(true);
        bold_16pt.setFontHeightInPoints((short) 16);

        //cell styles
        CellStyle wordWrapStyle = workbook.createCellStyle();
        wordWrapStyle.setWrapText(true);

        CellStyle alignCenterStyle = workbook.createCellStyle();
        alignCenterStyle.setAlignment(HorizontalAlignment.CENTER);

        CellStyle alignCenterBoldStyle = workbook.createCellStyle();
        alignCenterBoldStyle.setAlignment(HorizontalAlignment.CENTER);
        alignCenterBoldStyle.setFont(bold);

        CellStyle alignCenterBold_leftBorderStyle = workbook.createCellStyle();
        alignCenterBold_leftBorderStyle.setAlignment(HorizontalAlignment.CENTER);
        alignCenterBold_leftBorderStyle.setBorderLeft(BorderStyle.THIN);
        alignCenterBold_leftBorderStyle.setFont(bold);

        CellStyle alignCenterBold16ptLightGreenStyle = workbook.createCellStyle();
        alignCenterBold16ptLightGreenStyle.setAlignment(HorizontalAlignment.CENTER);
        alignCenterBold16ptLightGreenStyle.setFont(bold_16pt);
        alignCenterBold16ptLightGreenStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        alignCenterBold16ptLightGreenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        CellStyle alignCenterBold_with_boarders_Style = workbook.createCellStyle();
        alignCenterBold_with_boarders_Style.setFont(bold);
        alignCenterBold_with_boarders_Style.setAlignment(HorizontalAlignment.CENTER);
        alignCenterBold_with_boarders_Style.setBorderBottom(BorderStyle.THIN);
        alignCenterBold_with_boarders_Style.setBorderTop(BorderStyle.THIN);
        alignCenterBold_with_boarders_Style.setBorderLeft(BorderStyle.THIN);
        alignCenterBold_with_boarders_Style.setBorderRight(BorderStyle.THIN);

        CellStyle leftBoarderStyle = workbook.createCellStyle();
        leftBoarderStyle.setBorderLeft(BorderStyle.THIN);

        CellStyle boldStyle = workbook.createCellStyle();
        boldStyle.setFont(bold);

        CellStyle boldGrey25Style = workbook.createCellStyle();
        boldGrey25Style.setFont(bold);
        boldGrey25Style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        boldGrey25Style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        CellStyle grey25InfillStyle = workbook.createCellStyle();
        grey25InfillStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        grey25InfillStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        CellStyle wordWrapGrey25InfillStyle = workbook.createCellStyle();
        wordWrapGrey25InfillStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        wordWrapGrey25InfillStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        wordWrapGrey25InfillStyle.setWrapText(true);

        /*
         * create the first 7 lines as follows:
         * line 1: col A "Consolidated Matrix" starting at col J and every third col after that insert the partners
         * line 2: col A-I "Consolidation"
         * line 3: col A "Partners" col B list the partners, for every partner starting col J "M/O/C", "Qualifiers", "Design Notes"
         * line 4: col A "Version" col B version number
         * line 5: empty for now
         * line 6: empty for now
         * line 7: col A "Xpath", B "Field", C "Definition", D "Data Type", E "min", F "max", G "Usage", H "Qualifiers", J "Design Notes"
         */

        // Start with columns A - I, we will loop through for the partner columns
        Row row0 = sheet.createRow(0);
        row0.createCell(0).setCellValue("Consolidated Matrix");
        row0.getCell(0).setCellStyle(alignCenterBoldStyle);

        Row row1 = sheet.createRow(1);
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 9));
        row1.createCell(0).setCellValue("SPS Retailer Consolidation");
        row1.getCell(0).setCellStyle(alignCenterBold16ptLightGreenStyle);

        Row row2 = sheet.createRow(2);
        row2.createCell(0).setCellValue("Partners");
        row2.getCell(0).setCellStyle(boldStyle);
        row2.createCell(1).setCellValue(String.join(",", tradingPartners));
        row2.getCell(1).setCellStyle(boldStyle);
        row2.createCell(2).setCellValue(String.join(",", tradingPartners));
        row2.getCell(2).setCellStyle(boldStyle);

        Row row3 = sheet.createRow(3);
        row3.createCell(0).setCellValue("Version");
        row3.getCell(0).setCellStyle(boldStyle);
        row3.createCell(1).setCellValue("7.7");
        row3.getCell(1).setCellStyle(boldStyle);
        row3.createCell(2).setCellValue("7.7");
        row3.getCell(2).setCellStyle(boldStyle);

        //empty rows not used right now but need for styling
        Row row4 = sheet.createRow(4);

        Row row5 = sheet.createRow(5);

        Row row6 = sheet.createRow(6);
        //Styles for entire row
        row6.createCell(0).setCellValue("Xpath");
        row6.createCell(1).setCellValue("Field");
        row6.createCell(2).setCellValue("Field");
        row6.createCell(3).setCellValue("Definition");
        row6.createCell(4).setCellValue("Data Type");
        row6.createCell(5).setCellValue("min");
        row6.createCell(6).setCellValue("max");
        row6.createCell(7).setCellValue("Usage");
        row6.createCell(8).setCellValue("Qualifiers");
        row6.createCell(9).setCellValue("Design Notes");

        for (int i = 0; i <= 9; i++) {
            row6.getCell(i).setCellStyle(alignCenterBold_with_boarders_Style);
        }

        // columns to hide
        sheet.setColumnHidden(0, true);
        sheet.setColumnHidden(1, true);
        sheet.setColumnHidden(4, true);
        sheet.setColumnHidden(5, true);
        sheet.setColumnHidden(6, true);

        //add the trading partners starting at col J row 7 AND hide retailer columns from start
        int tpCount = 0;
        for (String tp : tradingPartners) {
            Cell cell = row0.createCell(PARTNER_STARTING_COLUMN + (tpCount * PARTNER_STARTING_COLUMN_OFFSET));
            cell.setCellStyle(alignCenterBold_leftBorderStyle);
            cell.setCellValue(tp);
            row6.createCell(PARTNER_STARTING_COLUMN + (tpCount * PARTNER_STARTING_COLUMN_OFFSET)).setCellValue("M/O/C");
            sheet.setColumnHidden(PARTNER_STARTING_COLUMN + (tpCount * PARTNER_STARTING_COLUMN_OFFSET), true);
            row6.createCell(PARTNER_STARTING_COLUMN + 1 + (tpCount * PARTNER_STARTING_COLUMN_OFFSET)).setCellValue("Qualifiers");
            sheet.setColumnHidden(PARTNER_STARTING_COLUMN + 1 + (tpCount * PARTNER_STARTING_COLUMN_OFFSET), true);
            row6.createCell(PARTNER_STARTING_COLUMN + 2 + (tpCount * PARTNER_STARTING_COLUMN_OFFSET)).setCellValue("Design Notes");
            sheet.setColumnHidden(PARTNER_STARTING_COLUMN + 2 + (tpCount * PARTNER_STARTING_COLUMN_OFFSET), true);
            partnerColumns.put(tp, PARTNER_STARTING_COLUMN + (tpCount * PARTNER_STARTING_COLUMN_OFFSET));
            sheet.addMergedRegion(new CellRangeAddress(0, 0, PARTNER_STARTING_COLUMN + (tpCount * PARTNER_STARTING_COLUMN_OFFSET), PARTNER_STARTING_COLUMN + (tpCount * PARTNER_STARTING_COLUMN_OFFSET) + 2));
            tpCount++;
        }


        //add one last column on row0
        row0.createCell(row0.getLastCellNum() + 2).setCellValue("Unhide for retailer specific information");
        row0.getCell(row0.getLastCellNum() - 1).setCellStyle(boldStyle);

        //start fields at this row
        int currentRowCount = 7;
        for (Node node : linkedNodes) {
            if (node.visible) {
                if (!node.isField() && node.firstChildIsAGroup()) {
                    //create a blank row before each group starts.
                    sheet.createRow(currentRowCount);
                    currentRowCount++;
                }
                StringBuilder hierarchyOffSet = new StringBuilder();
                Row currentRow = sheet.createRow(currentRowCount);
                currentRow.createCell(0).setCellValue(node.xpath);
                currentRow.getCell(0).setCellStyle((!node.isField() && node.firstChildIsAGroup() ? grey25InfillStyle : currentRow.getCell(0).getCellStyle()));
                for (int i = 0; i < node.hierarchyOffSet; i++) {
                    hierarchyOffSet.append(" ");
                }
                currentRow.createCell(1).setCellValue(node.name);
                currentRow.getCell(1).setCellStyle((!node.isField() && node.firstChildIsAGroup() ? boldGrey25Style : boldStyle));
                currentRow.createCell(2).setCellValue(hierarchyOffSet.toString() + node.name);
                currentRow.getCell(2).setCellStyle((!node.isField() && node.firstChildIsAGroup() ? boldGrey25Style : boldStyle));
                currentRow.createCell(3).setCellValue(node.definition);
                currentRow.getCell(3).setCellStyle((!node.isField() && node.firstChildIsAGroup() ? grey25InfillStyle : currentRow.getCell(3).getCellStyle()));
                currentRow.createCell(4).setCellValue(node.dataType);
                currentRow.getCell(4).setCellStyle((!node.isField() && node.firstChildIsAGroup() ? grey25InfillStyle : currentRow.getCell(4).getCellStyle()));
                currentRow.createCell(5).setCellValue(node.minLength);
                currentRow.getCell(5).setCellStyle((!node.isField() && node.firstChildIsAGroup() ? grey25InfillStyle : currentRow.getCell(5).getCellStyle()));
                if (node.isField()) {
                    currentRow.createCell(6).setCellValue(node.maxLength);
                } else {
                    currentRow.createCell(6).setCellValue(node.maxOccurs.equals("unbounded") ? 100000 : Integer.parseInt(node.maxOccurs));
                }
                currentRow.getCell(6).setCellStyle((!node.isField() && node.firstChildIsAGroup() ? grey25InfillStyle : currentRow.getCell(6).getCellStyle()));
                currentRow.createCell(7).setCellValue(node.usage);
                currentRow.getCell(7).setCellStyle((!node.isField() && node.firstChildIsAGroup() ? grey25InfillStyle : currentRow.getCell(7).getCellStyle()));
                Cell qualCell = currentRow.createCell(8);
                currentRow.getCell(8).setCellStyle((!node.isField() && node.firstChildIsAGroup() ? wordWrapGrey25InfillStyle : wordWrapStyle));
                qualCell.setCellValue(node.getQualifiersAndDescriptions().trim());


                //partner specifics start here

                if (!node.mandatoryFor.isEmpty()) {
                    String[] notes = node.mandatoryFor.trim().split(",");
                    for (String mandatoryPartner : notes) {
                        if (partnerColumns.containsKey(mandatoryPartner.trim())) {
                            if (node.isField()) {
                                currentRow.createCell(partnerColumns.get(mandatoryPartner.trim())).setCellValue("M");
                            } else {//group
                                currentRow.createCell(partnerColumns.get(mandatoryPartner.trim())).setCellValue(node.minLength + "/" + node.maxOccurs);
                            }
                        }
                    }
                }
                if (!node.optionalFor.isEmpty()) {
                    String[] notes = node.optionalFor.trim().split(",");
                    for (String mandatoryPartner : notes) {
                        if (partnerColumns.containsKey(mandatoryPartner.trim())) {
                            if (node.isField()) {
                                currentRow.createCell(partnerColumns.get(mandatoryPartner.trim())).setCellValue("O");
                            } else {//group
                                currentRow.createCell(partnerColumns.get(mandatoryPartner.trim())).setCellValue(node.minLength + "/" + node.maxOccurs);
                            }
                        }
                    }
                }
                if (node.mandatoryFor.isEmpty() && node.optionalFor.isEmpty()) {
                    for (String tp : tradingPartners) {
                        if (partnerColumns.containsKey(tp.trim())) {
                            //group
                            currentRow.createCell(partnerColumns.get(tp.trim())).setCellValue(node.minLength + "/" + (node.maxOccurs.equals("unbounded") ? 100000 : Integer.parseInt(node.maxOccurs)));
                        }
                    }
                }

                // take all of the partner specific notes and add them to the correct column
                Pattern regex = Pattern.compile("^(.*?)Notes for(.*?):(.*?)-----(.*)$", Pattern.DOTALL);
                Matcher regexMatcher = regex.matcher(node.designNotes.trim());
                while (regexMatcher.find()) {
                    if (partnerColumns.containsKey(regexMatcher.group(2).trim())) {
                        Cell cell = currentRow.createCell(partnerColumns.get(regexMatcher.group(2).trim()) + PARTNER_NOTES_COLUMN_OFFSET);
                        cell.setCellStyle(wordWrapStyle);
                        cell.setCellValue(regexMatcher.group(3));
                    }
                    String tempNotes = regexMatcher.group(1) + regexMatcher.group(4);
                    regexMatcher = regex.matcher(tempNotes.trim());
                }

                // for each qualifier for this field, add if its mandatory or optional for each partner
                for (Object qualPair : node.qualifiers.entrySet()) {

                    Qualifier qualifier = ((Map.Entry<String, Qualifier>) qualPair).getValue();
                    for (String partnerName : qualifier.mandatoryFor) {
                        if (partnerColumns.containsKey(partnerName)) {
                            if (currentRow.getCell(partnerColumns.get(partnerName) + PARTNER_QUALIFIERS_COLUMN_OFFSET) == null) {
                                Cell cell = currentRow.createCell(partnerColumns.get(partnerName) + PARTNER_QUALIFIERS_COLUMN_OFFSET);
                                cell.setCellStyle(wordWrapStyle);
                                cell.setCellValue(qualifier.value + ":" + qualifier.description + " (mandatory)");
                            } else {
                                String currentCellData = currentRow.getCell(partnerColumns.get(partnerName) + PARTNER_QUALIFIERS_COLUMN_OFFSET).getStringCellValue();
                                Cell cell = currentRow.getCell(partnerColumns.get(partnerName) + PARTNER_QUALIFIERS_COLUMN_OFFSET);
                                cell.setCellStyle(wordWrapStyle);
                                cell.setCellValue(currentCellData + "\r\n" + qualifier.value + ":" + qualifier.description + " (mandatory)");

                            }
                        }
                    }

                    for (String partnerName : qualifier.optionalFor) {

                        if (partnerColumns.containsKey(partnerName)) {
                            if (currentRow.getCell(partnerColumns.get(partnerName) + PARTNER_QUALIFIERS_COLUMN_OFFSET) == null) {
                                Cell cell = currentRow.createCell(partnerColumns.get(partnerName) + PARTNER_QUALIFIERS_COLUMN_OFFSET);
                                cell.setCellStyle(wordWrapStyle);
                                cell.setCellValue(qualifier.value + ":" + qualifier.description);
                            } else {
                                String currentCellData = currentRow.getCell(partnerColumns.get(partnerName) + PARTNER_QUALIFIERS_COLUMN_OFFSET).getStringCellValue();
                                Cell cell = currentRow.getCell(partnerColumns.get(partnerName) + PARTNER_QUALIFIERS_COLUMN_OFFSET);
                                cell.setCellStyle(wordWrapStyle);
                                cell.setCellValue(currentCellData + "\r\n" + qualifier.value + ":" + qualifier.description);

                            }
                        }
                    }
                    //end partner specifics
                }

                //have to set design notes in the matrix after partner columns are set (and removes unwanted notes)
                currentRow.createCell(9).setCellValue(cleanupDesignNotes(node.designNotes));
                currentRow.getCell(9).setCellStyle((!node.isField() && node.firstChildIsAGroup() ? boldGrey25Style : boldStyle));

                // move to next row
                currentRowCount++;
            }
        }

        //some styling for the entire sheet
        applyStyleToRange(sheet, leftBoarderStyle, 1, 10, sheet.getLastRowNum(), row0.getLastCellNum(), 3, true);
        sheet.setDefaultRowHeightInPoints((short) 24);
        sheet.setColumnWidth(2, 7000);
        sheet.setColumnWidth(3, 5000);
        sheet.setColumnWidth(7, 1500);
        sheet.setColumnWidth(8, 5000);
        sheet.setColumnWidth(9, 5000);

        try {
            FileOutputStream fileOut = new FileOutputStream("consolidatedReport.xlsx");
            workbook.write(fileOut);
            fileOut.close();
        } catch (Exception e) {
            e.printStackTrace();
            return 10;
        }

        return 0;
    }

    private void addHowToTab(XSSFWorkbook workbook) {
        XSSFSheet howToSheet = workbook.createSheet("How to Read");


        XSSFRow row0 = howToSheet.createRow(0);
        row0.createCell(0).setCellValue("How to info can go here");
        howToSheet.addMergedRegion(new CellRangeAddress(0, 30, 0, 30));
    }

    private ArrayList<String> getTradingPartners(LinkedList<Node> linkedNodes) {
        ArrayList<String> tradingPartners = new ArrayList<>();
        for (Node node : linkedNodes) {
            if (node.name.equalsIgnoreCase("TradingPartnerId")) {

                String[] tpArray = node.designNotes.replaceAll("Set mandatory for:", "").split(",");
                for (String str : tpArray) {
                    if (!str.trim().isEmpty()) // make sure we are not adding an empty one
                        tradingPartners.add(str.trim());
                }
                // exit after we find a TradingPartnerId
                break;
            }
        }
        //alphabetical order
        Collections.sort(tradingPartners, String.CASE_INSENSITIVE_ORDER);
        return tradingPartners;
    }

    private String cleanupDesignNotes(String notes) {
        String[] regexs = {"^(.*)Notes for.*-----(.*)$", "^(.*)Set mandatory for:.*,(.*)$", "^(.*)is required for:.* ,(.*)$"};
        for (String regularExpression : regexs) {
            Pattern regex = Pattern.compile(regularExpression, Pattern.DOTALL);
            Matcher regexMatcher = regex.matcher(notes.trim());
            while (regexMatcher.find()) {
                notes = regexMatcher.group(1) + regexMatcher.group(2);
                regexMatcher = regex.matcher(notes.trim());
            }
        }
        return notes;
    }

    private void applyStyleToRange(Sheet sheet, CellStyle style, int rowStart, int colStart, int rowEnd, int colEnd, int interval, boolean createCells) {
        for (int r = rowStart; r <= rowEnd; r++) {
            for (int c = colStart; c <= colEnd; c = c + interval) {
                Row row = sheet.getRow(r);
                if (row != null) {
                    Cell cell = row.getCell(c);

                    if (cell != null) {
                        cell.setCellStyle(style);
                    } else {
                        row.createCell(c).setCellStyle(style);
                    }
                }
            }
        }
    }
}
