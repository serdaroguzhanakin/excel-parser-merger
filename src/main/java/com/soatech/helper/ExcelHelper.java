package com.soatech.helper;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Random;

public class ExcelHelper {
    private static List<String> restrictedWords = Arrays.asList("gamb", "bet", "casino", "cash");
    private static String MAIL_SUFFIX = "@hotmail.com";

    public static boolean isRowListEmpty(ArrayList<Row> rowList) {
        return (rowList == null || rowList.get(0) == null || rowList.get(0).getCell(0) == null);
    }

    public static Workbook createWorkBook() {
        return new XSSFWorkbook();
    }

    public static Sheet createSheet(Workbook workbook, String name) {
        return workbook.createSheet(name);
    }

    public static void createHeader(Sheet sheet, ArrayList<Cell> headers) {
        Row headerRow = sheet.createRow(0);

        for (int i = 0; i < headers.size(); i++) {
            Cell cell = headerRow.createCell(i);
            if (headers.get(i) == null) {
                cell.setCellValue("#UNKNOWN_VALUE#");
                continue;
            }
            switch (headers.get(i).getCellTypeEnum()) {
                case STRING:
                    cell.setCellValue(headers.get(i).getStringCellValue());
                    break;
                case NUMERIC:
                    cell.setCellValue(headers.get(i).getNumericCellValue());
                    break;
                case _NONE:
                case BLANK:
                case ERROR:
                case FORMULA:
                    cell.setCellValue("#UNKNOWN_VALUE#");
                    break;
            }
        }
    }

    public static void createBody(Sheet sheet, ArrayList<Row> bodyRows, boolean shouldArrangeRestrictedWords) {
        int lineNumber = 1;
        for (Row bodyRow : bodyRows) {
            Row newRow = sheet.createRow(lineNumber++);

            for (int j = 0; j < bodyRow.getPhysicalNumberOfCells(); j++) {
                Cell cell = bodyRow.getCell(j);
                if (cell == null) {
                    newRow.createCell(j).setCellValue("");
                    continue;
                }
                switch (cell.getCellTypeEnum()) {
                    case STRING:
                        newRow.createCell(j)
                                .setCellValue(cell.getStringCellValue());
                        if (j == 7) {//BeneficiaryEmailAddress 7. element - BeneficiaryName 0. element
                            String arrangedEmail = arrangeRestrictedWords(cell.getStringCellValue(),
                                    bodyRow.getCell(0).getStringCellValue());
                            newRow.createCell(j)
                                    .setCellValue(arrangedEmail);
                        } else {
                            newRow.createCell(j)
                                    .setCellValue(cell.getStringCellValue());
                        }
                        break;
                    case NUMERIC:
                        newRow.createCell(j)
                                .setCellValue(cell.getNumericCellValue());
                        break;
                    case BOOLEAN:
                        newRow.createCell(j)
                                .setCellValue(cell.getBooleanCellValue());
                        break;
                    case _NONE:
                    case BLANK:
                    case ERROR:
                    case FORMULA:
                        newRow.createCell(j)
                                .setCellValue("#UNKNOWN_VALUE#");
                        break;
                }
            }
        }
    }

    public static void resizeAllColumns(Sheet sheet, ArrayList<Cell> headers) {
        for (int i = 0; i < headers.size(); i++) {
            sheet.autoSizeColumn(i);
        }
    }

    public static String arrangeRestrictedWords(String email, String beneficiaryName) {
        String arrangedEmail = email;
        String[] beneficiaryNames = beneficiaryName.split(" ");
        Random random = new Random(1000);

        if (restrictedWords.stream()
                .anyMatch(x -> email.trim().toLowerCase().contains(x))) {

            if (beneficiaryNames.length > 1) {
                arrangedEmail = (
                        beneficiaryNames[0].trim() + beneficiaryNames[1].trim()
                ).toLowerCase() + MAIL_SUFFIX;
            } else {
                arrangedEmail = (
                        beneficiaryNames[0].trim() + Math.abs(random.nextInt(100))
                ).toLowerCase() + MAIL_SUFFIX;
            }
        }
        return arrangedEmail;
    }
}
