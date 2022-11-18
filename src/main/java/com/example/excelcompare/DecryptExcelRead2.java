package com.example.excelcompare;
import java.io.File;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class DecryptExcelRead2 {
    public static void main(String[] args) throws IOException {

        // File object with specific path to read
        File xlsxFile = new File("파일경로");

        try {

            // Creating workbook using WorkbookFactory with password
            // It works for both excel format xls and xlsx
            Workbook workbook = WorkbookFactory.create(xlsxFile, "패스워드");

            // Reading the first sheet of the excel file
            Sheet sheet = workbook.getSheetAt(0);

            Iterator<Row> iterator = sheet.iterator();

            // Iterating all the rows
            while (iterator.hasNext()) {
                Row nextRow = iterator.next();
                Iterator<Cell> cellIterator = nextRow.cellIterator();

                // Iterating all the columns in a row
                while (cellIterator.hasNext()) {

                    Cell cell = cellIterator.next();

                    switch (cell.getCellType()) {
                        case STRING:
                            System.out.print(cell.getStringCellValue());
                            break;
                        case BOOLEAN:
                            System.out.print(cell.getBooleanCellValue());
                            break;
                        case NUMERIC:
                            System.out.print(cell.getNumericCellValue());
                            break;
                        default:
                            break;
                    }
                    System.out.print(" | ");
                }
                System.out.println();
            }

            workbook.close();
        } catch (EncryptedDocumentException | IOException ex) {
            throw new RuntimeException("Unable to process encrypted document", ex);
        }
    }
}
