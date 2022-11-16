package com.example.excelcompare;

import java.io.File;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;


public class DecryptExcelRead {
    public static void main(String[] args) throws IOException {
        String excelFilePath = "파일경로";
        boolean readOnly = true;

        NPOIFSFileSystem fileSystem = new NPOIFSFileSystem(new File(excelFilePath), readOnly);

        Biff8EncryptionKey.setCurrentUserPassword("****");

        Workbook workbook = new HSSFWorkbook(fileSystem);

        Sheet firstSheet = workbook.getSheetAt(0);
        Iterator<Row> iterator = firstSheet.iterator();

        while (iterator.hasNext()) {
            Row nextRow = iterator.next();
            Iterator<Cell> cellIterator = nextRow.cellIterator();

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();

                switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_STRING:
                        System.out.print(cell.getStringCellValue());
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        System.out.print(cell.getNumericCellValue());
                        break;
                }
                System.out.print("\t");
            }
            System.out.println();
        }

        workbook.close();
        fileSystem.close();
    }
}