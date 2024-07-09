package com.example.excelutilizer.v0;

import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelExample {

    public static void main(String[] args) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Example Sheet");

        for (int i = 0; i < 10; i++) {
            Row row = sheet.createRow(i);
            Cell cell = row.createCell(i); // 각 행의 첫 번째 셀에 값을 설정
            cell.setCellValue("Hello World" + i);
            System.out.println("cell = " + cell);
        }

        try (FileOutputStream outputStream = new FileOutputStream("Example.xlsx")) {
            workbook.write(outputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

}
