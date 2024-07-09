package com.example.excelutilizer.v1.service;

import com.example.excelutilizer.v1.annotation.ExcelColumn;
import com.fasterxml.jackson.databind.JsonNode;
import java.lang.reflect.Field;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

@Service
public class JsonToExcelServiceImpl implements JsonToExcelService {

    @Override
    public <T> Workbook convertJsonToExcel(JsonNode jsonNode, Class<T> clazz) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sheet1");

        // DTO 필드명을 첫 번째 행에 추가 - Header 세팅
        Row headerRow = sheet.createRow(0);
        Field[] dtoFields = clazz.getDeclaredFields();
        int cellIndex = 0;
        for (Field field : dtoFields) {
            if (field.isAnnotationPresent(ExcelColumn.class)) {
                ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);
                Cell currentCell = headerRow.getCell(cellIndex++);
                currentCell.setCellValue(annotation.name());
            }
        }

        // JSON 데이터를 row 에 추가
        int rowIndex = 1;
        for (JsonNode rowNode : jsonNode) {
            Row row = sheet.createRow(rowIndex++);
            cellIndex = 0;
            for (Field field : dtoFields) {
                if (field.isAnnotationPresent(ExcelColumn.class)) {
                    ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);
                    String columnName = annotation.name();
                    JsonNode fieldValue = rowNode.get(columnName);
                    Cell cell = row.createCell(cellIndex++);

                    if (fieldValue != null) {
                        cell.setCellValue(fieldValue.asText());
                    } else {
                        cell.setCellValue("");
                    }
                }
            }
        }

        return workbook;
    }
}
