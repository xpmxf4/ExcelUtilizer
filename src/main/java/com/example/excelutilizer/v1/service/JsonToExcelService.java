package com.example.excelutilizer.v1.service;

import com.fasterxml.jackson.databind.JsonNode;
import org.apache.poi.ss.usermodel.Workbook;

public interface JsonToExcelService {

    <T> Workbook convertJsonToExcel(JsonNode jsonNode, Class<T> clazz);
}
